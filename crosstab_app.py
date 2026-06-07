from __future__ import annotations

import argparse
import csv
import datetime as dt
import io
import json
import socket
import sys
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

try:
    import openpyxl
except ImportError as exc:  # pragma: no cover - shown to users at launch
    raise SystemExit(
        "openpyxl 모듈을 찾을 수 없습니다. run_crosstab.bat으로 실행하거나 Python 환경에 openpyxl을 설치해 주세요."
    ) from exc


APP_TITLE = "조건부확률 이중교차표"
MISSING_LABEL = "응답 없음"


def normalize_cell(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, dt.date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def is_question_header(header: str) -> bool:
    text = header.strip().lower()
    if not text:
        return False
    return text not in {"타임스탬프", "timestamp", "time stamp"}


def load_survey_data(workbook_path: Path) -> dict:
    if not workbook_path.exists():
        raise FileNotFoundError(f"{workbook_path} 파일을 찾을 수 없습니다.")

    workbook = openpyxl.load_workbook(workbook_path, data_only=True, read_only=True)
    worksheet = workbook.worksheets[0]
    iterator = worksheet.iter_rows(values_only=True)

    try:
        headers = [normalize_cell(value) for value in next(iterator)]
    except StopIteration as exc:
        raise ValueError("엑셀 파일에 데이터가 없습니다.") from exc

    question_columns = [
        (index, header)
        for index, header in enumerate(headers)
        if is_question_header(header)
    ]
    if len(question_columns) < 2:
        raise ValueError("교차표를 만들 질문 열이 2개 이상 필요합니다.")

    rows: list[list[str]] = []
    for row in iterator:
        normalized = [normalize_cell(value) for value in row]
        if not any(normalized):
            continue
        rows.append(
            [
                normalized[index] if index < len(normalized) else ""
                for index, _header in question_columns
            ]
        )

    questions = [
        {
            "id": position,
            "column": index + 1,
            "title": header,
            "label": f"Q{position + 1}. {header}",
        }
        for position, (index, header) in enumerate(question_columns)
    ]

    return {
        "appTitle": APP_TITLE,
        "workbook": str(workbook_path.name),
        "sheet": worksheet.title,
        "respondents": len(rows),
        "missingLabel": MISSING_LABEL,
        "questions": questions,
        "rows": rows,
    }


def build_check_crosstab(data: dict, first: int = 0, second: int = 1) -> list[list[object]]:
    row_labels: list[str] = []
    col_labels: list[str] = []
    row_seen: set[str] = set()
    col_seen: set[str] = set()
    counts: dict[tuple[str, str], int] = {}

    for row in data["rows"]:
        row_value = row[first] or MISSING_LABEL
        col_value = row[second] or MISSING_LABEL
        if row_value not in row_seen:
            row_seen.add(row_value)
            row_labels.append(row_value)
        if col_value not in col_seen:
            col_seen.add(col_value)
            col_labels.append(col_value)
        counts[(row_value, col_value)] = counts.get((row_value, col_value), 0) + 1

    table: list[list[object]] = [["", *col_labels, "합계"]]
    for row_label in row_labels:
        row_counts = [counts.get((row_label, col_label), 0) for col_label in col_labels]
        table.append([row_label, *row_counts, sum(row_counts)])
    total_row = ["합계"]
    for col_label in col_labels:
        total_row.append(sum(counts.get((row_label, col_label), 0) for row_label in row_labels))
    total_row.append(len(data["rows"]))
    table.append(total_row)
    return table


HTML = r"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>조건부확률 이중교차표</title>
  <style>
    :root {
      color-scheme: light;
      --bg: #f6f7f4;
      --panel: #ffffff;
      --ink: #1f2522;
      --muted: #68736d;
      --line: #dce3dc;
      --accent: #117865;
      --accent-strong: #0b5d4f;
      --accent-soft: #dff3ed;
      --warn: #9c4a1a;
      --shadow: 0 14px 34px rgba(31, 37, 34, 0.08);
    }

    * { box-sizing: border-box; }

    body {
      margin: 0;
      min-height: 100vh;
      background: var(--bg);
      color: var(--ink);
      font-family: "Malgun Gothic", "Segoe UI", system-ui, sans-serif;
      letter-spacing: 0;
    }

    .app {
      min-height: 100vh;
      display: grid;
      grid-template-rows: auto auto 1fr;
    }

    header {
      padding: 24px clamp(18px, 4vw, 42px) 18px;
      border-bottom: 1px solid var(--line);
      background: #fbfcfa;
    }

    h1 {
      margin: 0;
      font-size: clamp(24px, 3vw, 36px);
      line-height: 1.2;
      font-weight: 800;
    }

    .meta {
      display: flex;
      flex-wrap: wrap;
      gap: 8px 18px;
      margin-top: 10px;
      color: var(--muted);
      font-size: 14px;
    }

    .toolbar {
      display: grid;
      grid-template-columns: minmax(0, 1fr) minmax(0, 1fr) auto auto auto;
      gap: 12px;
      padding: 18px clamp(18px, 4vw, 42px);
      align-items: end;
      background: #ffffff;
      border-bottom: 1px solid var(--line);
    }

    label {
      display: grid;
      gap: 7px;
      color: var(--muted);
      font-size: 13px;
      font-weight: 700;
    }

    select, button {
      height: 42px;
      border-radius: 8px;
      border: 1px solid var(--line);
      background: var(--panel);
      color: var(--ink);
      font: inherit;
    }

    select {
      width: 100%;
      padding: 0 34px 0 12px;
    }

    button {
      padding: 0 14px;
      cursor: pointer;
      font-weight: 800;
      white-space: nowrap;
    }

    button.primary {
      color: #ffffff;
      border-color: var(--accent);
      background: var(--accent);
    }

    button:hover {
      border-color: var(--accent-strong);
    }

    main {
      padding: 22px clamp(18px, 4vw, 42px) 34px;
      display: grid;
      gap: 14px;
    }

    .status {
      min-height: 24px;
      color: var(--muted);
      font-size: 14px;
      display: flex;
      flex-wrap: wrap;
      gap: 8px 16px;
      align-items: center;
    }

    .warning {
      color: var(--warn);
      font-weight: 800;
    }

    .table-wrap {
      overflow: auto;
      border: 1px solid var(--line);
      border-radius: 8px;
      background: var(--panel);
      box-shadow: var(--shadow);
      max-height: calc(100vh - 230px);
    }

    table {
      width: 100%;
      min-width: 720px;
      border-collapse: separate;
      border-spacing: 0;
      font-size: 14px;
    }

    th, td {
      border-right: 1px solid var(--line);
      border-bottom: 1px solid var(--line);
      padding: 10px 12px;
      text-align: center;
      vertical-align: middle;
    }

    th {
      position: sticky;
      top: 0;
      z-index: 2;
      background: #eef4ef;
      font-weight: 800;
    }

    th:first-child,
    td:first-child {
      position: sticky;
      left: 0;
      z-index: 1;
      background: #fbfcfa;
      text-align: left;
      min-width: 210px;
      max-width: 320px;
      overflow-wrap: anywhere;
      font-weight: 800;
    }

    th:first-child {
      z-index: 3;
      background: #e7efe9;
    }

    td.count {
      font-variant-numeric: tabular-nums;
      font-weight: 800;
      min-width: 74px;
    }

    .total {
      background: #f4f7f1;
      font-weight: 900;
    }

    .empty {
      padding: 36px;
      color: var(--muted);
      text-align: center;
    }

    @media (max-width: 920px) {
      .toolbar {
        grid-template-columns: 1fr;
      }

      button {
        width: 100%;
      }

      .table-wrap {
        max-height: none;
      }
    }
  </style>
</head>
<body>
  <div class="app">
    <header>
      <h1>조건부확률 이중교차표</h1>
      <div class="meta">
        <span id="sourceName">불러오는 중</span>
        <span id="sheetName"></span>
        <span id="respondentCount"></span>
      </div>
    </header>

    <section class="toolbar" aria-label="질문 선택">
      <label>
        행 질문
        <select id="rowQuestion"></select>
      </label>
      <label>
        열 질문
        <select id="colQuestion"></select>
      </label>
      <button class="primary" id="swapBtn" type="button">행/열 바꾸기</button>
      <button id="copyBtn" type="button">복사</button>
      <button id="csvBtn" type="button">CSV 저장</button>
    </section>

    <main>
      <div class="status" id="status"></div>
      <div class="table-wrap" id="tableWrap">
        <div class="empty">데이터를 불러오는 중입니다.</div>
      </div>
    </main>
  </div>

  <script>
    const state = {
      data: null,
      table: [],
      rowLabels: [],
      colLabels: [],
    };

    const rowSelect = document.getElementById("rowQuestion");
    const colSelect = document.getElementById("colQuestion");
    const statusEl = document.getElementById("status");
    const tableWrap = document.getElementById("tableWrap");

    function valueLabel(value) {
      const trimmed = String(value ?? "").trim();
      return trimmed || state.data.missingLabel || "응답 없음";
    }

    function escapeHtml(value) {
      return String(value).replace(/[&<>"']/g, (char) => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      }[char]));
    }

    function fillSelects() {
      const options = state.data.questions.map((question) => {
        const text = escapeHtml(question.label);
        return `<option value="${question.id}">${text}</option>`;
      }).join("");
      rowSelect.innerHTML = options;
      colSelect.innerHTML = options;
      rowSelect.value = "0";
      colSelect.value = state.data.questions.length > 1 ? "1" : "0";
    }

    function buildTable() {
      const rowIndex = Number(rowSelect.value);
      const colIndex = Number(colSelect.value);
      const rowLabels = [];
      const colLabels = [];
      const rowSeen = new Set();
      const colSeen = new Set();
      const counts = new Map();

      for (const row of state.data.rows) {
        const rowValue = valueLabel(row[rowIndex]);
        const colValue = valueLabel(row[colIndex]);
        if (!rowSeen.has(rowValue)) {
          rowSeen.add(rowValue);
          rowLabels.push(rowValue);
        }
        if (!colSeen.has(colValue)) {
          colSeen.add(colValue);
          colLabels.push(colValue);
        }
        const key = `${rowValue}\u0000${colValue}`;
        counts.set(key, (counts.get(key) || 0) + 1);
      }

      const body = [];
      let maxCount = 0;
      for (const rowLabel of rowLabels) {
        const values = colLabels.map((colLabel) => {
          const count = counts.get(`${rowLabel}\u0000${colLabel}`) || 0;
          maxCount = Math.max(maxCount, count);
          return count;
        });
        body.push({ rowLabel, values, total: values.reduce((sum, count) => sum + count, 0) });
      }

      const colTotals = colLabels.map((colLabel) => {
        return rowLabels.reduce((sum, rowLabel) => sum + (counts.get(`${rowLabel}\u0000${colLabel}`) || 0), 0);
      });

      state.rowLabels = rowLabels;
      state.colLabels = colLabels;
      state.table = [
        ["", ...colLabels, "합계"],
        ...body.map((item) => [item.rowLabel, ...item.values, item.total]),
        ["합계", ...colTotals, state.data.rows.length],
      ];

      renderTable(body, colLabels, colTotals, maxCount);
      renderStatus(rowIndex, colIndex);
    }

    function heatStyle(count, maxCount) {
      if (!maxCount || count === 0) return "";
      const intensity = Math.max(0.14, count / maxCount);
      return `background: rgba(17, 120, 101, ${0.10 + intensity * 0.42});`;
    }

    function renderTable(body, colLabels, colTotals, maxCount) {
      if (!state.data.rows.length) {
        tableWrap.innerHTML = '<div class="empty">응답 데이터가 없습니다.</div>';
        return;
      }

      let html = "<table><thead><tr><th>구분</th>";
      for (const label of colLabels) {
        html += `<th>${escapeHtml(label)}</th>`;
      }
      html += '<th class="total">합계</th></tr></thead><tbody>';

      for (const item of body) {
        html += `<tr><td>${escapeHtml(item.rowLabel)}</td>`;
        for (const count of item.values) {
          html += `<td class="count" style="${heatStyle(count, maxCount)}">${count}</td>`;
        }
        html += `<td class="count total">${item.total}</td></tr>`;
      }

      html += '<tr><td class="total">합계</td>';
      for (const total of colTotals) {
        html += `<td class="count total">${total}</td>`;
      }
      html += `<td class="count total">${state.data.rows.length}</td></tr>`;
      html += "</tbody></table>";
      tableWrap.innerHTML = html;
    }

    function renderStatus(rowIndex, colIndex) {
      const rowTitle = state.data.questions[rowIndex]?.title || "";
      const colTitle = state.data.questions[colIndex]?.title || "";
      const warning = rowIndex === colIndex
        ? '<span class="warning">같은 질문이 선택되었습니다.</span>'
        : "";
      statusEl.innerHTML = [
        `<span>행 선택지 ${state.rowLabels.length}개</span>`,
        `<span>열 선택지 ${state.colLabels.length}개</span>`,
        `<span>전체 ${state.data.rows.length}명</span>`,
        warning,
      ].filter(Boolean).join("");
      document.title = `${rowTitle} × ${colTitle}`;
    }

    function tableToDelimited(delimiter) {
      return state.table.map((row) => row.map((cell) => {
        const text = String(cell);
        if (text.includes(delimiter) || text.includes("\n") || text.includes('"')) {
          return `"${text.replace(/"/g, '""')}"`;
        }
        return text;
      }).join(delimiter)).join("\n");
    }

    async function copyTable() {
      const text = tableToDelimited("\t");
      await navigator.clipboard.writeText(text);
      statusEl.insertAdjacentHTML("beforeend", "<span>표를 클립보드에 복사했습니다.</span>");
    }

    function downloadCsv() {
      const csv = "\ufeff" + tableToDelimited(",");
      const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "conditional_probability_crosstab.csv";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    }

    async function loadData() {
      const response = await fetch("/api/data", { cache: "no-store" });
      if (!response.ok) {
        const text = await response.text();
        throw new Error(text || "데이터를 불러오지 못했습니다.");
      }
      state.data = await response.json();
      document.getElementById("sourceName").textContent = state.data.workbook;
      document.getElementById("sheetName").textContent = `시트: ${state.data.sheet}`;
      document.getElementById("respondentCount").textContent = `응답: ${state.data.respondents}명`;
      fillSelects();
      buildTable();
    }

    rowSelect.addEventListener("change", buildTable);
    colSelect.addEventListener("change", buildTable);
    document.getElementById("swapBtn").addEventListener("click", () => {
      const rowValue = rowSelect.value;
      rowSelect.value = colSelect.value;
      colSelect.value = rowValue;
      buildTable();
    });
    document.getElementById("copyBtn").addEventListener("click", copyTable);
    document.getElementById("csvBtn").addEventListener("click", downloadCsv);

    loadData().catch((error) => {
      tableWrap.innerHTML = `<div class="empty">${escapeHtml(error.message)}</div>`;
      statusEl.innerHTML = '<span class="warning">오류가 발생했습니다.</span>';
    });
  </script>
</body>
</html>
"""


class CrosstabHandler(BaseHTTPRequestHandler):
    workbook_path: Path

    def log_message(self, format: str, *args: object) -> None:
        return

    def send_text(self, text: str, content_type: str, status: int = 200) -> None:
        encoded = text.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", f"{content_type}; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(encoded)

    def do_GET(self) -> None:
        path = unquote(urlparse(self.path).path)
        if path == "/":
            self.send_text(HTML, "text/html")
            return

        if path == "/api/data":
            try:
                payload = json.dumps(load_survey_data(self.workbook_path), ensure_ascii=False)
                self.send_text(payload, "application/json")
            except Exception as exc:  # pragma: no cover - browser-facing error
                self.send_text(str(exc), "text/plain", status=500)
            return

        self.send_text("Not found", "text/plain", status=404)


def find_available_port(preferred: int) -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        try:
            sock.bind(("127.0.0.1", preferred))
            return preferred
        except OSError:
            sock.bind(("127.0.0.1", 0))
            return int(sock.getsockname()[1])


def run_server(workbook_path: Path, port: int, open_browser: bool) -> None:
    CrosstabHandler.workbook_path = workbook_path.resolve()
    server = ThreadingHTTPServer(("127.0.0.1", find_available_port(port)), CrosstabHandler)
    url = f"http://127.0.0.1:{server.server_port}/"

    print(f"{APP_TITLE} 실행 중")
    print(f"엑셀 파일: {workbook_path.resolve()}")
    print(f"주소: {url}")
    print("종료하려면 이 창에서 Ctrl+C를 누르세요.")

    if open_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n종료합니다.")
    finally:
        server.server_close()


def write_check_summary(workbook_path: Path) -> None:
    data = load_survey_data(workbook_path)
    print(f"workbook={data['workbook']}")
    print(f"sheet={data['sheet']}")
    print(f"respondents={data['respondents']}")
    print(f"questions={len(data['questions'])}")
    print("sample_questions:")
    for question in data["questions"][:5]:
        print(f"- {question['label']}")

    table = build_check_crosstab(data)
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerows(table)
    print("sample_crosstab:")
    print(output.getvalue().strip())


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=APP_TITLE)
    parser.add_argument(
        "workbook",
        nargs="?",
        default="data.xlsx",
        help="읽을 엑셀 파일 경로입니다. 기본값: data.xlsx",
    )
    parser.add_argument("--port", type=int, default=8765, help="웹 서버 포트입니다. 기본값: 8765")
    parser.add_argument("--no-browser", action="store_true", help="브라우저 자동 실행을 끕니다.")
    parser.add_argument("--check", action="store_true", help="엑셀 인식과 기본 교차표 계산만 확인합니다.")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    workbook_path = Path(args.workbook)
    if not workbook_path.is_absolute():
        workbook_path = Path.cwd() / workbook_path

    if args.check:
        write_check_summary(workbook_path)
        return 0

    run_server(workbook_path, args.port, not args.no_browser)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
