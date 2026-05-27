import os
from pathlib import Path
import fitz
import re
import copy
import openpyxl
from typing import Callable
from openpyxl.styles import Alignment, Border, Side


def _check_cancel(cancel_check: Callable[[], None] | None) -> None:
  if cancel_check:
    cancel_check()

def extract_text_from_pdf(pdf_path):
  if not os.path.exists(pdf_path):
    raise FileNotFoundError(f"PDF 파일을 찾을 수 없습니다: {pdf_path}")
  
  print(f"\n*** 사용자인터페이스설계서에서 텍스트 추출 중: {pdf_path}")
  text_content = ""
  with fitz.open(pdf_path) as doc:
    for page_num in range(len(doc)):
      page = doc.load_page(page_num)
      text_content += page.get_text()

  return text_content.strip()

def extract_req_mapping_from_pdf(pdf_path):
  """사용자인터페이스설계서 PDF에서 {화면ID: 요구사항ID} 매핑을 만든다."""
  text = extract_text_from_pdf(pdf_path)
  mapping = {}

  normalized = re.sub(r'[ \t]+', ' ', text)

  pair_pattern = re.compile(
    r'요구\s*사항\s*ID\s*'
    r'(SFR-[A-Z0-9]+(?:-[A-Z0-9]+)*)'
    r'[\s\S]{0,500}?'
    r'화면\s*ID\s*'
    r'(UI-[A-Z0-9]+(?:-[A-Z0-9]+)*)',
    re.IGNORECASE,
  )

  for match in pair_pattern.finditer(normalized):
    req_id = match.group(1).strip()
    screen_id = match.group(2).strip()
    mapping[screen_id] = req_id

  return mapping

def extract_unit_test_from_excel(xlsx_path):
  if not os.path.exists(xlsx_path):
    raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {xlsx_path}")
  
  print(f"단위시험 케이스에서 데이터 추출 중: {xlsx_path}")
  wb = openpyxl.load_workbook(xlsx_path, data_only=True)

  unit_test_data = []

  for sheet in wb.worksheets:
    header_row_idx = None
    headers = []
    common_info = {}

    # 1. 상단 정보 추출 및 헤더(항목명)가 있는 행 찾기
    first_row_val = sheet.cell(row=1, column=1).value
    if first_row_val:
      common_info["타이틀"] = str(first_row_val).strip()

    for row_idx, row in enumerate(sheet.iter_rows(values_only=True), 1):
      if not any(row):
        continue

      # 셀 값들의 띄어쓰기와 줄바꿈을 제거하여 검색용 리스트 생성
      search_row = [str(cell).replace(" ", "").replace("\n", "") for cell in row if cell is not None]

      # '순서'(또는 '순번')과 '예상'이라는 단어가 같은 행에 모두 존재하면 헤더
      has_seq = any('순번' in val or '순서' in val for val in search_row)
      has_result = any('예상' in val or '결과' in val for val in search_row)

      if has_seq and has_result:
        header_row_idx = row_idx
        headers = [str(cell).strip().replace(" ", "_") if cell else f"Col_{i}" for i, cell in enumerate(row)]
        break
      else:
        # 헤더가 아니라면 상단의 메타 정보 행이므로 '사전조건', '화면ID' 등을 찾아 common_info에 저장
        for c_idx, cell in enumerate(row):
          if isinstance(cell, str):
            clean_label = cell.replace(" ", "")
            key_map = {
              "단위시험ID": "단위시험_ID",
              "단위시험명": "단위시험_명",
              "사전조건" : "사전조건",
              "화면ID": "화면_ID"
            }
            if clean_label in key_map:
              # 바로 오른쪽 칸에 있는 값을 가져옴
              val = row[c_idx + 1] if c_idx + 1 < len(row) and row[c_idx + 1] is not None else ""
              common_info[key_map[clean_label]] = str(val).strip()

    if not header_row_idx:
      continue

    # 2. 헤더 다음 행부터 실제 데이터 읽기
    for row in sheet.iter_rows(min_row=header_row_idx + 1, values_only=True):
      if not any(row):
        continue

      # 헤더와 값을 1:1로 매핑하여 딕셔너리 생성
      row_data = {}
      for i, cell in enumerate(row):
        if i < len(headers):
          row_data[headers[i]] = str(cell).strip() if cell is not None else ""

      # 3. 추출해 둔 상단 메타 정보를 개별 스텝 데이터에 합침
      row_data.update(common_info)

      unit_test_data.append(row_data)

  return unit_test_data

def build_scenario_id(screen_id):
  screen_id = str(screen_id or "").strip()
  if screen_id.startswith("UI-"):
    return "AT-" + screen_id[3:]
  if screen_id == "UI":
    return "AT"
  return screen_id

def extract_scenario_case_name(title, unit_test_name):
  if title and " - " in title:
    return title.rsplit(" - ", 1)[-1].strip()
  if unit_test_name:
    return unit_test_name.strip()
  return ""

def normalize_label(value):
  return re.sub(r"\s+", "", str(value or ""))

def extract_cover_author_from_workbook(workbook_path):
  try:
    wb = openpyxl.load_workbook(workbook_path, data_only=True)
  except Exception:
    return ""

  try:
    cover_sheets = [ws for ws in wb.worksheets if "표지" in ws.title]
    for ws in cover_sheets:
      for row in ws.iter_rows():
        cells = list(row)
        for index, cell in enumerate(cells):
          if normalize_label(cell.value) != "작성자":
            continue
          for next_cell in cells[index + 1:]:
            if next_cell.value not in (None, ""):
              return str(next_cell.value).strip()

    for ws in wb.worksheets:
      for row in ws.iter_rows():
        cells = list(row)
        for cell in cells:
          if normalize_label(cell.value) != "작성자":
            continue
          for next_row in range(cell.row + 1, min(cell.row + 6, ws.max_row) + 1):
            value = ws.cell(row=next_row, column=cell.column).value
            if value not in (None, ""):
              return str(value).strip()
  finally:
    wb.close()

  return ""

def build_test_scenarios_from_unit_tests(unit_test_data, author=""):
  scenarios = []
  sequence_by_scenario = {}
  scenario_name_by_id = {}

  if unit_test_data:
    print(f"[통합시험 시나리오] 단위시험 데이터 {len(unit_test_data)}건 변환 중...")

  for row in unit_test_data:
    screen_id = row.get("화면_ID", "")
    scenario_id = build_scenario_id(screen_id)
    sequence_by_scenario[scenario_id] = sequence_by_scenario.get(scenario_id, 0) + 1

    scenario_name = scenario_name_by_id.get(scenario_id)
    if scenario_name is None:
      scenario_name = row.get("단위시험_명", "") or extract_scenario_case_name(row.get("타이틀", ""), "")
      scenario_name_by_id[scenario_id] = scenario_name

    scenarios.append({
      "시스템": "",
      "작성자": author,
      "테스트_기간": "",
      "시나리오ID": scenario_id,
      "시나리오명": scenario_name,
      "요구사항_ID": row.get("요구사항_ID", ""),
      "케이스명": extract_scenario_case_name(row.get("타이틀", ""), row.get("단위시험_명", "")),
      "순번": sequence_by_scenario[scenario_id],
      "업무처리내용": "",
      "시험항목": row.get("테스트_케이스", ""),
      "사전조건": row.get("사전조건", ""),
      "입력자료": row.get("테스트_데이터", ""),
      "예상결과": row.get("예상_결과", ""),
      "화면ID": row.get("화면_ID", "")
    })

  return scenarios
  
def copy_worksheet_template(source_ws, target_ws):
  """외부 엑셀 시트의 너비, 높이, 병합, 셀 서식을 완벽하게 복사하는 도우미 함수"""
  # 1. 열 너비 복사
  for key, dim in source_ws.column_dimensions.items():
    target_ws.column_dimensions[key].min = dim.min
    target_ws.column_dimensions[key].max = dim.max
    target_ws.column_dimensions[key].width = dim.width
    target_ws.column_dimensions[key].hidden = dim.hidden
  
  # 2. 행 높이 복사
  for key, dim in source_ws.row_dimensions.items():
    target_ws.row_dimensions[key].height = dim.height
    target_ws.row_dimensions[key].hidden = dim.hidden
  
  # 3. 병합된 셀 복사
  for m_range in source_ws.merged_cells.ranges:
    target_ws.merge_cells(str(m_range))
  
  # 4. 셀 데이터 및 스타일 복사
  for row in source_ws.iter_rows():
    for cell in row:
      new_cell = target_ws.cell(row=cell.row, column=cell.column, value=cell.value)
      if cell.has_style:
        new_cell.font = copy.copy(cell.font)
        new_cell.border = copy.copy(cell.border)
        new_cell.fill = copy.copy(cell.fill)
        new_cell.number_format = copy.copy(cell.number_format)
        new_cell.protection = copy.copy(cell.protection)
        new_cell.alignment = copy.copy(cell.alignment)

def save_test_scenarios_to_excel(test_scenarios, base_workbook_path, output_dir: Path, scenario_sheet_form_path, base_filename="generated_TS"):
  if not test_scenarios:
    print("저장할 데이터가 없습니다.")
    return
  
  base_workbook_path = Path(base_workbook_path)
  output_dir = Path(output_dir)
  scenario_sheet_form_path = Path(scenario_sheet_form_path)
  
  if not base_workbook_path.exists():
    print(f"업로드된 기존 시나리오 파일을 찾을 수 없습니다: {base_workbook_path}")
    return
  
  if not scenario_sheet_form_path.exists():
    print(f"서버에 표준 양식 파일({scenario_sheet_form_path})이 없습니다.")
    return
  
  output_dir.mkdir(parents=True, exist_ok=True)
  
  # 파일명 자동 증가
  counter = 1
  while True:
    output_filename = f"{base_filename}_{counter}.xlsx"
    full_path = output_dir / output_filename
    if not full_path.exists():
      break
    counter += 1

  print(f"기존 파일 로드 중: {base_workbook_path}")
  wb = openpyxl.load_workbook(base_workbook_path)

  print(f"외부 표준 양식 로드 중: {scenario_sheet_form_path}")
  form_wb = openpyxl.load_workbook(scenario_sheet_form_path)
  source_ws = form_wb.active

  # 1. '개정이력'과 '작성방법' 인덱스 찾기
  sheet_names = wb.sheetnames
  start_idx = -1
  end_idx = -1

  for i, name in enumerate(sheet_names):
    if "개정이력" in name:
      start_idx = i
    elif "작성방법" in name:
      end_idx = i

  # 3. 기존 시나리오 시트 삭제 (개정이력과 작성방법 사이)
  if start_idx != -1 and end_idx != -1 and start_idx < end_idx:
      sheets_to_remove = sheet_names[start_idx + 1 : end_idx]
      for sheet_name in sheets_to_remove:
        del wb[sheet_name]
      insert_idx = start_idx + 1
  else:
      insert_idx = 1

  border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
  center_align = Alignment(horizontal="center", vertical="center")
  left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)

  # 시나리오ID 기준으로 데이터 그룹화
  grouped = {}
  for ts in test_scenarios:
    ts_id = ts.get("시나리오ID", "미분류")
    if ts_id not in grouped:
      grouped[ts_id] = []
    grouped[ts_id].append(ts)

  # 4. '양식' 시트를 복사하여 새 데이터 채우기
  for ts_id, group in grouped.items():
    sheet_title = str(ts_id)[:31]
    
    # 새 시트를 원하는 위치에 바로 생성
    new_ws = wb.create_sheet(title=sheet_title, index=insert_idx)
    insert_idx += 1

    copy_worksheet_template(source_ws, new_ws)
    
    common = group[0]

    # 상단 메타 정보 입력
    new_ws['B2'] = common.get("시스템", "")
    new_ws['F2'] = common.get("작성자", "")
    new_ws['B4'] = common.get("시나리오ID", "")
    new_ws['F4'] = common.get("시나리오명", "")
    new_ws['I4'] = common.get("요구사항_ID", "")

    # 7행부터 본문 데이터 입력
    for i, ts in enumerate(group):
      r = 7 + i 
      
      new_ws.cell(row=r, column=1, value=ts.get("케이스명", ""))
      new_ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
      
      data_mapping = [
        (3, i + 1, center_align),
        (4, ts.get("업무처리내용", ""), left_align),
        (5, ts.get("시험항목", ""), left_align),
        (6, ts.get("사전조건", ""), left_align),
        (7, "", left_align),
        (8, ts.get("예상결과", ""), left_align),
        (9, ts.get("화면ID", ""), left_align)
      ]

      # 셀 병합 및 테두리 복원
      for col in range(1, 10):
        cell = new_ws.cell(row=r, column=col)
        cell.border = border

        # 7행의 폰트를 새로 생성되는 행에 적용
        reference_font = new_ws.cell(row=7, column=col).font
        if reference_font:
          cell.font = copy.copy(reference_font)

        if col in [1, 2]:
           cell.alignment = center_align

      # 값 입력 및 정렬 적용
      for col, val, align in data_mapping:
        cell = new_ws.cell(row=r, column=col, value=str(val) if val else "")
        cell.alignment = align

  wb.save(full_path)
  print(f"통합시험 시나리오 저장 완료: {full_path.resolve()}")
  return full_path

def find_workbook_section_indices(wb):
  start_idx = -1
  end_idx = -1
  for i, name in enumerate(wb.sheetnames):
    if "개정이력" in name:
      start_idx = i
    elif "작성방법" in name:
      end_idx = i
  return start_idx, end_idx

def save_integration_test_results_to_excel(test_scenarios, base_workbook_path, output_dir: Path, result_sheet_form_path, base_filename="generated_TR"):
  if not test_scenarios:
    print("저장할 데이터가 없습니다.")
    return

  base_workbook_path = Path(base_workbook_path)
  output_dir = Path(output_dir)
  result_sheet_form_path = Path(result_sheet_form_path)

  if not base_workbook_path.exists():
    print(f"업로드된 기존 통합시험 결과서 파일을 찾을 수 없습니다: {base_workbook_path}")
    return

  if not result_sheet_form_path.exists():
    print(f"서버에 통합시험 결과서 표준 양식 파일({result_sheet_form_path})이 없습니다.")
    return

  output_dir.mkdir(parents=True, exist_ok=True)

  counter = 1
  while True:
    output_filename = f"{base_filename}_{counter}.xlsx"
    full_path = output_dir / output_filename
    if not full_path.exists():
      break
    counter += 1

  print(f"기존 통합시험 결과서 파일 로드 중: {base_workbook_path}")
  wb = openpyxl.load_workbook(base_workbook_path)

  print(f"외부 통합시험 결과서 표준 양식 로드 중: {result_sheet_form_path}")
  form_wb = openpyxl.load_workbook(result_sheet_form_path)
  source_ws = form_wb.active

  start_idx, end_idx = find_workbook_section_indices(wb)
  sheet_names = wb.sheetnames
  if start_idx != -1 and end_idx != -1 and start_idx < end_idx:
    sheets_to_remove = sheet_names[start_idx + 1 : end_idx]
    for sheet_name in sheets_to_remove:
      del wb[sheet_name]
    insert_idx = start_idx + 1
  else:
    insert_idx = 1

  border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
  center_align = Alignment(horizontal="center", vertical="center")
  left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
  max_result_column = max(source_ws.max_column, 13)

  grouped = {}
  for ts in test_scenarios:
    ts_id = ts.get("시나리오ID", "미분류")
    if ts_id not in grouped:
      grouped[ts_id] = []
    grouped[ts_id].append(ts)

  for ts_id, group in grouped.items():
    sheet_title = str(ts_id)[:31]
    new_ws = wb.create_sheet(title=sheet_title, index=insert_idx)
    insert_idx += 1

    copy_worksheet_template(source_ws, new_ws)
    common = group[0]

    new_ws['B2'] = common.get("시스템", "")
    new_ws['F2'] = common.get("작성자", "")
    new_ws['B4'] = common.get("시나리오ID", "")
    new_ws['F4'] = common.get("시나리오명", "")
    new_ws['I4'] = common.get("요구사항_ID", "")

    for i, ts in enumerate(group):
      r = 7 + i

      new_ws.cell(row=r, column=1, value=ts.get("케이스명", ""))
      new_ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)

      data_mapping = [
        (3, i + 1, center_align),
        (4, ts.get("업무처리내용", ""), left_align),
        (5, ts.get("시험항목", ""), left_align),
        (6, ts.get("사전조건", ""), left_align),
        (7, "", left_align),
        (8, ts.get("예상결과", ""), left_align),
        (9, ts.get("화면ID", ""), left_align),
      ]

      for col in range(1, max_result_column + 1):
        cell = new_ws.cell(row=r, column=col)
        cell.border = border
        reference_font = new_ws.cell(row=7, column=col).font
        if reference_font:
          cell.font = copy.copy(reference_font)
        if col in [1, 2, 3]:
          cell.alignment = center_align

      for col, val, align in data_mapping:
        cell = new_ws.cell(row=r, column=col, value=str(val) if val else "")
        cell.alignment = align

  form_wb.close()
  wb.save(full_path)
  print(f"통합시험 결과서 저장 완료: {full_path.resolve()}")
  return full_path

def generate_test_scenarios(
    template_xlsx_path: Path,
    tc_xlsx_path: Path,
    ui_pdf_path: Path | None,
    output_dir: Path,
    form_path: Path,
    req_mapping: dict[str, str] | None = None,
    unit_test_data: list[dict] | None = None,
    log_progress: bool = True,
    cancel_check: Callable[[], None] | None = None,
) -> dict:
    _check_cancel(cancel_check)
    template_xlsx_path = Path(template_xlsx_path)
    tc_xlsx_path = Path(tc_xlsx_path)
    output_dir = Path(output_dir)
    form_path = Path(form_path)

    output_dir.mkdir(parents=True, exist_ok=True)

    cover_author = extract_cover_author_from_workbook(template_xlsx_path)
    _check_cancel(cancel_check)
    _check_cancel(cancel_check)
    req_mapping = {}
    if ui_pdf_path:
        ui_pdf_path = Path(ui_pdf_path)
        if ui_pdf_path.exists():
            req_mapping = extract_req_mapping_from_pdf(ui_pdf_path)
            _check_cancel(cancel_check)

    unit_test_data = extract_unit_test_from_excel(tc_xlsx_path)
    _check_cancel(cancel_check)
    if not unit_test_data:
        return {
            "ok": False,
            "error": "단위시험 케이스 엑셀에서 데이터를 찾지 못했습니다.",
            "files": [],
        }

    missing_req_screen_ids = sorted({
        str(data.get("화면_ID", "")).strip()
        for data in unit_test_data
        if str(data.get("화면_ID", "")).strip()
        and str(data.get("화면_ID", "")).strip() not in req_mapping
    })
    if missing_req_screen_ids:
        return {
            "ok": False,
            "error": "사용자인터페이스설계서에서 화면 ID에 해당하는 요구사항 ID를 찾지 못했습니다.",
            "missing_screen_ids": missing_req_screen_ids[:10],
            "missing_screen_count": len(missing_req_screen_ids),
            "files": [],
        }

    for data in unit_test_data:
        screen_id = data.get("화면_ID", "").strip()
        data["요구사항_ID"] = req_mapping.get(screen_id, "")

    ts_result = build_test_scenarios_from_unit_tests(unit_test_data, author=cover_author)
    _check_cancel(cancel_check)

    if not ts_result:
        return {
            "ok": False,
            "error": "통합시험 시나리오로 변환할 수 있는 단위시험 케이스 데이터가 없습니다.",
            "files": [],
        }

    excel_path = save_test_scenarios_to_excel(
        ts_result,
        base_workbook_path=template_xlsx_path,
        output_dir=output_dir,
        scenario_sheet_form_path=form_path,
    )
    _check_cancel(cancel_check)

    files = []
    if excel_path:
        files.append({
            "kind": "xlsx",
            "path": str(excel_path),
            "name": excel_path.name,
        })

    return {
        "ok": True,
        "count": len(ts_result),
        "files": files,
    }

def generate_integration_test_results(
    template_xlsx_path: Path,
    tc_xlsx_path: Path,
    ui_pdf_path: Path | None,
    output_dir: Path,
    form_path: Path,
    cancel_check: Callable[[], None] | None = None,
) -> dict:
    _check_cancel(cancel_check)
    template_xlsx_path = Path(template_xlsx_path)
    tc_xlsx_path = Path(tc_xlsx_path)
    output_dir = Path(output_dir)
    form_path = Path(form_path)

    output_dir.mkdir(parents=True, exist_ok=True)

    cover_author = extract_cover_author_from_workbook(template_xlsx_path)
    req_mapping = dict(req_mapping or {})
    if not req_mapping and ui_pdf_path:
        ui_pdf_path = Path(ui_pdf_path)
        if ui_pdf_path.exists():
            req_mapping = extract_req_mapping_from_pdf(ui_pdf_path)
            _check_cancel(cancel_check)

    if unit_test_data is None:
        unit_test_data = extract_unit_test_from_excel(tc_xlsx_path)
    _check_cancel(cancel_check)
    if not unit_test_data:
        return {
            "ok": False,
            "error": "단위시험 케이스 엑셀에서 데이터를 찾지 못했습니다.",
            "files": [],
        }

    missing_req_screen_ids = sorted({
        str(data.get("화면_ID", "")).strip()
        for data in unit_test_data
        if str(data.get("화면_ID", "")).strip()
        and str(data.get("화면_ID", "")).strip() not in req_mapping
    })
    if missing_req_screen_ids:
        return {
            "ok": False,
            "error": "사용자인터페이스설계서에서 화면 ID에 해당하는 요구사항 ID를 찾지 못했습니다.",
            "missing_screen_ids": missing_req_screen_ids[:10],
            "missing_screen_count": len(missing_req_screen_ids),
            "files": [],
        }

    for data in unit_test_data:
        _check_cancel(cancel_check)
        _check_cancel(cancel_check)
        screen_id = data.get("화면_ID", "").strip()
        data["요구사항_ID"] = req_mapping.get(screen_id, "")

    result_rows = build_test_scenarios_from_unit_tests(unit_test_data, author=cover_author)
    _check_cancel(cancel_check)
    if not result_rows:
        return {
            "ok": False,
            "error": "통합시험 결과서로 변환할 수 있는 단위시험 케이스 데이터가 없습니다.",
            "files": [],
        }

    excel_path = save_integration_test_results_to_excel(
        result_rows,
        base_workbook_path=template_xlsx_path,
        output_dir=output_dir,
        result_sheet_form_path=form_path,
    )
    _check_cancel(cancel_check)

    files = []
    if excel_path:
        files.append({
            "kind": "xlsx",
            "path": str(excel_path),
            "name": excel_path.name,
        })

    return {
        "ok": True,
        "count": len(result_rows),
        "files": files,
    }
