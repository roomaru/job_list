#!/usr/bin/env python3
import html
import json
import re
import sys
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import quote, urljoin
from urllib.error import HTTPError
from urllib.request import Request, urlopen
from xml.sax.saxutils import escape


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR.parent
STATE_PATH = BASE_DIR / "job_postings_state.json"
XLSX_PATH = OUTPUT_DIR / "sap_abap_jobs.xlsx"

KEYWORDS = ("ABAP", "SAP ERP", "SAP", "ERP")
SEARCH_TERMS = (
    "SAP ABAP 신입",
    "SAP ERP 신입",
    "ABAP 신입",
    "SAP 운영 신입",
    "SAP ERP 운영 신입",
)
USER_AGENT = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"
INVALID_JOB_IDS = set()


def fetch(url):
    req = Request(url, headers={"User-Agent": USER_AGENT, "Accept-Language": "ko-KR,ko;q=0.9"})
    with urlopen(req, timeout=30) as response:
        return response.read().decode("utf-8", "ignore")


def clean_text(value):
    value = html.unescape(value or "")
    value = re.sub(r"<[^>]+>", " ", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def decode_escaped(value):
    value = value.replace(r"\u0026", "&")
    value = value.replace(r"\/", "/")
    value = value.replace(r"\"", '"')
    value = value.replace(r"\\", "\\")
    return html.unescape(value)


def is_relevant_job(title, details):
    if is_title_relevant(title, details):
        return True
    return False


def is_title_relevant(title, details=""):
    title_u = (title or "").upper()
    details_u = (details or "").upper()
    if any(term in title_u for term in ("ORACLE", "D365", "MS ERP")) and "SAP" not in title_u and "ABAP" not in title_u:
        return False
    if "ABAP" in title_u:
        return True
    if "SAP ERP" in title_u:
        return True
    if "SAP" in title_u and any(term in title_u for term in ("ERP", "운영", "AMS", "BASIS", "SUCCESSFACTORS")):
        return True
    if ("ERP 운영" in title or "ERP운영" in title or "ERP 개발" in title or "ERP개발" in title) and "SAP" in details_u:
        return True
    return False


def has_listing_sap_signal(details):
    details_u = (details or "").upper()
    return "ABAP" in details_u or "SAP" in details_u or "SAP ERP" in details_u


def strip_noise_sections(page):
    page = re.sub(r"<script[\s\S]*?</script>", " ", page, flags=re.I)
    page = re.sub(r"<style[\s\S]*?</style>", " ", page, flags=re.I)
    for marker in ("jv_footer", "jv_company", "similar_recruit", "recommend", "powered_by"):
        idx = page.find(marker)
        if idx != -1:
            page = page[:idx]
    return page


def clean_page_text(page):
    page = strip_noise_sections(page)
    text = html.unescape(re.sub(r"<[^>]+>", " ", page))
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def detail_has_relevant_role(text):
    text_u = (text or "").upper()
    if "ABAP" in text_u:
        return True
    direct_terms = ("SAP ERP", "SAP 운영", "SAP 개발", "SAP 모듈", "SAP AMS", "SAP BASIS", "SUCCESSFACTORS")
    if any(term in text_u for term in direct_terms):
        return True
    for match in re.finditer(r"ERP\s*(운영|개발|유지보수|컨설턴트|컨설팅)", text, flags=re.I):
        window = text_u[max(0, match.start() - 120) : match.end() + 120]
        if "SAP" in window:
            return True
    return False


def fetch_saramin_detail_text(rec_idx, searchword):
    data = (
        "rec_idx="
        + quote(rec_idx)
        + "&rec_seq=0&view_type=search&t_ref=&t_ref_content=&t_ref_scnid=&search_uuid=&refer=&searchType=search&searchword="
        + quote(searchword)
        + "&ref_dp=SRI_050_VIEW_MIX_RCT_NONMEM&dpId=&recommendRecIdx=&referNonce=&trainingStudentCode="
    ).encode()
    req = Request(
        "https://www.saramin.co.kr/zf_user/jobs/relay/view-ajax",
        data=data,
        headers={
            "User-Agent": USER_AGENT,
            "Accept-Language": "ko-KR,ko;q=0.9",
            "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
            "X-Requested-With": "XMLHttpRequest",
            "Referer": "https://www.saramin.co.kr/zf_user/jobs/relay/view?rec_idx=" + rec_idx,
        },
    )
    with urlopen(req, timeout=30) as response:
        return clean_page_text(response.read().decode("utf-8", "ignore"))


def fetch_detail_text(url):
    return clean_page_text(fetch(url))


def looks_entry_level(*parts):
    text = " ".join(part or "" for part in parts)
    return (
        "신입" in text
        or "경력무관" in text
        or "년수무관" in text
        or "ENTRY_NEWBIE" in text
        or '"isStarter":"1"' in text
        or r'\"isStarter\":\"1\"' in text
    )


def normalize_deadline(period):
    if not period:
        return ""
    period = clean_text(period)
    if "2070-01-01" in period or "상시" in period:
        return "상시채용"
    return period


def parse_iso_date(value):
    if not value:
        return None
    matches = re.findall(r"(\d{4})-(\d{2})-(\d{2})", value)
    if not matches:
        return None
    match = matches[-1]
    return date(int(match[0]), int(match[1]), int(match[2]))


def parse_local_date(value):
    if not value:
        return None
    try:
        return date.fromisoformat(value[:10])
    except ValueError:
        return None


def relative_deadline_date(period, reference_date):
    if "오늘마감" in period:
        return reference_date
    if "내일마감" in period:
        return reference_date + timedelta(days=1)
    return None


def is_past_deadline(period, today, reference_date=None):
    if not period or period == "상시채용":
        return False
    reference_date = reference_date or today
    relative = relative_deadline_date(period, reference_date)
    if relative:
        return relative < today
    parsed = parse_iso_date(period)
    if parsed:
        return parsed < today
    match = re.search(r"~\s*(\d{2})/(\d{2})", period)
    if match:
        candidate = date(today.year, int(match.group(1)), int(match.group(2)))
        if candidate < today and today.month >= 11 and candidate.month <= 2:
            candidate = date(today.year + 1, candidate.month, candidate.day)
        return candidate < today
    return False


def extract_escaped(block, field):
    match = re.search(rf'\\"{re.escape(field)}\\":\\"(.*?)\\"', block)
    return decode_escaped(match.group(1)) if match else ""


def extract_boolish(block, field):
    match = re.search(rf'\\"{re.escape(field)}\\":(true|false|null|\\".*?\\")', block)
    return match.group(1) if match else ""


def parse_jobkorea(term):
    url = "https://www.jobkorea.co.kr/Search/?stext=" + quote(term)
    page = fetch(url)
    jobs = []
    for match in re.finditer(r'\\"id\\":\\"(\d+)\\",\\"legacyJobNo\\":\\"(.*?)\\"', page):
        start = match.start()
        block = page[start : start + 5000]
        if r"applicationPeriod" not in block:
            continue
        job_id = match.group(1)
        legacy_no = decode_escaped(match.group(2))
        title = extract_escaped(block, "title")
        company = extract_escaped(block, "postingCompanyName") or extract_escaped(block, "companyName")
        start_date = extract_escaped(block, "start")
        end_date = extract_escaped(block, "end")
        feature = " / ".join(
            clean_text(extract_escaped(block, field))
            for field in (
                "_internal_featureWorkCode",
                "_internal_featureIndustryCode",
                "_internal_featureToolCode",
                "_internal_featureSkillCode",
                "jobClassificationOrIndustry",
            )
            if extract_escaped(block, field)
        )
        entry_level = looks_entry_level(title, feature, block)
        condition = "신입 가능" if entry_level else ""
        if feature:
            condition = (condition + " / " + feature).strip(" /")
        if not (title and company and legacy_no):
            continue
        job_key = "jobkorea:" + legacy_no
        link = "https://www.jobkorea.co.kr/Recruit/GI_Read/" + legacy_no
        direct_match = is_title_relevant(title, feature)
        detail_match = False
        detail_text = ""
        try:
            detail_text = fetch_detail_text(link)
        except HTTPError as exc:
            if exc.code in (404, 410):
                INVALID_JOB_IDS.add(job_key)
                continue
        except Exception:
            detail_text = ""
        if not direct_match and has_listing_sap_signal(feature):
            detail_match = detail_has_relevant_role(detail_text)
        if not (direct_match or detail_match):
            continue
        if not entry_level:
            continue
        jobs.append(
            {
                "id": job_key,
                "site": "잡코리아",
                "title": title,
                "link": link,
                "company": company,
                "condition": (condition or "신입 / SAP 관련") + (" / 상세 본문 확인" if detail_match else ""),
                "period": normalize_deadline(f"{start_date[:10]} ~ {end_date[:10]}"),
            }
        )
    return jobs


def parse_saramin(term):
    url = "https://www.saramin.co.kr/zf_user/search/recruit?searchType=search&searchword=" + quote(term)
    page = fetch(url)
    blocks = re.split(r'<div[^>]+class="item_recruit"[^>]*', page)[1:]
    jobs = []
    for block in blocks:
        title_match = re.search(r'<h2 class="job_tit">.*?<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', block, re.S)
        rec_match = re.search(r"rec_idx=(\d+)", title_match.group(1) if title_match else block)
        company_match = re.search(r'<strong class="corp_name">.*?<a[^>]*>(.*?)</a>', block, re.S)
        date_match = re.search(r'<span class="date">(.*?)</span>', block, re.S)
        condition_match = re.search(r'<div class="job_condition">(.*?)</div>', block, re.S)
        sector_match = re.search(r'<div class="job_sector">(.*?)</div>', block, re.S)
        if not (rec_match and title_match and company_match):
            continue
        title = clean_text(title_match.group(2))
        link = urljoin("https://www.saramin.co.kr", html.unescape(title_match.group(1)))
        company = clean_text(company_match.group(1))
        condition = " / ".join(
            part for part in (clean_text(condition_match.group(1) if condition_match else ""), clean_text(sector_match.group(1) if sector_match else "")) if part
        )
        period = normalize_deadline(clean_text(date_match.group(1) if date_match else ""))
        direct_match = is_title_relevant(title, condition)
        detail_match = False
        if not direct_match and has_listing_sap_signal(condition):
            try:
                detail_match = detail_has_relevant_role(fetch_saramin_detail_text(rec_match.group(1), term))
            except Exception:
                detail_match = False
        if not (direct_match or detail_match):
            continue
        if not looks_entry_level(title, condition):
            continue
        jobs.append(
            {
                "id": "saramin:" + rec_match.group(1),
                "site": "사람인",
                "title": title,
                "link": link,
                "company": company,
                "condition": (condition or "신입 / SAP 관련") + (" / 상세 본문 확인" if detail_match else ""),
                "period": period,
            }
        )
    return jobs


def load_state():
    if not STATE_PATH.exists():
        return {"jobs": {}}
    try:
        return json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {"jobs": {}}


def save_state(state):
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def collect_jobs():
    collected = {}
    errors = []
    for term in SEARCH_TERMS:
        for parser in (parse_jobkorea, parse_saramin):
            try:
                for job in parser(term):
                    collected[job["id"]] = job
            except Exception as exc:
                errors.append(f"{parser.__name__}({term}): {exc}")
    return collected, errors


def merge_state(state, collected, today):
    today_s = today.isoformat()
    jobs = state.setdefault("jobs", {})
    for job_id, job in collected.items():
        existing = jobs.get(job_id, {})
        existing.update(job)
        existing.setdefault("first_seen", today_s)
        existing["last_seen"] = today_s
        existing["missing_count"] = 0
        existing["status"] = "마감" if is_past_deadline(job.get("period", ""), today) else "모집중"
        jobs[job_id] = existing

    for job_id, job in list(jobs.items()):
        if job_id in INVALID_JOB_IDS:
            job["missing_count"] = int(job.get("missing_count", 0)) + 1
            job["status"] = "마감"
            continue
        if job_id in collected:
            continue
        job["missing_count"] = int(job.get("missing_count", 0)) + 1
        reference_date = parse_local_date(job.get("last_seen", "")) or today
        if is_past_deadline(job.get("period", ""), today, reference_date) or job["missing_count"] >= 3:
            job["status"] = "마감"
    state["last_updated"] = datetime.now().isoformat(timespec="seconds")
    return state


def col_name(index):
    name = ""
    while index:
        index, rem = divmod(index - 1, 26)
        name = chr(65 + rem) + name
    return name


def safe_xml_text(value):
    text = str(value or "")
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    return escape(text)


def safe_xml_attr(value):
    return escape(str(value or ""), {'"': "&quot;"})


def sheet_xml(rows, rel_ids):
    widths = [34, 58, 24, 52, 24, 12, 14, 14]
    cols = "".join(f'<col min="{i+1}" max="{i+1}" width="{width}" customWidth="1"/>' for i, width in enumerate(widths))
    xml_rows = []
    for r_idx, row in enumerate(rows, start=1):
        cells = []
        for c_idx, value in enumerate(row, start=1):
            ref = f"{col_name(c_idx)}{r_idx}"
            style = ' s="1"' if r_idx == 1 else ""
            text = safe_xml_text(value)
            cells.append(f'<c r="{ref}" t="inlineStr"{style}><is><t>{text}</t></is></c>')
        xml_rows.append(f'<row r="{r_idx}">{"".join(cells)}</row>')
    hyperlinks = "".join(f'<hyperlink ref="{cell}" r:id="{rid}"/>' for cell, rid in rel_ids)
    hyperlink_xml = f"<hyperlinks>{hyperlinks}</hyperlinks>" if hyperlinks else ""
    last_row = max(len(rows), 1)
    dimension = f"A1:H{last_row}"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<dimension ref="{dimension}"/>'
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f"<cols>{cols}</cols>"
        f'<sheetData>{"".join(xml_rows)}</sheetData>{hyperlink_xml}</worksheet>'
    )


def sheet_rels(rows):
    rels = []
    for idx, row in enumerate(rows[1:], start=2):
        if len(row) > 1 and row[1]:
            rels.append((f"B{idx}", f"rId{len(rels)+1}", row[1]))
    xml = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
    xml.append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
    for _, rid, target in rels:
        xml.append(f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{safe_xml_attr(target)}" TargetMode="External"/>')
    xml.append("</Relationships>")
    return [(cell, rid) for cell, rid, _ in rels], "".join(xml)


def write_xlsx(state):
    headers = ["공고명", "공고 링크", "회사", "조건", "모집 기간", "사이트", "최초 수집일", "최근 확인일"]
    rows_by_status = {"모집중": [headers], "마감": [headers]}
    items = sorted(
        state["jobs"].values(),
        key=lambda item: (item.get("status") != "모집중", item.get("site", ""), item.get("company", ""), item.get("title", "")),
    )
    for job in items:
        status = "마감" if job.get("status") == "마감" else "모집중"
        rows_by_status[status].append(
            [
                job.get("title", ""),
                job.get("link", ""),
                job.get("company", ""),
                job.get("condition", ""),
                job.get("period", ""),
                job.get("site", ""),
                job.get("first_seen", ""),
                job.get("last_seen", ""),
            ]
        )

    active_rels, active_rel_xml = sheet_rels(rows_by_status["모집중"])
    closed_rels, closed_rel_xml = sheet_rels(rows_by_status["마감"])
    active_xml = sheet_xml(rows_by_status["모집중"], active_rels)
    closed_xml = sheet_xml(rows_by_status["마감"], closed_rels)

    with zipfile.ZipFile(XLSX_PATH, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", ROOT_RELS)
        zf.writestr("xl/workbook.xml", WORKBOOK_XML)
        zf.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        zf.writestr("xl/styles.xml", STYLES_XML)
        zf.writestr("xl/worksheets/sheet1.xml", active_xml)
        zf.writestr("xl/worksheets/sheet2.xml", closed_xml)
        zf.writestr("xl/worksheets/_rels/sheet1.xml.rels", active_rel_xml)
        zf.writestr("xl/worksheets/_rels/sheet2.xml.rels", closed_rel_xml)


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

WORKBOOK_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="모집중" sheetId="1" r:id="rId1"/>
<sheet name="마감" sheetId="2" r:id="rId2"/>
</sheets>
</workbook>"""

WORKBOOK_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="11"/><name val="맑은 고딕"/></font><font><b/><sz val="11"/><name val="맑은 고딕"/><color rgb="FFFFFFFF"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def main():
    today = date.today()
    state = load_state()
    collected, errors = collect_jobs()
    state = merge_state(state, collected, today)
    if errors:
        state["last_errors"] = errors
    else:
        state.pop("last_errors", None)
    save_state(state)
    write_xlsx(state)
    active = sum(1 for job in state["jobs"].values() if job.get("status") == "모집중")
    closed = sum(1 for job in state["jobs"].values() if job.get("status") == "마감")
    print(f"updated {XLSX_PATH} active={active} closed={closed} collected={len(collected)}")
    if errors:
        print("warnings:")
        for error in errors:
            print(" - " + error)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
