import os
import sys
import json
import base64
import logging
import urllib.request
import urllib.error
import urllib.parse
from datetime import datetime
import xlwt
import xlrd
try:
    import win32com.client
except ImportError:
    win32com = None
from playwright.sync_api import sync_playwright

# Setup logging
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
log_dir = os.path.join(BASE_DIR, "logs")
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, "absent_tracker.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(log_file, encoding="utf-8"),
        logging.StreamHandler(sys.stdout)
    ]
)

CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

def load_config():
    default_config = {
        "erp_url": "https://online.myschoolerp.com/Login.aspx",
        "erp_accounts": [
            {"name": "Hindi Medium (HM)", "username": "MY1386_2_BABU", "password": "Ssvm@2003"},
            {"name": "English Medium (EM)", "username": "MY1386_4_BABU", "password": "Ssvm@2003"}
        ],
        "export_directory": BASE_DIR,
        "logo_path": os.path.join(BASE_DIR, "SSVM Logo 2000x.jpg"),
        "filename_suffix": "absent.xls",
        "date_format": "%d%m%Y",
        "github_owner": "blc2050",
        "github_repo": "ssvm",
        "github_branch": "main",
        "github_token": ""
    }
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                default_config.update(cfg)
        except Exception as e:
            logging.warning(f"Could not load config file, using defaults: {e}")

    # Fallback if configured directory or logo doesn't exist
    export_dir = default_config.get("export_directory", "")
    if not export_dir or not os.path.exists(export_dir):
        default_config["export_directory"] = BASE_DIR

    logo_file = default_config.get("logo_path", "")
    if not logo_file or not os.path.exists(logo_file):
        for candidate in ["SSVM Logo 2000x.jpg", "logo.jpg"]:
            cand_path = os.path.join(BASE_DIR, candidate)
            if os.path.exists(cand_path):
                default_config["logo_path"] = cand_path
                break

    return default_config

def get_formatted_status_text(dt=None):
    if dt is None:
        dt = datetime.now()
    # Format: Day (DD Month Name YYYY) -> e.g. Thursday (06 August 2026)
    return dt.strftime("%A (%d %B %Y)")

def extract_rows_from_raw_excel(raw_file_path, status_text):
    """
    Reads exported ERP excel file.
    The ERP exports an HTML table saved with a .xls extension.
    This parses the HTML table directly (fast, resilient, no Excel COM process lock).
    Falls back to xlrd and Excel COM if necessary.
    """
    logging.info(f"Reading exported file: {raw_file_path}")
    
    # Strategy 1: Direct HTML Table Parsing (Standard for myschoolerp export)
    try:
        with open(raw_file_path, "r", encoding="utf-8", errors="ignore") as f:
            preview = f.read(4096)
        if "<table" in preview.lower():
            import lxml.html
            with open(raw_file_path, "r", encoding="utf-8", errors="ignore") as f:
                tree = lxml.html.fromstring(f.read())
            table = tree.find(".//table")
            if table is not None:
                headers = [th.text_content().strip() for th in table.xpath(".//tr[1]/th")]
                col_map = {h: idx for idx, h in enumerate(headers)}
                records = []
                for tr in table.xpath(".//tr[position()>1]"):
                    cells = [td.text_content().strip() for td in tr.xpath("./td")]
                    if not cells or len(cells) < len(headers):
                        continue
                    scholar = cells[col_map.get("Scholar No", 0)]
                    name = cells[col_map.get("Student Name", 1)]
                    father = cells[col_map.get("Father Name", 2)]
                    cls = cells[col_map.get("Class Name", 3)]
                    mobile = cells[col_map.get("Mobile No", 4)]
                    if not name and not scholar:
                        continue
                    records.append({
                        "Mobile No": mobile,
                        "Student Name": name,
                        "Father Name": father,
                        "Scholar No": scholar,
                        "Class Name": cls,
                        "Status": status_text
                    })
                if records:
                    logging.info(f"Successfully extracted {len(records)} records via HTML table parser.")
                    return records
    except Exception as html_err:
        logging.warning(f"HTML table parsing failed, falling back: {html_err}")

    # Strategy 2: xlrd (Binary BIFF8 Excel)
    try:
        wb = xlrd.open_workbook(raw_file_path)
        ws = wb.sheet_by_index(0)
        headers = [str(ws.cell_value(0, c)).strip() for c in range(ws.ncols)]
        col_map = {h: idx for idx, h in enumerate(headers) if h}
        records = []
        for r in range(1, ws.nrows):
            scholar = ws.cell_value(r, col_map.get("Scholar No", 0))
            name = ws.cell_value(r, col_map.get("Student Name", 1))
            father = ws.cell_value(r, col_map.get("Father Name", 2))
            cls = ws.cell_value(r, col_map.get("Class Name", 3))
            mobile = ws.cell_value(r, col_map.get("Mobile No", 4))
            if not name and not scholar:
                continue
            records.append({
                "Mobile No": mobile,
                "Student Name": name,
                "Father Name": father,
                "Scholar No": scholar,
                "Class Name": cls,
                "Status": status_text
            })
        if records:
            logging.info(f"Successfully extracted {len(records)} records via xlrd.")
            return records
    except Exception:
        pass

    # Strategy 3: Safe Excel COM Fallback (Windows only)
    if win32com is None:
        return records

    excel = None
    records = []
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        try:
            excel.Visible = False
            excel.DisplayAlerts = False
        except Exception:
            pass
        wb = excel.Workbooks.Open(os.path.abspath(raw_file_path))
        ws = wb.Worksheets(1)
        row_count = ws.UsedRange.Rows.Count
        col_count = ws.UsedRange.Columns.Count
        headers = [ws.Cells(1, c).Value for c in range(1, col_count + 1)]
        col_map = {str(h).strip(): idx for idx, h in enumerate(headers, start=1) if h}
        for r in range(2, row_count + 1):
            scholar_no = ws.Cells(r, col_map.get("Scholar No", 1)).Value
            student_name = ws.Cells(r, col_map.get("Student Name", 2)).Value
            father_name = ws.Cells(r, col_map.get("Father Name", 3)).Value
            class_name = ws.Cells(r, col_map.get("Class Name", 4)).Value
            mobile_no = ws.Cells(r, col_map.get("Mobile No", 5)).Value
            if not student_name and not scholar_no:
                continue
            records.append({
                "Mobile No": mobile_no,
                "Student Name": student_name,
                "Father Name": father_name,
                "Scholar No": scholar_no,
                "Class Name": class_name,
                "Status": status_text
            })
        wb.Close(False)
    finally:
        if excel:
            try:
                excel.Quit()
            except Exception:
                pass
            del excel
        
    return records

def save_combined_excel(all_records, target_excel_path):
    """
    Saves combined records to target_excel_path using xlwt BIFF8 format.
    """
    wb_out = xlwt.Workbook(encoding="utf-8")
    ws_out = wb_out.add_sheet("Sheet1")
    
    target_columns = ["Medium", "Mobile No", "Student Name", "Father Name", "Scholar No", "Class Name", "Status"]
    
    header_style = xlwt.easyxf("font: name Aptos Narrow, bold on, height 220; align: horiz left;")
    general_style = xlwt.easyxf("font: name Aptos Narrow, height 220; align: horiz left;")
    text_style = xlwt.easyxf("font: name Aptos Narrow, height 220; align: horiz left;", num_format_str="@")
    
    # Headers
    for c_idx, col_name in enumerate(target_columns):
        ws_out.write(0, c_idx, col_name, header_style)
        
    # Column widths matching 05082026absent.xls
    col_widths = [6000, 4500, 9500, 9000, 4200, 4500, 9500]
    for idx, w in enumerate(col_widths):
        ws_out.col(idx).width = w

    # Write Records
    for r_idx, row_dict in enumerate(all_records, start=1):
        ws_out.write(r_idx, 0, str(row_dict.get("Medium") or ""), text_style)
        ws_out.write(r_idx, 1, row_dict["Mobile No"], general_style if isinstance(row_dict["Mobile No"], (int, float)) else text_style)
        ws_out.write(r_idx, 2, str(row_dict["Student Name"] or ""), text_style)
        ws_out.write(r_idx, 3, str(row_dict["Father Name"] or ""), text_style)
        ws_out.write(r_idx, 4, row_dict["Scholar No"], general_style if isinstance(row_dict["Scholar No"], (int, float)) else text_style)
        ws_out.write(r_idx, 5, str(row_dict["Class Name"] or ""), text_style)
        ws_out.write(r_idx, 6, str(row_dict["Status"]), text_style)

    wb_out.save(target_excel_path)
    logging.info(f"Successfully saved combined Excel file: {target_excel_path}")

def run_part_1_erp_extraction(config, p):
    """
    PART 1: Log in to HM & EM ERP accounts, scrape attendance, format & save excel file.
    """
    today_str = datetime.now().strftime(config["date_format"])
    output_filename = f"{today_str}{config['filename_suffix']}"
    output_filepath = os.path.join(config["export_directory"], output_filename)
    status_text = get_formatted_status_text()

    logging.info("========================================================")
    logging.info("PART 1: ERP ATTENDANCE SCRAPING (HM + EM)")
    logging.info("========================================================")

    combined_records = []
    account_counts = {}
    temp_files = []

    try:
        browser = p.chromium.launch(headless=True)
    except Exception:
        browser = p.chromium.launch(channel="msedge", headless=True)

    for acc in config["erp_accounts"]:
        acc_name = acc["name"]
        username = acc["username"]
        password = acc["password"]
        
        logging.info(f"--> Scraping account: {acc_name} ({username})")
        page = browser.new_page()
        page.on("dialog", lambda d: d.accept())

        # Step 1: Login ERP
        page.goto(config["erp_url"], timeout=60000, wait_until="domcontentloaded")
        page.fill("input[name*='User']", username)
        page.fill("input[type='password']", password)
        page.click("input[type='submit']")
        try:
            page.wait_for_url(lambda u: "login.aspx" not in u.lower(), timeout=30000)
        except Exception:
            logging.warning(f"[{acc_name}] Page remained on Login.aspx after submit; waiting for networkidle.")
            try:
                page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass

        # Step 2: Navigate to Continuous Attendance page (with retries & session recovery)
        report_url = "https://online.myschoolerp.com/StudentAttendence/ContinuousAttendance.aspx"
        for nav_attempt in range(1, 4):
            try:
                page.goto(report_url, wait_until="domcontentloaded", timeout=60000)
                # If redirected back to Login.aspx, session was dropped; re-authenticate
                if "login.aspx" in page.url.lower():
                    logging.warning(f"[{acc_name}] Redirected to Login.aspx; re-authenticating...")
                    page.fill("input[name*='User']", username)
                    page.fill("input[type='password']", password)
                    page.click("input[type='submit']")
                    try:
                        page.wait_for_url(lambda u: "login.aspx" not in u.lower(), timeout=25000)
                    except Exception:
                        pass
                    page.goto(report_url, wait_until="domcontentloaded", timeout=60000)

                # Ensure #btnSearch is present on the page before proceeding
                page.wait_for_selector("#btnSearch", timeout=20000, state="attached")
                break
            except Exception as nav_err:
                logging.warning(f"[{acc_name}] Report page navigation attempt {nav_attempt} failed: {nav_err}")
                if nav_attempt == 3:
                    raise
                page.wait_for_timeout(4000)
        page.wait_for_timeout(1000)

        # Step 3: Search & Export
        for search_attempt in range(1, 4):
            try:
                if not page.query_selector("#btnSearch"):
                    logging.warning(f"[{acc_name}] #btnSearch not present on {page.url}; re-navigating to report page...")
                    page.goto(report_url, wait_until="domcontentloaded", timeout=45000)
                    page.wait_for_selector("#btnSearch", timeout=20000, state="attached")
                page.click("#btnSearch", timeout=30000)
                page.wait_for_load_state("networkidle", timeout=30000)
                break
            except Exception as search_err:
                logging.warning(f"[{acc_name}] Search click attempt {search_attempt} failed: {search_err}")
                if search_attempt == 3:
                    raise
                page.wait_for_timeout(4000)
        page.wait_for_timeout(2000)

        temp_path = os.path.join(config["export_directory"], f"temp_{username}_{today_str}.xls")
        temp_files.append(temp_path)

        # Detect empty result grid (no absent students today)
        grid_text = ""
        try:
            grid_text = page.inner_text("#grdSearch")
        except Exception:
            pass
        if "No Record Found" in grid_text or not grid_text.strip():
            logging.warning(f"[{acc_name}] No absent records found for today; skipping export.")
            page.close()
            account_counts[acc_name] = 0
            continue

        # Step 3b: Download with retries (handles transient failures / new-tab downloads)
        download = None
        for attempt in range(1, 4):
            try:
                with page.expect_download(timeout=25000) as download_info:
                    page.click("#BtnExportToxls")
                download = download_info.value
                break
            except Exception as e:
                logging.warning(f"[{acc_name}] Download attempt {attempt} failed: {e}")
                # Some ERP versions trigger the download from a new tab
                for candidate in browser.contexts[-1].pages:
                    if candidate != page:
                        try:
                            with candidate.expect_download(timeout=15000) as dl2:
                                pass
                            download = dl2.value
                            break
                        except Exception:
                            candidate.close()
                if download:
                    break
                page.wait_for_timeout(3000)
        if download is None:
            raise RuntimeError(f"[{acc_name}] Could not download export after 3 attempts")
        download.save_as(temp_path)
        
        page.close()

        # Step 4: Extract records
        acc_records = extract_rows_from_raw_excel(temp_path, status_text)
        for rec in acc_records:
            rec["Medium"] = acc_name
        account_counts[acc_name] = len(acc_records)
        logging.info(f"[{acc_name}] Extracted count: {len(acc_records)}")

        # Append EM below HM
        combined_records.extend(acc_records)

    browser.close()

    # Step 5: Save final combined Excel file
    save_combined_excel(combined_records, output_filepath)

    # Clean up temp files
    for tf in temp_files:
        if os.path.exists(tf):
            try:
                os.remove(tf)
            except Exception:
                pass

    # Step 6: Verify Counts (HM + EM = Total)
    hm_count = account_counts.get(config["erp_accounts"][0]["name"], 0)
    em_count = account_counts.get(config["erp_accounts"][1]["name"], 0)
    total_count = len(combined_records)

    logging.info("--------------------------------------------------------")
    logging.info("COUNT VERIFICATION")
    logging.info(f"  Hindi Medium (HM) Count    : {hm_count}")
    logging.info(f"  English Medium (EM) Count  : {em_count}")
    logging.info(f"  Combined Total in Excel    : {total_count}")
    if hm_count + em_count == total_count:
        logging.info(f"SUCCESS VERIFIED: HM ({hm_count}) + EM ({em_count}) = Total ({total_count})")
    else:
        logging.error(f"MISMATCH DETECTED: HM ({hm_count}) + EM ({em_count}) != Total ({total_count})")
    logging.info("--------------------------------------------------------")

    return output_filepath

def normalize_mobile(mobile):
    """Normalize mobile number to international format for wa.me/tel links."""
    if mobile is None:
        return None
    if isinstance(mobile, float) and mobile.is_integer():
        mobile = int(mobile)
    num = str(mobile).strip().replace(" ", "").replace("-", "")
    if not num or not num.isdigit():
        return None
    if num.startswith("+"):
        num = num[1:]
    if num.startswith("91") and len(num) == 12:
        return num
    if num.startswith("0") and len(num) == 11:
        num = num[1:]
    if len(num) == 10:
        return "91" + num
    return num

CLASS_LIST = [
    "NURSERY", "1ST", "2ND", "3RD", "4TH", "5TH",
    "6TH A", "6TH B", "7TH A", "7TH B", "8TH A", "8TH B",
    "9TH A", "9TH B", "10TH A", "10TH B", "10TH C",
    "11TH SCI", "11TH ARTS", "11TH AGR",
    "12TH AGR", "12TH SCI", "12TH ARTS",
]

def class_order_index(class_name):
    """Index of a class in the school master class list (from attendance report images)."""
    c = str(class_name or "").strip().upper()
    if c in CLASS_LIST:
        return CLASS_LIST.index(c)
    return 1000

def medium_short(medium):
    """Short display name for a medium."""
    m = str(medium or "")
    if "English" in m:
        return "English Medium"
    if "Hindi" in m:
        return "Hindi Medium"
    return (m or "Hindi Medium").strip()

def classify_class(class_name):
    """Classify a class name into: 10th, 12th or other."""
    if not class_name:
        return "other"
    c = str(class_name).strip().upper()
    if c.startswith("10") or c.startswith("X "):
        return "10th"
    if c.startswith("12") or c.startswith("XII"):
        return "12th"
    return "other"

def group_by_class_medium(records):
    """Groups records by (class, medium), sorted by master class order, then name-wise."""
    groups = {}
    for rec in records:
        cls = str(rec.get("Class Name") or "").strip() or "UNKNOWN"
        med = medium_short(rec.get("Medium"))
        groups.setdefault((cls, med), []).append(rec)
    for lst in groups.values():
        lst.sort(key=lambda r: str(r.get("Student Name") or "").lower())
    return sorted(groups.items(), key=lambda kv: (class_order_index(kv[0][0]), kv[0][1]))

def class_records_map(records):
    """Maps class name -> list of records (all mediums)."""
    groups = {}
    for rec in records:
        cls = str(rec.get("Class Name") or "").strip() or "UNKNOWN"
        groups.setdefault(cls, []).append(rec)
    return groups

def read_records_from_excel(excel_path):
    """Reads combined excel file back into a list of record dicts."""
    records = []
    wb = xlrd.open_workbook(excel_path)
    ws = wb.sheet_by_index(0)
    headers = [str(ws.cell_value(0, c)).strip() for c in range(ws.ncols)]
    col_map = {h: idx for idx, h in enumerate(headers) if h}
    for r in range(1, ws.nrows):
        record = {h: ws.cell_value(r, col_map[h]) for h in col_map}
        if record.get("Student Name") or record.get("Scholar No"):
            records.append(record)
    return records

def build_whatsapp_message(rec):
    """Builds the pre-filled WhatsApp message text for a student record."""
    name = str(rec.get("Student Name") or "").strip()
    father = str(rec.get("Father Name") or "").strip()
    scholar_raw = rec.get("Scholar No")
    if isinstance(scholar_raw, float) and scholar_raw.is_integer():
        scholar_raw = int(scholar_raw)
    cls = str(rec.get("Class Name") or "").strip()
    date_text = str(rec.get("Status") or "").strip()
    return (
        "*सरस्वती उच्च माध्यमिक विद्या मंदिर, मण्डली*\n"
        "कल्याणपुर (बालोतरा) | 📞 9413030806, 8854030806\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        "*दैनिक अनुपस्थिति सूचना (सत्र: 2026-27)*\n\n"
        "आदरणीय अभिभावक, सादर प्रणाम।\n"
        "आपका सुपुत्र/सुपुत्री:\n"
        f"👤 *विद्यार्थी का नाम:* {name}\n"
        f"👨 *पिता का नाम:* {father}\n"
        f"📋 *प्रवेश क्रमांक (SR No.):* {scholar_raw}\n"
        f"🏫 *कक्षा:* {cls}\n"
        f"📅 *दिनांक:* {date_text}\n"
        "आज विद्यालय में *अनुपस्थित* रहे हैं।\n\n"
        "📢 *अनुपस्थिति के नियम:*\n"
        "● दैनिक उपस्थिति अनिवार्य है।\n"
        "● नियमित उपस्थिति से ही विद्यार्थी की शैक्षणिक प्रगति संभव है।\n"
        "● अवकाश की स्थिति में विद्यालय में पूर्व सूचना देना अनिवार्य है।\n\n"
        "🙏 *निवेदन:* विद्यार्थी को प्रतिदिन नियमित रूप से विद्यालय भेजें।\n\n"
        "*धन्यवाद!*\n"
        "*श्री सरस्वती उच्च माध्यमिक विद्या मंदिर, मण्डली*"
    )

def build_html(records, status_text):
    """Builds a modern glassy mobile-first HTML page with 3 tabs:
    Board Class (10th & 12th), Local Class (all others), Attendance Report (class-wise counts)."""
    all_groups = group_by_class_medium(records)
    board_groups = [(key, rows) for key, rows in all_groups if classify_class(key[0]) in ("10th", "12th")]
    local_groups = [(key, rows) for key, rows in all_groups if classify_class(key[0]) not in ("10th", "12th")]

    def render_rows(rows):
        html = []
        for rec in rows:
            mobile = normalize_mobile(rec.get("Mobile No"))
            name = str(rec.get("Student Name") or "").strip()
            father = str(rec.get("Father Name") or "").strip()
            scholar_raw = rec.get("Scholar No")
            if isinstance(scholar_raw, float) and scholar_raw.is_integer():
                scholar_raw = int(scholar_raw)
            scholar = str(scholar_raw or "").strip()
            cls = str(rec.get("Class Name") or "").strip()
            if mobile:
                wa_msg = urllib.parse.quote(build_whatsapp_message(rec), safe="")
                wa_link = f"https://wa.me/{mobile}?text={wa_msg}"
                tel_num = mobile[2:] if len(mobile) == 12 and mobile.startswith("91") else mobile
                tel_link = f"tel:{tel_num}"
                actions = (
                    f'<a class="btn wa" href="{wa_link}" target="_blank" rel="noopener" '
                    f'title="WhatsApp"><svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor">'
                    f'<path d="M12 2a10 10 0 0 0-8.6 15.1L2 22l5-1.3A10 10 0 1 0 12 2zm0 18.2a8.2 8.2 0 0 1-4.2-1.2l-.3-.2-3 .8.8-2.9-.2-.3A8.2 8.2 0 1 1 12 20.2zm4.5-6.1c-.2-.1-1.5-.7-1.7-.8s-.4-.1-.6.1-.7.8-.8 1-.3.2-.5.1a6.7 6.7 0 0 1-2-1.2 7.5 7.5 0 0 1-1.4-1.7c-.1-.2 0-.4.1-.5l.4-.5a2 2 0 0 0 .3-.4.4.4 0 0 0 0-.4c0-.1-.6-1.4-.8-1.9s-.4-.4-.6-.4h-.5a1 1 0 0 0-.7.3 3 3 0 0 0-.9 2.2 5.2 5.2 0 0 0 1.1 2.7 11.8 11.8 0 0 0 4.5 4 5.2 5.2 0 0 0 3.1 1c1 0 1.5-.7 1.7-1.4s.2-1.2.1-1.4-.3-.2-.5-.3z"/></svg> '
                    f'<span>WhatsApp</span></a>'
                    f'<a class="btn call" href="{tel_link}" title="Call"><svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor">'
                    f'<path d="M6.6 10.8a15.5 15.5 0 0 0 6.6 6.6l2.2-2.2a1 1 0 0 1 1-.2 11.4 11.4 0 0 0 3.6.6 1 1 0 0 1 1 1V20a1 1 0 0 1-1 1A17 17 0 0 1 3 4a1 1 0 0 1 1-1h3.5a1 1 0 0 1 1 1 11.4 11.4 0 0 0 .6 3.6 1 1 0 0 1-.2 1z"/></svg> '
                    f'<span>Call</span></a>'
                )
            else:
                actions = '<span class="no-mobile">No mobile</span>'
            html.append(
                f'<div class="row">'
                f'<div class="cell scholar" data-label="SR No.">{scholar}</div>'
                f'<div class="cell name">{name}</div>'
                f'<div class="cell father" data-label="Father">{father}</div>'
                f'<div class="cell cls" data-label="Class">{cls}</div>'
                f'<div class="cell actions">{actions}</div>'
                f'</div>'
            )
        return "".join(html)

    def render_class_group(key, rows):
        cls, med = key
        return (
            f'<div class="group">'
            f'<div class="group-title"><span class="gt-name">🏫 {cls} ({med})</span>'
            f'<span class="gt-count">{len(rows)}</span></div>'
            f'<div class="list">{render_rows(rows)}</div>'
            f'</div>'
        )

    def render_list_tab(key, label, groups):
        if not groups:
            return f'<section class="tab-panel" id="tab-{key}" data-label="{label}"><div class="empty">✨ No absent students in {label} today.</div></section>'
        inner = "".join(render_class_group(k, rows) for k, rows in groups)
        return f'<section class="tab-panel" id="tab-{key}" data-label="{label}">{inner}</section>'

    def render_report():
        by_class = class_records_map(records)
        total_hm = 0
        total_em = 0
        rows_html = []
        listed = set()
        for cls in CLASS_LIST:
            rows = by_class.get(cls, [])
            hm = sum(1 for r in rows if "Hindi" in str(r.get("Medium") or ""))
            em = sum(1 for r in rows if "English" in str(r.get("Medium") or ""))
            total_hm += hm
            total_em += em
            listed.add(cls)
            rows_html.append(
                f'<div class="rrow"><span class="rcell rcls">{cls}</span>'
                f'<span class="rcell num">{hm}</span><span class="rcell num">{em}</span>'
                f'<span class="rcell num rtot">{hm + em}</span></div>'
            )
        extra_classes = sorted(
            (c for c in by_class if c not in listed),
            key=lambda c: (class_order_index(c), c)
        )
        for cls in extra_classes:
            rows = by_class[cls]
            hm = sum(1 for r in rows if "Hindi" in str(r.get("Medium") or ""))
            em = sum(1 for r in rows if "English" in str(r.get("Medium") or ""))
            total_hm += hm
            total_em += em
            rows_html.append(
                f'<div class="rrow"><span class="rcell rcls">{cls}</span>'
                f'<span class="rcell num">{hm}</span><span class="rcell num">{em}</span>'
                f'<span class="rcell num rtot">{hm + em}</span></div>'
            )
        grand = (
            f'<div class="rrow grand"><span class="rcell rcls">Grand Total</span>'
            f'<span class="rcell num">{total_hm}</span><span class="rcell num">{total_em}</span>'
            f'<span class="rcell num rtot">{total_hm + total_em}</span></div>'
        )
        return (
            f'<section class="tab-panel" id="tab-report" data-label="Attendance Report">'
            f'<div class="report-title">📊 Absent Count Class Wise'
            f'<span class="report-date">{status_text}</span></div>'
            f'<div class="report">'
            f'<div class="rhead"><span>Class</span><span>Hindi Medium</span>'
            f'<span>English Medium</span><span>Total</span></div>'
            f'{"".join(rows_html)}{grand}'
            f'</div></section>'
        )

    total = len(records)
    tab_defs = [
        ("board", "Board Class", render_list_tab("board", "Board Class", board_groups), sum(len(r) for _, r in board_groups)),
        ("local", "Local Class", render_list_tab("local", "Local Class", local_groups), sum(len(r) for _, r in local_groups)),
        ("report", "Attendance Report", render_report(), total),
    ]

    tab_buttons = "".join(
        f'<button class="tab-btn{" active" if i == 0 else ""}" data-tab="tab-{key}">'
        f'{label} <span class="count">{count}</span></button>'
        for i, (key, label, _, count) in enumerate(tab_defs)
    )

    html_doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SSVM Daily Attendance - {status_text}</title>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; -webkit-tap-highlight-color: transparent; }}
html {{ overflow-x: hidden; }}
body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; background: linear-gradient(160deg, #eef2ff 0%, #f8fafc 45%, #ecfeff 100%); color: #0f172a; min-height: 100vh; overflow-x: hidden; }}
.blob {{ position: fixed; border-radius: 50%; filter: blur(90px); opacity: 0.45; z-index: -1; pointer-events: none; }}
.blob1 {{ width: 340px; height: 340px; background: #818cf8; top: -80px; left: -80px; }}
.blob2 {{ width: 300px; height: 300px; background: #38bdf8; top: 30%; right: -100px; }}
.blob3 {{ width: 320px; height: 320px; background: #f0abfc; bottom: -100px; left: 20%; }}
header {{ text-align: center; padding: 30px 16px 18px; position: relative; }}
.logo-wrap {{ display: flex; justify-content: center; margin-bottom: 12px; }}
.logo-img {{ width: 84px; height: 84px; border-radius: 50%; object-fit: cover; border: 3px solid rgba(255,255,255,0.9); box-shadow: 0 8px 24px rgba(30,58,138,0.25); background: #fff; }}
header h1 {{ font-size: 1.5rem; font-weight: 800; letter-spacing: 0.2px; background: linear-gradient(135deg, #1e3a8a, #2563eb, #0891b2); -webkit-background-clip: text; background-clip: text; -webkit-text-fill-color: transparent; line-height: 1.4; }}
.date-pill {{ display: inline-block; margin-top: 12px; background: rgba(255,255,255,0.65); backdrop-filter: blur(12px); -webkit-backdrop-filter: blur(12px); border: 1px solid rgba(255,255,255,0.8); border-radius: 999px; padding: 8px 20px; font-size: 0.95rem; font-weight: 600; color: #334155; box-shadow: 0 4px 16px rgba(59,130,246,0.12); }}
.badge {{ display: inline-block; margin-top: 10px; background: linear-gradient(135deg, #4f46e5, #2563eb); color: #fff; border-radius: 999px; padding: 7px 20px; font-size: 0.85rem; font-weight: 700; box-shadow: 0 6px 18px rgba(79,70,229,0.35); }}
main {{ max-width: 900px; margin: 0 auto; padding: 6px 14px 50px; }}
.tabs {{ display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 14px; }}
.tab-btn {{ border: 1px solid rgba(255,255,255,0.85); background: rgba(255,255,255,0.55); backdrop-filter: blur(14px); -webkit-backdrop-filter: blur(14px); color: #334155; padding: 10px 16px; border-radius: 999px; font-size: 0.9rem; font-weight: 700; cursor: pointer; box-shadow: 0 4px 14px rgba(15,23,42,0.08); transition: all 0.18s ease; }}
.tab-btn:active {{ transform: scale(0.96); }}
.tab-btn.active {{ background: linear-gradient(135deg, #4f46e5, #2563eb); color: #fff; border-color: transparent; box-shadow: 0 6px 20px rgba(79,70,229,0.4); }}
.tab-btn .count {{ display: inline-block; background: rgba(148,163,184,0.25); border-radius: 999px; padding: 1px 8px; font-size: 0.72rem; margin-left: 5px; }}
.tab-btn.active .count {{ background: rgba(255,255,255,0.25); color: #fff; }}
.tab-panel {{ display: none; }}
.tab-panel.active {{ display: block; animation: fadeIn 0.25s ease; }}
@keyframes fadeIn {{ from {{ opacity: 0; transform: translateY(6px); }} to {{ opacity: 1; transform: translateY(0); }} }}
.group {{ margin-bottom: 18px; }}
.group-title {{ display: flex; align-items: center; gap: 10px; margin: 0 4px 8px; }}
.gt-name {{ font-weight: 800; font-size: 1rem; color: #1e3a8a; }}
.gt-count {{ background: rgba(79,70,229,0.12); color: #4338ca; border-radius: 999px; padding: 2px 10px; font-size: 0.75rem; font-weight: 800; }}
.list {{ display: block; background: rgba(255,255,255,0.6); backdrop-filter: blur(18px); -webkit-backdrop-filter: blur(18px); border: 1px solid rgba(255,255,255,0.9); border-radius: 18px; padding: 6px; box-shadow: 0 10px 34px rgba(15,23,42,0.1); }}
.row {{ display: grid; grid-template-columns: 88px minmax(0,1.3fr) minmax(0,1.1fr) 104px 200px; grid-template-areas: "scholar name father cls actions"; column-gap: 8px; align-items: center; padding: 10px 14px; border-radius: 12px; transition: background 0.15s ease; }}
.row + .row {{ border-top: 1px solid rgba(148,163,184,0.18); }}
.row:hover {{ background: rgba(255,255,255,0.75); }}
.cell {{ min-width: 0; }}
.cell.scholar {{ grid-area: scholar; font-weight: 800; color: #4338ca; font-size: 0.9rem; }}
.cell.name {{ grid-area: name; font-weight: 700; font-size: 1rem; color: #0f172a; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
.cell.father {{ grid-area: father; color: #64748b; font-size: 0.9rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
.cell.cls {{ grid-area: cls; color: #2563eb; font-weight: 700; font-size: 0.9rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
.cell.actions {{ grid-area: actions; display: flex; gap: 7px; justify-content: flex-end; }}
.btn {{ display: inline-flex; align-items: center; gap: 6px; padding: 8px 12px; border-radius: 10px; font-size: 0.8rem; font-weight: 700; text-decoration: none; white-space: nowrap; transition: all 0.15s ease; }}
.btn:active {{ transform: scale(0.94); }}
.btn.wa {{ background: rgba(18,140,75,0.12); color: #128c4b; border: 1px solid rgba(18,140,75,0.25); }}
.btn.wa:hover {{ background: #128c4b; color: #fff; }}
.btn.call {{ background: rgba(29,78,216,0.10); color: #1d4ed8; border: 1px solid rgba(29,78,216,0.22); }}
.btn.call:hover {{ background: #1d4ed8; color: #fff; }}
.no-mobile {{ color: #94a3b8; font-size: 0.75rem; }}
.empty {{ background: rgba(255,255,255,0.6); backdrop-filter: blur(18px); -webkit-backdrop-filter: blur(18px); border: 1px solid rgba(255,255,255,0.9); border-radius: 18px; padding: 40px 16px; text-align: center; color: #64748b; font-size: 0.95rem; box-shadow: 0 10px 34px rgba(15,23,42,0.1); }}
.report-title {{ display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 6px; font-weight: 800; font-size: 1.05rem; color: #1e3a8a; margin-bottom: 10px; }}
.report-date {{ font-size: 0.8rem; color: #64748b; font-weight: 600; }}
.report {{ background: rgba(255,255,255,0.6); backdrop-filter: blur(18px); -webkit-backdrop-filter: blur(18px); border: 1px solid rgba(255,255,255,0.9); border-radius: 18px; padding: 8px; box-shadow: 0 10px 34px rgba(15,23,42,0.1); }}
.rhead, .rrow {{ display: grid; grid-template-columns: minmax(0,1fr) 100px 100px 80px; gap: 8px; align-items: center; padding: 10px 14px; }}
.rhead {{ font-size: 0.68rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.6px; color: #64748b; }}
.rhead span:nth-child(n+2), .rrow .num {{ text-align: right; }}
.rrow + .rrow {{ border-top: 1px solid rgba(148,163,184,0.18); }}
.rcls {{ font-weight: 700; font-size: 0.92rem; }}
.num {{ font-weight: 700; font-size: 0.92rem; color: #334155; }}
.rtot {{ color: #4338ca; }}
.rrow.grand {{ background: linear-gradient(135deg, #4f46e5, #2563eb); border-radius: 12px; margin-top: 4px; }}
.rrow.grand .num, .rrow.grand .rcls {{ color: #fff; }}
footer {{ text-align: center; color: #94a3b8; font-size: 0.75rem; padding: 18px 12px 30px; }}
@media (max-width: 680px) {{
  header {{ padding: 24px 12px 14px; }}
  .logo-img {{ width: 64px; height: 64px; }}
  header h1 {{ font-size: 1.1rem; }}
  main {{ padding: 4px 10px 40px; }}
  .row {{ grid-template-columns: minmax(0,1fr) auto; grid-template-areas: "name actions" "scholar scholar" "cls cls"; row-gap: 3px; padding: 13px 12px; }}
  .row + .row {{ border-top: 1px solid rgba(148,163,184,0.18); }}
  .cell.father {{ display: none; }}
  .cell.name {{ font-size: 1.05rem; }}
  .cell.scholar, .cell.cls {{ font-size: 0.9rem; }}
  .cell.scholar::before, .cell.cls::before {{ content: attr(data-label) ": "; color: #94a3b8; font-weight: 600; }}
  .cell.actions {{ gap: 7px; }}
  .btn {{ padding: 9px 11px; font-size: 0.78rem; }}
  .btn svg {{ width: 14px; height: 14px; }}
  .tabs {{ gap: 7px; }}
  .tab-btn {{ padding: 9px 14px; font-size: 0.85rem; }}
  .date-pill {{ font-size: 0.88rem; padding: 7px 16px; }}
  .badge {{ font-size: 0.8rem; }}
  .gt-name {{ font-size: 0.95rem; }}
  .rhead, .rrow {{ grid-template-columns: minmax(0,1fr) 76px 76px 64px; gap: 6px; padding: 11px 10px; }}
  .rhead {{ font-size: 0.6rem; }}
  .report-title {{ font-size: 0.95rem; }}
}}
</style>
</head>
<body>
<div class="blob blob1"></div>
<div class="blob blob2"></div>
<div class="blob blob3"></div>
<header>
<div class="logo-wrap"><img class="logo-img" src="logo.jpg" alt="School Logo"></div>
<h1>Shri Saraswati Sen Sec Vidhya Mandir Mandli Kalyanpur Balotra</h1>
<span class="date-pill">📅 {status_text}</span>
<br>
<span class="badge">Total Absent Today: {total}</span>
</header>
<main>
<nav class="tabs">{tab_buttons}</nav>
{tab_defs[0][2]}
{tab_defs[1][2]}
{tab_defs[2][2]}
</main>
<footer>Updated daily by automated script</footer>
<script>
document.querySelectorAll(".tab-btn").forEach(function(btn) {{
  btn.addEventListener("click", function() {{
    document.querySelectorAll(".tab-btn").forEach(function(b) {{ b.classList.remove("active"); }});
    document.querySelectorAll(".tab-panel").forEach(function(p) {{ p.classList.remove("active"); }});
    btn.classList.add("active");
    document.getElementById(btn.dataset.tab).classList.add("active");
  }});
}});
</script>
</body>
</html>"""
    return html_doc

def github_api(config, method, path, payload=None, retries=3, timeout=90):
    """Performs a GitHub API request with retry logic and returns parsed JSON (or None)."""
    import time
    url = f"https://api.github.com{path}"
    token = config.get("github_token", "")
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github+json",
        "User-Agent": "DailyAttendance-Publisher",
    }
    data = json.dumps(payload).encode("utf-8") if payload is not None else None

    for attempt in range(1, retries + 1):
        req = urllib.request.Request(url, data=data, method=method, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                body = resp.read().decode("utf-8")
                return json.loads(body) if body else None
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace")
            logging.warning(f"GitHub API {method} {path} -> HTTP {e.code}: {body[:300]}")
            if e.code == 404:
                return None
            if attempt == retries:
                return None
        except Exception as e:
            logging.warning(f"GitHub API {method} {path} attempt {attempt}/{retries} failed: {e}")
            if attempt == retries:
                raise
        time.sleep(3 * attempt)
    return None

def enable_github_pages(config):
    """Ensures GitHub Pages is enabled on the target branch (root)."""
    owner = config["github_owner"]
    repo = config["github_repo"]
    branch = config["github_branch"]
    source = {"source": {"branch": branch, "path": "/"}}
    existing = github_api(config, "GET", f"/repos/{owner}/{repo}/pages")
    if existing is not None:
        logging.info(f"GitHub Pages already enabled: {existing.get('html_url', '')}")
        return
    result = github_api(config, "POST", f"/repos/{owner}/{repo}/pages", source)
    if result is not None:
        logging.info(f"GitHub Pages created: {result.get('html_url', result.get('status', ''))}")
    else:
        logging.info("GitHub Pages could not be enabled via API; user may need to enable manually in repo settings.")

def publish_to_github(config, html_doc):
    """Pushes index.html to the repo via GitHub Contents API and enables Pages."""
    owner = config["github_owner"]
    repo = config["github_repo"]
    branch = config["github_branch"]
    token = config.get("github_token", "")
    if not token:
        raise RuntimeError("github_token is missing from config.json")

    path = "index.html"
    content_b64 = base64.b64encode(html_doc.encode("utf-8")).decode("utf-8")

    sha = None
    existing = github_api(config, "GET", f"/repos/{owner}/{repo}/contents/{path}?ref={branch}")
    if existing and "sha" in existing:
        sha = existing["sha"]

    payload = {
        "message": f"Update daily attendance {datetime.now().strftime('%d-%m-%Y %H:%M')}",
        "content": content_b64,
        "branch": branch,
    }
    if sha:
        payload["sha"] = sha

    result = github_api(config, "PUT", f"/repos/{owner}/{repo}/contents/{path}", payload)
    if result is None:
        raise RuntimeError(f"Failed to push {path} to {owner}/{repo}")

    # Keep repo root clean (only .nojekyll + index.html + logo.jpg are required)
    nojekyll = base64.b64encode(b"").decode("utf-8")
    nk_sha = None
    nk_existing = github_api(config, "GET", f"/repos/{owner}/{repo}/contents/.nojekyll?ref={branch}")
    if nk_existing and "sha" in nk_existing:
        nk_sha = nk_existing["sha"]
    nk_payload = {
        "message": "Add .nojekyll",
        "content": nojekyll,
        "branch": branch,
    }
    if nk_sha:
        nk_payload["sha"] = nk_sha
    github_api(config, "PUT", f"/repos/{owner}/{repo}/contents/.nojekyll", nk_payload)

    logo_path = config.get("logo_path", "")
    if logo_path and os.path.exists(logo_path):
        with open(logo_path, "rb") as lf:
            logo_b64 = base64.b64encode(lf.read()).decode("utf-8")
        lg_sha = None
        lg_existing = github_api(config, "GET", f"/repos/{owner}/{repo}/contents/logo.jpg?ref={branch}")
        if lg_existing and "sha" in lg_existing:
            lg_sha = lg_existing["sha"]
        lg_payload = {
            "message": "Update school logo",
            "content": logo_b64,
            "branch": branch,
        }
        if lg_sha:
            lg_payload["sha"] = lg_sha
        github_api(config, "PUT", f"/repos/{owner}/{repo}/contents/logo.jpg", lg_payload)
        logging.info("Uploaded logo.jpg to repo")
    else:
        logging.warning("logo_path not configured or file missing; skipping logo upload")

    enable_github_pages(config)
    logging.info(f"Published to https://{owner}.github.io/{repo}/")
    return f"https://{owner}.github.io/{repo}/"

def run_part_2_github_publish(config, excel_path):
    """
    PART 2: Read combined excel, generate modern HTML page (10th / 12th / Others tabs)
    with WhatsApp & Call buttons, push to GitHub Pages.
    """
    logging.info("========================================================")
    logging.info("PART 2: GITHUB PAGES PUBLISH")
    logging.info("========================================================")

    records = read_records_from_excel(excel_path)
    logging.info(f"Read {len(records)} records from {excel_path}")
    if not records:
        logging.warning("No records found in excel file; still publishing page.")

    status_text = get_formatted_status_text()
    html_doc = build_html(records, status_text)
    local_html = os.path.join(config["export_directory"], "index.html")
    with open(local_html, "w", encoding="utf-8") as f:
        f.write(html_doc)
    logging.info(f"Generated {local_html}")

    if os.environ.get("GITHUB_ACTIONS") == "true":
        site_url = f"https://{config['github_owner']}.github.io/{config['github_repo']}/"
        logging.info("Running in GitHub Actions: index.html generated. Git commit step will push changes.")
        logging.info("--------------------------------------------------------")
        logging.info(f"SUCCESS: Attendance page will be published: {site_url}")
        logging.info("--------------------------------------------------------")
        return site_url

    site_url = publish_to_github(config, html_doc)
    logging.info("--------------------------------------------------------")
    logging.info(f"SUCCESS: Attendance page published: {site_url}")
    logging.info("--------------------------------------------------------")
    return site_url

def show_desktop_notification(title, message):
    """Shows a non-intrusive Windows desktop balloon notification."""
    try:
        import subprocess
        ps_cmd = (
            f"[void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms');"
            f"$n = New-Object System.Windows.Forms.NotifyIcon;"
            f"$n.Icon = [System.Drawing.SystemIcons]::Information;"
            f"$n.BalloonTipTitle = '{title}';"
            f"$n.BalloonTipText = '{message}';"
            f"$n.Visible = $True;"
            f"$n.ShowBalloonTip(5000);"
            f"Start-Sleep -Milliseconds 600;"
            f"$n.Dispose();"
        )
        subprocess.run(["powershell", "-NoProfile", "-Command", ps_cmd], timeout=10, capture_output=True)
    except Exception:
        pass

def main():
    config = load_config()
    with sync_playwright() as p:
        # Part 1: ERP Extraction & Excel Generation
        excel_path = run_part_1_erp_extraction(config, p)

    # Part 2: GitHub Pages Publishing
    site_url = run_part_2_github_publish(config, excel_path)

    logging.info("========================================================")
    logging.info("PART 1 + PART 2 COMPLETED SUCCESSFULLY!")
    logging.info("========================================================")

    # Show desktop notification so user sees when automated run completes
    show_desktop_notification(
        "SSVM Daily Attendance",
        "Attendance successfully scraped and published to GitHub Pages!"
    )

if __name__ == "__main__":
    try:
        main()
        print("\n========================================================")
        print("SUCCESS: Part 1 (ERP Data Scraping) + Part 2 (GitHub Pages) Completed!")
        print("========================================================")
    except Exception as err:
        logging.error(f"Execution failed: {err}", exc_info=True)
        sys.exit(1)
