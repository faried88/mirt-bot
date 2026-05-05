import logging
import os
import math
import time
import tempfile
import re
from datetime import datetime
from io import BytesIO

import requests
import pandas as pd
import numpy as np
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ContextTypes, ConversationHandler, filters
)

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ── Config ───────────────────────────────────────────────────────────────────
TOKEN         = "8687940667:AAE2jfP57jzaRu5LE0BneU3IlXCcrTwJIeQ"
ALLOWED_USER  = 910169628
DRIVE_FOLDER  = "1RmxvQKfY8e9-BL3gphaVNwRApNQa2YOp"
BASE_URL      = "https://sps-bs.cigalah.com.sa"
LOGIN_URL     = f"{BASE_URL}/Report/Identity/login"
PIVOT_URL     = f"{BASE_URL}/Report/DevPivotViewer.aspx"
MiRT_URL      = f"{BASE_URL}/Report/MiRT/index.html"
MIRT_USER     = "M.faried88"
MIRT_PASS     = "handle"

SERVICE_ACCOUNT_INFO = {
    "type": "service_account",
    "project_id": "mirt-495420",
    "private_key_id": "eea07b7b52c951a0003fa1664127ea81f271dd4f",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQCe5s79PXfwcCyK\n+z0ljxAMwDTRfcEZO3CqF7RA057THK1k72veCpIUkmle1oCcCFe6LTg0+9pVcGae\nhKAmK0e8YwBu6W/UEwAFyYB0ickvr5hcwzg870V0Lnk4dwupA34mlBqhx6JCaY4C\nR3PJ1vzVGwoUbhfU5M8Be7TJwor1DVrNugHZm7zVGnnrtJoyq0a8CZLD3nOHdzul\ngj4o8xnOTBZVN55eZmoVHnHwmO8wDG53cpnX2keESsZTV/DdLAUp5JkY/f7HLBjg\njc4s5cAt8giAC1QREiSN4pXhCCe97fecbNaSX056u46ByOnbzkkJV2uPRsAtrprU\nK4TJGHTvAgMBAAECggEAOIArY4JizlykifqRoHRBKbeCUGcdrSIkimaJUm+szrYo\ntXYobbhmfugcjXtKGbEhuHJxxO00kiK4am8QHuJOzJ6LPeTFPaxP2r7ubQG9RrZy\nP7GuooQVtxz7P2ec/sjeJ0uMOLAqcuDjfM35TvCh0AigSelnkeyV6poZC5CgJke8\n7cS0gXsUjeZz7oZra5PQlv54ymXxHx0vXhH6MAOhCV2Qks3BUdKeyy3sL/YZ48HV\nVETfxCRRKrGnFtjg1i/h10ntev7qDBW0FdF0N8y+Ju4Pgm7tGOnvc7HnTM6pESF6\nEoRfzNJtdihdygMQRHT17qIYqAb+LPZxaxswNfZTsQKBgQDKw3MM+XFQSSg3SGHH\nh8fqGXCYxeGT7vfBKenL1gWT0eWr5CfONllOC9uHA/ChEci9D0Y6O5uWDQK+YdKu\ndBhOJww2S8V6KWkhmfse/jGAa9zgGAnYeVPufE4nK6QXRPYagLChbp2hHJt5PTlF\n7UIqIMo9l8TWk9QXuKj7crQUJwKBgQDInzm3O5flbwKIMFI9xGqVR5G9Q7pfB4FU\nah0KkGJHaVEX26cKgZXdJSsQTJIMu5rUqLCvWv0mKsHEHLd6g4XP6fF1NuvCqd7e\nJUfFuOpuXPT0Ml5nswrJ392WG0MFTWdmPjbkev3XxB5ZEUU7J8fHpHNklh99QqoH\nK/AX7iot+QKBgBsWPyFllVinXUL9XWqdXfyNB3ixPrBXhSt94OjFH5uet7Ld2N94\nbTe658nCofuyd4GiL7yJyAAkntA2G0II6lJObxg1yRzHuW6utlhulshUIH6jV3Ve\nx/KdEoezEcm2AbaKqI34TACA5NgucJ9B0cv082+E/du4heXhWlm0+g+TAoGBAIOd\nyiyGoSE5Ec0s/ldda5shx+AF9dfwQY2SzBipHoDA/B2N0emXmCzr/HOF+G74CRyo\nyrlQFTIb7ODvAgQTEw+S6ADBFiywavEMPijeJpZez6kA/mRD1rkX7/RRUEfDPymZ\neUOt2KjcFhjStruXXn6ASd/ciS4RNSDdV3crnWppAoGBAMawFgwhlAylTCUs6YTH\nrRRrYtb+3dD1NX9QqlpK47oy1tq65U/8JViQ3UaA08KTwqNnXSNaxx6DjeuGk3l4\n+XAfV50ubEUW109RRChAXkrUXmm15A5wBY1TwfuivK76BKpl+rRE1fVCYqmf0CBH\nPfmpTBj1LBBlqmE1ekE9cp1X\n-----END PRIVATE KEY-----\n",
    "client_email": "mirt-bot@mirt-495420.iam.gserviceaccount.com",
    "client_id": "115674558839677734201",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/mirt-bot%40mirt-495420.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}

logging.basicConfig(format="%(asctime)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

WAIT_DATE = 0

# ════════════════════════════════════════════════════════════════════════════
# GOOGLE DRIVE
# ════════════════════════════════════════════════════════════════════════════

def get_drive_service():
    creds = Credentials.from_service_account_info(
        SERVICE_ACCOUNT_INFO,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)

def upload_to_drive(file_bytes, filename, folder_id):
    service = get_drive_service()
    meta = {"name": filename, "parents": [folder_id]}
    media = MediaIoBaseUpload(
        BytesIO(file_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True
    )
    f = service.files().create(body=meta, media_body=media, fields="id,webViewLink").execute()
    return f.get("webViewLink", "")

def create_drive_subfolder(name, parent_id):
    service = get_drive_service()
    meta = {"name": name, "mimeType": "application/vnd.google-apps.folder", "parents": [parent_id]}
    f = service.files().create(body=meta, fields="id").execute()
    return f.get("id")

# ════════════════════════════════════════════════════════════════════════════
# MIRT DOWNLOADER — HTTP requests (no Selenium)
# ════════════════════════════════════════════════════════════════════════════

BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36",
    "Accept-Language": "en-US,en;q=0.9,ar;q=0.8",
    "Accept-Encoding": "gzip, deflate, br, zstd",
    "Connection": "keep-alive",
    "Sec-Ch-Ua": '"Google Chrome";v="147", "Not.A/Brand";v="8", "Chromium";v="147"',
    "Sec-Ch-Ua-Mobile": "?0",
    "Sec-Ch-Ua-Platform": '"Windows"',
}

REPORT_NAMES = {
    "attendance": "Attendance Report",
    "f2f":        "Master F2F",
    "market":     "Market Intelligence Pivot",
    "sv":         "S.V Visit",
    "oos":        "OOS Report",
    "callcenter": "Call Center",
}

def mirt_login():
    s = requests.Session()
    # Step 1: visit main page to get session cookie
    s.get(MiRT_URL, headers={**BROWSER_HEADERS, "Accept": "text/html,application/xhtml+xml"}, timeout=30)
    # Step 2: login with JSON
    r = s.post(LOGIN_URL, json={
        "stayLoggedIn": False,
        "userName": MIRT_USER,
        "password": MIRT_PASS,
    }, headers={
        **BROWSER_HEADERS,
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Origin": BASE_URL,
        "Referer": MiRT_URL,
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-origin",
        "X-Requested-With": "XMLHttpRequest",
    }, timeout=30)
    if r.status_code != 200:
        raise Exception(f"Login failed: {r.status_code} — {r.text[:200]}")
    return s

def mirt_get_pivot_page(s):
    r = s.get(PIVOT_URL, headers={
        **BROWSER_HEADERS,
        "Accept": "text/html,application/xhtml+xml",
        "Referer": MiRT_URL,
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
    }, timeout=30)
    vs  = re.search(r'id="__VIEWSTATE"\s+value="([^"]*)"', r.text)
    vsg = re.search(r'id="__VIEWSTATEGENERATOR"\s+value="([^"]*)"', r.text)
    dpv = re.search(r'id="DevPivotDocumentViewer"\s+value="([^"]*)"', r.text)
    hm  = re.search(r'id="DevPivotDocumentViewer\$HM"\s+value="([^"]*)"', r.text)
    fvm = re.search(r'id="DevPivotDocumentViewer\$FVM"\s+value="([^"]*)"', r.text)
    fm  = re.search(r'id="DevPivotDocumentViewer\$FM"\s+value="([^"]*)"', r.text)
    return {
        "__VIEWSTATE":          vs.group(1)  if vs  else "",
        "__VIEWSTATEGENERATOR": vsg.group(1) if vsg else "367F0B34",
        "DevPivotDocumentViewer":      dpv.group(1) if dpv else "",
        "DevPivotDocumentViewer$HM":   hm.group(1)  if hm  else "",
        "DevPivotDocumentViewer$FVM":  fvm.group(1) if fvm else "",
        "DevPivotDocumentViewer$FM":   fm.group(1)  if fm  else "",
    }

def mirt_load_report(s, report_key, date_str):
    """Click the report in the menu via AJAX"""
    report_name = REPORT_NAMES.get(report_key, "")
    # MiRT uses callback to load reports
    params = {
        "report": report_name,
        "date":   date_str,
        "_dc":    str(int(time.time() * 1000)),
    }
    s.get(f"{BASE_URL}/Report/MiRT/getVersion", params=params, headers={
        **BROWSER_HEADERS,
        "Accept": "application/json",
        "Referer": MiRT_URL,
        "X-Requested-With": "XMLHttpRequest",
    }, timeout=30)
    time.sleep(3)

def mirt_export_excel(s, vs_data, date_str):
    """POST to DevPivotViewer to export Excel"""
    data = {
        "__EVENTTARGET":   "",
        "__EVENTARGUMENT": "",
        **vs_data,
        "configField":  '{"data":"12|#|type|4|6|1excel#"}',
        "configButton": "Config",
        "DXScript": "1_9,1_10,1_253,1_21,1_62,1_11,1_12,1_13,1_18,1_17,7_7,7_5,7_6,7_9,7_4,1_15,1_22,1_31,1_32,1_39,1_180,1_181,1_186,1_30",
        "DXCss": "7_13,1_66,1_67,7_11,7_12,1_72,1_71,1_205,1_207,1_204",
    }
    r = s.post(PIVOT_URL, data=data, headers={
        **BROWSER_HEADERS,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Content-Type": "application/x-www-form-urlencoded",
        "Origin": BASE_URL,
        "Referer": PIVOT_URL,
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
    }, timeout=120)
    ct = r.headers.get("Content-Type", "")
    if "excel" in ct.lower() or "octet" in ct.lower() or (
        r.status_code == 200 and len(r.content) > 5000 and b"PK" in r.content[:10]
    ):
        return r.content
    return None

def download_report_from_mirt(report_key, date_str):
    s = mirt_login()
    vs = mirt_get_pivot_page(s)
    mirt_load_report(s, report_key, date_str)
    # Refresh viewstate after loading report
    vs2 = mirt_get_pivot_page(s)
    content = mirt_export_excel(s, vs2, date_str)
    if not content:
        # Try with original viewstate
        content = mirt_export_excel(s, vs, date_str)
    return content

# ════════════════════════════════════════════════════════════════════════════
# EXCEL HELPERS
# ════════════════════════════════════════════════════════════════════════════

def set_header_style(ws, header_row=1, bg="1F4E79"):
    fill = PatternFill("solid", fgColor=bg)
    font = Font(bold=True, color="FFFFFF")
    for cell in ws[header_row]:
        cell.fill = fill; cell.font = font
        cell.alignment = Alignment(horizontal="center")

def auto_col_width(ws):
    for col in ws.columns:
        ml = 0
        cl = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value: ml = max(ml, len(str(cell.value)))
            except: pass
        ws.column_dimensions[cl].width = min(ml + 4, 40)

def col_fmt_date(ws, cl, start_row=2):
    for cell in ws[cl][start_row-1:]: cell.number_format = "DD-MM-YY"

def col_fmt_number(ws, cl, start_row=2):
    for cell in ws[cl][start_row-1:]:
        if cell.value not in (None,""):
            try: cell.value = float(cell.value); cell.number_format = "0"
            except: pass

def col_fmt_integer(ws, cl, start_row=2):
    for cell in ws[cl][start_row-1:]:
        if cell.value not in (None,""):
            try: cell.value = int(float(str(cell.value).replace(" ",""))); cell.number_format = "0"
            except: pass

def df_to_sheet(wb, df, sheet_name):
    if sheet_name in wb.sheetnames: del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    for c, col in enumerate(df.columns, 1): ws.cell(1, c, col)
    for r, row in enumerate(df.itertuples(index=False), 2):
        for c, val in enumerate(row, 1): ws.cell(r, c, val)
    set_header_style(ws); auto_col_width(ws)
    return ws

# ════════════════════════════════════════════════════════════════════════════
# PROCESSORS
# ════════════════════════════════════════════════════════════════════════════

def process_attendance(fb, day):
    df = pd.read_excel(BytesIO(fb), header=4)
    df.columns = df.columns.str.strip()
    df = df[df["City"].notna() & (df["City"].astype(str).str.strip() != "") & (df["City"].astype(str).str.strip() != "City")].reset_index(drop=True)
    df["Project Type"] = df.groupby("Tablet #")["Project Type"].transform(lambda x: x.ffill().bfill())
    non_agg = {c:(c,"first") for c in df.columns if c not in ("Tablet #","Check In","Check Out")}
    agg = df.groupby("Tablet #", as_index=False).agg(**non_agg, **{"Check In":("Check In","min"),"Check Out":("Check Out","max")})
    try: agg = agg[df.columns]
    except: pass
    agg = agg.sort_values("Check In").drop_duplicates("Tablet #", keep="last").reset_index(drop=True)
    wb = load_workbook(BytesIO(fb))
    for s in wb.sheetnames: del wb[s]
    ws = df_to_sheet(wb, agg, "Detailed Attendance")
    cols = list(agg.columns)
    if "Date"        in cols: col_fmt_date(ws,    get_column_letter(cols.index("Date")+1))
    if "Outlet Code" in cols: col_fmt_number(ws,  get_column_letter(cols.index("Outlet Code")+1))
    if "Mobile"      in cols: col_fmt_integer(ws, get_column_letter(cols.index("Mobile")+1))
    out = BytesIO(); wb.save(out); return out.getvalue()

def process_callcenter(fb, day):
    df = pd.read_excel(BytesIO(fb), header=8).dropna(axis=1,how="all").dropna(how="all")
    df.columns = df.columns.str.strip()
    wb = load_workbook(BytesIO(fb))
    for s in wb.sheetnames: del wb[s]
    df_to_sheet(wb, df, "Callcenter Report")
    out = BytesIO(); wb.save(out); return out.getvalue()

def _f2f_qr(row):
    p = str(row.get("Project","")).strip()
    op = str(row.get("Our Products","")).strip()
    if p == "Davidoff":
        if any(k in op for k in ["DVD Gold","DVD Slim Gold","DVD Slim White","DVD White"]): return "DVD PL"
        if "Evolve" in op: return "Evolve"
    return p

def process_f2f(fb, day):
    df = pd.read_excel(BytesIO(fb), header=0)
    df.columns = df.columns.str.strip()
    for col in df.columns:
        df = df[~(df[col].astype(str).str.strip().str.lower() == "grand total")]
    df = df.reset_index(drop=True)
    if "Status" in df.columns and "Mobile" in df.columns:
        df.loc[df["Status"].astype(str).str.strip() == "Refused to flip","Mobile"] = "0000000000"
    df["Project QR"] = df.apply(_f2f_qr, axis=1)
    wb = load_workbook(BytesIO(fb))
    for s in wb.sheetnames: del wb[s]
    ws_n = df_to_sheet(wb, df, "Numbers")
    if "Mobile" in df.columns and "Status" in df.columns:
        al = df[df["Status"].astype(str).str.strip().isin(["Agreed","Loyal"])]
        dv = set(al[al.duplicated("Mobile",keep=False)]["Mobile"].astype(str))
        mi = list(df.columns).index("Mobile")+1
        yf = PatternFill("solid",fgColor="FFFF00")
        for r in range(2, ws_n.max_row+1):
            if str(ws_n.cell(r,mi).value) in dv: ws_n.cell(r,mi).fill = yf
    df_to_sheet(wb, df, "Main")
    if "QR Code" in df.columns:
        hq = df[df["QR Code"].notna() & (df["QR Code"].astype(str).str.strip()!="")]
        nq = df[df["QR Code"].isna()  | (df["QR Code"].astype(str).str.strip()=="")]
        qdf = pd.concat([hq,nq],ignore_index=True)
        ws_q = df_to_sheet(wb, qdf, "QR")
        qi = list(qdf.columns).index("QR Code")+1
        al2 = qdf[qdf["Status"].astype(str).str.strip().isin(["Agreed","Loyal"])] if "Status" in qdf.columns else qdf
        dq = set(al2[al2.duplicated("QR Code",keep=False)]["QR Code"].astype(str)) if not al2.empty else set()
        for r in range(2, ws_q.max_row+1):
            c = ws_q.cell(r,qi)
            if c.value in (None,""): c.fill = PatternFill("solid",fgColor="FFFF00")
            elif str(c.value) in dq: c.fill = PatternFill("solid",fgColor="FFA500")
    out = BytesIO(); wb.save(out); return out.getvalue()

def _strip_ms(fb):
    dr = pd.read_excel(BytesIO(fb), header=None)
    ms = next((i for i,row in dr.iterrows() if row.astype(str).str.upper().str.contains("MS CHOICE").any()), None)
    if ms is not None: dr = dr.drop(index=ms).reset_index(drop=True)
    df = dr.copy(); df.columns = df.iloc[0].astype(str).str.strip()
    df = df.iloc[1:].reset_index(drop=True)
    return df.loc[:, df.columns.notna() & (df.columns!="nan")]

def _pt(row):
    pr = str(row.get("Project","")).strip().upper()
    pq = str(row.get("Project Type?","")).strip()
    pt = str(row.get("Project Type","")).strip()
    pqv = pq if "GITANES" in pr else ("DVD" if "DAVIDOFF" in pr else pq)
    ptf = (pqv if not pt or pt=="nan" else ("Davidoff" if pt.upper()=="DVD" else pt))
    return pqv, ptf

def process_market(fb, day):
    df = _strip_ms(fb)
    if "Outlet ID" in df.columns:
        df = df[df["Outlet ID"].notna() & (df["Outlet ID"].astype(str).str.strip()!="")]
    if "Date" in df.columns: df["Date"] = df["Date"].ffill().bfill()
    if "MS Report Time" in df.columns: df["MS Report Time"] = df["MS Report Time"].ffill().bfill()
    if "Project Type?" in df.columns and "Project Type" in df.columns:
        r = df.apply(_pt,axis=1); df["Project Type?"]=[x[0] for x in r]; df["Project Type"]=[x[1] for x in r]
    cols = list(df.columns)
    ordered = [c for c in ["Date","MS Report Time"] if c in cols]
    df = df[ordered + [c for c in cols if c not in ordered]]
    wb = load_workbook(BytesIO(fb))
    for s in wb.sheetnames: del wb[s]
    ws = df_to_sheet(wb, df, "Market Report")
    cols = list(df.columns)
    if "Date"        in cols: col_fmt_date(ws,    get_column_letter(cols.index("Date")+1))
    if "Outlet Code" in cols: col_fmt_number(ws,  get_column_letter(cols.index("Outlet Code")+1))
    if "Mobile"      in cols: col_fmt_integer(ws, get_column_letter(cols.index("Mobile")+1))
    out = BytesIO(); wb.save(out); return out.getvalue()

FS = {"Ihab Kader","Mohammed Faried","Ahmad Gamal","Ibrahim Fawaz"}
CT = {"Mohammed Izzat","Mohamed Abdul Aziz"}

def process_sv(fb, day):
    df = _strip_ms(fb)
    if "Date" in df.columns: df["Date"] = df["Date"].ffill().bfill()
    if "MS Report Time" in df.columns: df["MS Report Time"] = df["MS Report Time"].ffill().bfill()
    if "Project Type?" in df.columns and "Project Type" in df.columns:
        r = df.apply(_pt,axis=1); df["Project Type?"]=[x[0] for x in r]; df["Project Type"]=[x[1] for x in r]
    sv_col  = next((c for c in df.columns if "s.v" in c.lower() or "coacher" in c.lower()), None)
    sup_col = next((c for c in df.columns if "supervisor" in c.lower()), None)
    if sv_col:
        def cl(n):
            n=str(n).strip()
            return ("Field Supervisor",n) if n in FS else (("Control",n) if n in CT else ("S.V",n))
        r2=df[sv_col].apply(cl); df["Type"]=[x[0] for x in r2]; df["Coacher Name"]=[x[1] for x in r2]
    wb = load_workbook(BytesIO(fb))
    for s in wb.sheetnames: del wb[s]
    ws = df_to_sheet(wb, df, "S.V Report")
    if sup_col and "Coacher Name" in list(df.columns):
        cc=ws.max_column+1
        hc=ws.cell(1,cc,"Compare"); hc.fill=PatternFill("solid",fgColor="1F4E79"); hc.font=Font(bold=True,color="FFFFFF")
        sl=get_column_letter(list(df.columns).index(sup_col)+1)
        cll=get_column_letter(list(df.columns).index("Coacher Name")+1)
        for r in range(2,ws.max_row+1): ws.cell(r,cc,f"={sl}{r}={cll}{r}")
    if "Date" in list(df.columns): col_fmt_date(ws, get_column_letter(list(df.columns).index("Date")+1))
    out = BytesIO(); wb.save(out); return out.getvalue()

def process_oos(fb, day):
    df = _strip_ms(fb)
    for col in ["MS Report Time","Promoter ID"]:
        if col in df.columns: df=df.drop(columns=[col])
    if "Outlet ID" in df.columns:
        df=df[df["Outlet ID"].notna() & (df["Outlet ID"].astype(str).str.strip()!="")]
    if "Date" in df.columns: df["Date"]=df["Date"].ffill().bfill()
    if "Project Type?" in df.columns and "Project Type" in df.columns:
        r=df.apply(_pt,axis=1); df["Project Type?"]=[x[0] for x in r]; df["Project Type"]=[x[1] for x in r]
    wb = load_workbook(BytesIO(fb))
    for s in wb.sheetnames: del wb[s]
    ws = df_to_sheet(wb, df, "OOS Report")
    cols = list(df.columns)
    if "Date"        in cols: col_fmt_date(ws, get_column_letter(cols.index("Date")+1))
    if "Outlet Code" in cols: col_fmt_number(ws, get_column_letter(cols.index("Outlet Code")+1))
    for col in cols:
        if any(k in col.lower() for k in ["qty","quantity","count"]):
            col_fmt_number(ws, get_column_letter(cols.index(col)+1))
    out = BytesIO(); wb.save(out); return out.getvalue()

PROCESSORS = {
    "attendance": (process_attendance, "attendance report"),
    "callcenter": (process_callcenter, "Callcenter Report"),
    "f2f":        (process_f2f,        "f2f report"),
    "market":     (process_market,     "Market Report"),
    "sv":         (process_sv,         "S.V Report"),
    "oos":        (process_oos,        "OOS Report"),
}

REPORT_LABELS = {
    "attendance": "📋 Attendance",
    "callcenter": "📞 Callcenter",
    "f2f":        "🤝 F2F",
    "market":     "📊 Market",
    "sv":         "👔 S.V",
    "oos":        "📦 OOS",
}

# ════════════════════════════════════════════════════════════════════════════
# BOT HANDLERS
# ════════════════════════════════════════════════════════════════════════════

def is_allowed(update):
    return update.effective_user.id == ALLOWED_USER

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update): return ConversationHandler.END
    await update.message.reply_text(
        "👋 *أهلاً يا محمد!*\n\n"
        "📅 اكتب التاريخ:\n`YYYY-MM-DD`\nمثال: `2026-05-05`",
        parse_mode="Markdown"
    )
    return WAIT_DATE

async def handle_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_allowed(update): return ConversationHandler.END
    date_text = update.message.text.strip()
    try:
        date_obj = datetime.strptime(date_text, "%Y-%m-%d")
    except ValueError:
        await update.message.reply_text("⚠️ صيغة غلط!\nاكتب زي كده: `2026-05-05`", parse_mode="Markdown")
        return WAIT_DATE

    day          = str(date_obj.day)
    date_display = date_obj.strftime("%d/%m/%Y")

    msg = await update.message.reply_text(
        f"✅ *{date_display}*\n\n⏳ بدأت تنزيل الـ 6 تقارير...", parse_mode="Markdown"
    )

    try:
        sub_folder_id = create_drive_subfolder(f"تقارير {date_display}", DRIVE_FOLDER)
    except:
        sub_folder_id = DRIVE_FOLDER

    results = {}
    done = 0

    for key, (processor, fname_prefix) in PROCESSORS.items():
        label = REPORT_LABELS[key]
        await msg.edit_text(
            f"🔄 *{label}* — جاري التنزيل...\n({done}/{len(PROCESSORS)} خلصوا)",
            parse_mode="Markdown"
        )
        try:
            file_bytes = download_report_from_mirt(key, date_text)
            if not file_bytes:
                raise Exception("الملف مش اتنزل")

            await msg.edit_text(
                f"⚙️ *{label}* — جاري المعالجة...\n({done}/{len(PROCESSORS)} خلصوا)",
                parse_mode="Markdown"
            )
            result_bytes = processor(file_bytes, day)
            out_name     = f"{fname_prefix} {day}.xlsx"
            drive_link   = upload_to_drive(result_bytes, out_name, sub_folder_id)
            results[key] = ("✅", label, result_bytes, out_name)
        except Exception as e:
            logger.exception(f"Error on {key}")
            results[key] = ("❌", label, None, str(e)[:80])
        done += 1

    # Summary
    lines = [f"📊 *تقارير {date_display}*\n"]
    all_ok = True
    for key, data in results.items():
        if data[0] == "✅":
            lines.append(f"✅ {data[1]}")
        else:
            lines.append(f"❌ {data[1]}\n   _{data[3]}_")
            all_ok = False

    folder_link = f"https://drive.google.com/drive/folders/{sub_folder_id}"
    lines.append(f"\n📁 [افتح الفولدر]({folder_link})")
    await msg.edit_text("\n".join(lines), parse_mode="Markdown")

    # Send files on Telegram
    for key, data in results.items():
        if data[0] == "✅" and data[2]:
            await update.message.reply_document(
                document=BytesIO(data[2]),
                filename=data[3],
                caption=f"📎 {data[1]}"
            )

    keyboard = [
        [InlineKeyboardButton("📅 تاريخ تاني", callback_data="again")],
        [InlineKeyboardButton("🚪 خروج", callback_data="exit")]
    ]
    await update.message.reply_text("عايز تاريخ تاني؟", reply_markup=InlineKeyboardMarkup(keyboard))
    return WAIT_DATE

async def callback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "again":
        await query.edit_message_text("📅 اكتب التاريخ:\n`YYYY-MM-DD`", parse_mode="Markdown")
        return WAIT_DATE
    await query.edit_message_text("👋 لما تحتاجني ابعت /start")
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("❌ تم الإلغاء.")
    return ConversationHandler.END

def main():
    app = Application.builder().token(TOKEN).build()
    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={WAIT_DATE: [
            MessageHandler(filters.TEXT & ~filters.COMMAND, handle_date),
            CallbackQueryHandler(callback_handler),
        ]},
        fallbacks=[CommandHandler("cancel", cancel)],
        per_message=False,
    )
    app.add_handler(conv)
    logger.info("✅ MiRT Bot started!")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
