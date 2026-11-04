import io
import os
import uuid
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Mall Vehicle Counter", page_icon="🚗", layout="centered")

TZ = ZoneInfo("Asia/Bangkok")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
FONT_PATH = "fonts/Sarabun-Regular.ttf"  # ฟอนต์ไทยสำหรับ PDF
REC = "VehicleRecored"  # สะกดตามที่กำหนดไว้

H = {
    "Users": ["uid", "email", "display_name", "role", "status", "created_at", "last_login"],
    "Branches": ["branch_id", "branch_name", "status"],
    "Gates": ["gate_id", "branch_id", "gate_name", "status"],
    REC: ["record_id", "entry_date", "branch_id", "gate_id", "vehicle_count", "note",
          "created_by_uid", "created_by_email", "created_at", "updated_by", "updated_at", "is_deleted"],
    "AuditLogs": ["log_id", "record_id", "action", "old_value", "new_value", "changed_by", "changed_at"],
    "UserBranchAccess": ["uid", "branch_id", "status"],
    "Settings": ["setting_key", "setting_value"],
}


def now():
    return datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")


def today():
    return datetime.now(TZ).date()


def not_deleted(df):
    return df[df.is_deleted.astype(str).str.upper() != "TRUE"]


# ---------------------------------------------------------------- Google Sheets
@st.cache_resource
def get_book():
    creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=SCOPES)
    book = gspread.authorize(creds).open_by_key(st.secrets["sheet_id"])
    have = {w.title: w for w in book.worksheets()}
    for name, cols in H.items():
        ws = have.get(name)
        if ws is None:
            ws = book.add_worksheet(title=name, rows=1000, cols=len(cols))
        if not ws.row_values(1):
            ws.update(values=[cols], range_name="A1")
            if name == REC:  # เก็บวันที่เป็นข้อความ กันชีตแปลงรูปแบบเอง
                ws.format("B:B", {"numberFormat": {"type": "TEXT"}})
    for t in ("Sheet1", "ชีต1"):  # ลบชีตเปล่าที่ Google สร้างให้
        w = have.get(t)
        if w is not None and not w.get_all_values():
            book.del_worksheet(w)
    if len(book.worksheet("Branches").get_all_values()) < 2:
        book.worksheet("Branches").append_rows(
            [["B01", "ลาดกระบัง", "active"], ["B02", "รามคำแหง", "active"], ["B03", "กาญจนาภิเษก", "active"]])
        book.worksheet("Gates").append_rows(
            [["B01-G1", "B01", "ประตู 1", "active"],
             ["B02-G1", "B02", "ประตู 1", "active"], ["B02-G2", "B02", "ประตู 2", "active"],
             ["B03-G1", "B03", "ประตู 1", "active"], ["B03-G2", "B03", "ประตู 2", "active"],
             ["B03-G3", "B03", "ประตู 3", "active"]])
    admin = st.secrets.get("bootstrap_admin_email", "")
    if admin and len(book.worksheet("Users").get_all_values()) < 2:
        book.worksheet("Users").append_row(
            [uuid.uuid4().hex[:10], admin.lower(), "Admin", "admin", "active", now(), ""])
    return book


@st.cache_data(ttl=30, show_spinner=False)
def read(name):
    rows = get_book().worksheet(name).get_all_records(numericise_ignore=["all"])
    return pd.DataFrame(rows, columns=H[name]).fillna("")


def fresh(name):
    read.clear()
    return read(name)


def append(name, rows):
    get_book().worksheet(name).append_rows(rows, value_input_option="RAW")
    read.clear()


def rewrite(name, df):
    ws = get_book().worksheet(name)
    ws.clear()
    ws.update(values=[H[name]] + df[H[name]].astype(object).values.tolist(), range_name="A1")
    read.clear()


def audit(record_id, action, old, new, by):
    append("AuditLogs", [[uuid.uuid4().hex[:12], record_id, action, old, new, by, now()]])


# ---------------------------------------------------------------- Auth
def login():
    if not st.user.is_logged_in:
        st.title("🚗 Mall Vehicle Counter")
        st.write("เข้าสู่ระบบด้วยบัญชี Google ที่ผู้ดูแลระบบอนุญาตไว้")
        st.button("เข้าสู่ระบบด้วย Google", on_click=st.login, type="primary")
        st.stop()
    email = (st.user.email or "").lower()
    users = read("Users")
    m = users[(users.email.str.lower() == email) & (users.status == "active")]
    if m.empty:
        st.error(f"อีเมล {email} ยังไม่ได้รับสิทธิ์ใช้งาน กรุณาติดต่อผู้ดูแลระบบ")
        st.button("ออกจากระบบ", on_click=st.logout)
        st.stop()
    u = m.iloc[0].to_dict()
    if not st.session_state.get("logged"):
        get_book().worksheet("Users").update_cell(int(m.index[0]) + 2, 7, now())
        st.session_state.logged = True
        read.clear()
    return u


def allowed_branches(u):
    br = read("Branches")
    br = br[br.status == "active"]
    if u["role"] == "admin":
        return br
    acc = read("UserBranchAccess")
    ids = acc[(acc.uid == u["uid"]) & (acc.status == "active")].branch_id
    return br[br.branch_id.isin(ids)]


# ---------------------------------------------------------------- Save logic
def save_records(u, ds, items):
    """items = [(branch_id, gate_id, count, note)] -> upsert ด้วย gate_id + entry_date"""
    ws = get_book().worksheet(REC)
    cur = fresh(REC)
    new_rows, updates, logs = [], [], []
    for b, g, c, n in items:
        hit = cur[(cur.gate_id == g) & (cur.entry_date == ds)]
        if hit.empty:
            rid = uuid.uuid4().hex[:12]
            new_rows.append([rid, ds, b, g, int(c), n, u["uid"], u["email"], now(), "", "", "FALSE"])
            logs.append([uuid.uuid4().hex[:12], rid, "create", "", f"count={int(c)};note={n}", u["email"], now()])
        else:
            r = hit.iloc[0]
            changed = (str(r.vehicle_count) != str(int(c)) or r.note != n
                       or str(r.is_deleted).upper() == "TRUE")
            if not changed:
                continue
            row = int(hit.index[0]) + 2
            updates.append({"range": f"E{row}:F{row}", "values": [[int(c), n]]})
            updates.append({"range": f"J{row}:L{row}", "values": [[u["email"], now(), "FALSE"]]})
            logs.append([uuid.uuid4().hex[:12], r.record_id, "update",
                         f"count={r.vehicle_count};note={r.note}", f"count={int(c)};note={n}",
                         u["email"], now()])
    if new_rows:
        ws.append_rows(new_rows, value_input_option="RAW")
    if updates:
        ws.batch_update(updates, value_input_option="RAW")
    if logs:
        get_book().worksheet("AuditLogs").append_rows(logs, value_input_option="RAW")
    read.clear()
    return len(new_rows), len(logs) - len(new_rows)


# ---------------------------------------------------------------- Pages
def page_entry(u):
    st.header("📝 บันทึกจำนวนรถเข้า")
    br = allowed_branches(u)
    if br.empty:
        st.warning("ยังไม่ได้รับสิทธิ์สาขาใด กรุณาติดต่อผู้ดูแลระบบ")
        return
    gates = read("Gates")
    gates = gates[gates.status == "active"]
    ds = st.date_input("วันที่", value=today()).strftime("%Y-%m-%d")
    names = dict(zip(br.branch_id, br.branch_name))
    sel = st.multiselect("เลือกสาขา (เลือกได้หลายสาขา)", list(names), format_func=names.get,
                         default=list(names)[:1])
    rec = not_deleted(read(REC))
    items = []
    with st.form("entry"):
        for b in sel:
            st.subheader(names[b])
            ex = rec[(rec.entry_date == ds) & (rec.branch_id == b)]
            gb = gates[gates.branch_id == b]
            for _, g in gb.iterrows():
                e = ex[ex.gate_id == g.gate_id]
                v = int(float(e.vehicle_count.iloc[0])) if not e.empty and str(e.vehicle_count.iloc[0]).strip() else 0
                cnt = st.number_input(f"{g.gate_name} (คัน)", min_value=0, step=1, value=v,
                                      key=f"c_{g.gate_id}_{ds}")
                items.append([b, g.gate_id, cnt])
            note = st.text_input("หมายเหตุ", value=ex.note.iloc[0] if not ex.empty else "",
                                 key=f"n_{b}_{ds}")
            for it in items:
                if it[0] == b and len(it) == 3:
                    it.append(note)
        ok = st.form_submit_button("💾 บันทึก", type="primary", use_container_width=True)
    if ok and items:
        n_new, n_upd = save_records(u, ds, [tuple(i) for i in items])
        st.success(f"บันทึกแล้ว: เพิ่มใหม่ {n_new} / แก้ไข {n_upd} รายการ")


def load_df(u):
    br = allowed_branches(u)
    rec = not_deleted(read(REC)).copy()
    rec = rec[rec.branch_id.isin(br.branch_id)]
    rec["entry_date"] = pd.to_datetime(rec.entry_date, errors="coerce")
    rec["vehicle_count"] = pd.to_numeric(rec.vehicle_count, errors="coerce").fillna(0).astype(int)
    rec = rec.merge(br[["branch_id", "branch_name"]], on="branch_id")
    rec = rec.merge(read("Gates")[["gate_id", "gate_name"]], on="gate_id", how="left")
    return rec.dropna(subset=["entry_date"])


def add_period(df, mode):
    d = df.entry_date
    if mode == "รายวัน":
        df["period"] = d.dt.strftime("%Y-%m-%d")
    elif mode == "รายสัปดาห์":  # แสดงเป็นวันจันทร์ต้นสัปดาห์
        df["period"] = (d - pd.to_timedelta(d.dt.weekday, unit="D")).dt.strftime("%Y-%m-%d")
    elif mode == "รายเดือน":
        df["period"] = d.dt.strftime("%Y-%m")
    else:
        df["period"] = d.dt.strftime("%Y")
    return df


def to_excel(sheets):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, d in sheets.items():
            d.to_excel(w, sheet_name=name[:31], index=False)
    return buf.getvalue()


def to_pdf(title, df):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    font = "Helvetica"
    if os.path.exists(FONT_PATH):
        pdfmetrics.registerFont(TTFont("TH", FONT_PATH))
        font = "TH"
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=landscape(A4), leftMargin=24, rightMargin=24, topMargin=24, bottomMargin=24)
    data = [list(map(str, df.columns))] + df.astype(str).values.tolist()
    t = Table(data, repeatRows=1)
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font), ("FONTSIZE", (0, 0), (-1, -1), 11),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f2f2f2")]),
    ]))
    doc.build([Paragraph(title, ParagraphStyle("t", fontName=font, fontSize=16)), Spacer(1, 8), t])
    return buf.getvalue()


def page_dashboard(u):
    st.header("📊 Dashboard")
    df = load_df(u)
    if df.empty:
        st.info("ยังไม่มีข้อมูล")
        return
    view = st.radio("รูปแบบ", ["รายงานแต่ละสาขา", "เปรียบเทียบสาขา"], horizontal=True)
    mode = st.selectbox("ช่วงเวลา", ["รายวัน", "รายสัปดาห์", "รายเดือน", "รายปี"])
    c1, c2 = st.columns(2)
    start = c1.date_input("ตั้งแต่", value=today() - timedelta(days=30))
    end = c2.date_input("ถึง", value=today())
    all_b = sorted(df.branch_name.unique())
    sel = st.multiselect("สาขา", all_b, default=all_b)
    df = df[(df.entry_date.dt.date >= start) & (df.entry_date.dt.date <= end) & df.branch_name.isin(sel)].copy()
    if df.empty:
        st.info("ไม่มีข้อมูลในช่วงที่เลือก")
        return
    skip0 = st.checkbox("ไม่นับวันที่ยอดเป็น 0 (เช่น เครื่องนับเสีย)", value=True)
    if skip0:
        day_tot = df.groupby(["branch_name", "entry_date"]).vehicle_count.transform("sum")
        n0 = df[day_tot == 0].groupby("branch_name").entry_date.nunique().sum()
        df = df[day_tot > 0].copy()
        if n0:
            st.caption(f"ไม่นับ {n0} วัน (สาขา-วัน) ที่ยอดรวมเป็น 0")
        if df.empty:
            st.info("ไม่มีข้อมูลในช่วงที่เลือก")
            return
    df = add_period(df, mode)

    if view == "รายงานแต่ละสาขา":
        tabs = st.tabs(sel)
        for t, bn in zip(tabs, sel):
            with t:
                x = df[df.branch_name == bn]
                if x.empty:
                    st.info("ไม่มีข้อมูล")
                    continue
                tot = x.groupby("period").vehicle_count.sum()
                a, b, c = st.columns(3)
                a.metric("รวม (คัน)", f"{tot.sum():,}")
                b.metric(f"เฉลี่ย/{mode[3:]}", f"{tot.mean():,.0f}")
                c.metric("สูงสุด", f"{tot.max():,}")
                st.bar_chart(tot)
                st.dataframe(x.pivot_table(index="period", columns="gate_name", values="vehicle_count",
                                           aggfunc="sum", fill_value=0), use_container_width=True)
    else:
        pv = df.pivot_table(index="period", columns="branch_name", values="vehicle_count",
                            aggfunc="sum", fill_value=0)
        st.bar_chart(pv)
        st.dataframe(pv, use_container_width=True)
        share = (pv.sum() / pv.values.sum() * 100).round(1).rename("สัดส่วน %").to_frame()
        share["รวม (คัน)"] = pv.sum()
        st.dataframe(share, use_container_width=True)

    summary = df.pivot_table(index="period", columns="branch_name", values="vehicle_count",
                             aggfunc="sum", fill_value=0).reset_index()
    detail = df[["entry_date", "branch_name", "gate_name", "vehicle_count", "note"]].sort_values(
        ["entry_date", "branch_name", "gate_name"])
    detail["entry_date"] = detail.entry_date.dt.strftime("%Y-%m-%d")
    st.divider()
    e1, e2 = st.columns(2)
    e1.download_button("⬇️ Excel", to_excel({"สรุป": summary, "รายละเอียด": detail}),
                       file_name=f"vehicle_{mode}_{start}_{end}.xlsx", use_container_width=True)
    e2.download_button("⬇️ PDF", to_pdf(f"Vehicle report {mode} {start} - {end}", summary),
                       file_name=f"vehicle_{mode}_{start}_{end}.pdf", mime="application/pdf",
                       use_container_width=True)


def page_manage(u):
    st.header("🛠️ แก้ไข / ลบข้อมูล")
    ds = st.date_input("วันที่", value=today()).strftime("%Y-%m-%d")
    cur = not_deleted(fresh(REC))
    cur = cur[cur.entry_date == ds]
    if cur.empty:
        st.info("ไม่มีข้อมูลในวันนี้")
        return
    lab = read("Gates").merge(read("Branches"), on="branch_id", suffixes=("_g", "_b"))
    lab = dict(zip(lab.gate_id, lab.branch_name + " - " + lab.gate_name))
    rid = st.selectbox("เลือกรายการ", list(cur.record_id),
                       format_func=lambda r: f"{lab.get(cur[cur.record_id == r].gate_id.iloc[0], r)} "
                                             f"({cur[cur.record_id == r].vehicle_count.iloc[0]} คัน)")
    r = cur[cur.record_id == rid].iloc[0]
    row = int(cur.index[cur.record_id == rid][0]) + 2
    cnt = st.number_input("จำนวนรถ", min_value=0, step=1, value=int(float(r.vehicle_count or 0)), key=f"ec_{rid}")
    note = st.text_input("หมายเหตุ", value=r.note, key=f"en_{rid}")
    ws = get_book().worksheet(REC)
    a, b = st.columns(2)
    if a.button("💾 บันทึกการแก้ไข", use_container_width=True):
        ws.batch_update([{"range": f"E{row}:F{row}", "values": [[int(cnt), note]]},
                         {"range": f"J{row}:K{row}", "values": [[u["email"], now()]]}], value_input_option="RAW")
        audit(rid, "update", f"count={r.vehicle_count};note={r.note}", f"count={int(cnt)};note={note}", u["email"])
        st.success("แก้ไขแล้ว")
        st.rerun()
    if b.button("🗑️ ลบรายการ", use_container_width=True):
        ws.batch_update([{"range": f"J{row}:L{row}", "values": [[u["email"], now(), "TRUE"]]}], value_input_option="RAW")
        audit(rid, "delete", f"count={r.vehicle_count};note={r.note}", "", u["email"])
        st.success("ลบแล้ว (soft delete)")
        st.rerun()


def page_import(u):
    st.header("📥 นำเข้าข้อมูลเดิม")
    st.caption("ไฟล์ .xlsx/.csv ต้องมีคอลัมน์: entry_date, branch_name, gate_name, vehicle_count, note (note ไม่บังคับ). "
               "ชื่อสาขา/ประตูต้องตรงกับชีต Branches/Gates. ถ้ามีข้อมูลวัน+ประตูเดิมอยู่แล้ว จะถูกเขียนทับ")
    f = st.file_uploader("เลือกไฟล์", type=["xlsx", "csv"])
    if not f:
        return
    df = pd.read_csv(f) if f.name.lower().endswith(".csv") else pd.read_excel(f)
    st.dataframe(df.head(20))
    if not st.button("นำเข้าข้อมูล", type="primary"):
        return
    m = read("Gates").merge(read("Branches"), on="branch_id", suffixes=("_g", "_b"))
    df["entry_date"] = pd.to_datetime(df["entry_date"], errors="coerce")
    df = df.merge(m[["gate_id", "branch_id", "gate_name", "branch_name"]], on=["branch_name", "gate_name"], how="left")
    bad = df[df.gate_id.isna() | df.entry_date.isna()]
    if len(bad):
        st.error(f"พบ {len(bad)} แถวที่จับคู่สาขา/ประตูหรือวันที่ไม่ได้ (แก้ไฟล์แล้วนำเข้าใหม่)")
        st.dataframe(bad)
        return
    if "note" not in df:
        df["note"] = ""
    new = pd.DataFrame({
        "record_id": [uuid.uuid4().hex[:12] for _ in range(len(df))],
        "entry_date": df.entry_date.dt.strftime("%Y-%m-%d"),
        "branch_id": df.branch_id, "gate_id": df.gate_id,
        "vehicle_count": pd.to_numeric(df.vehicle_count, errors="coerce").fillna(0).astype(int),
        "note": df.note.fillna("").astype(str),
        "created_by_uid": u["uid"], "created_by_email": u["email"], "created_at": now(),
        "updated_by": "", "updated_at": "", "is_deleted": "FALSE"})
    merged = pd.concat([fresh(REC), new]).drop_duplicates(["gate_id", "entry_date"], keep="last")
    rewrite(REC, merged)
    audit("-", "import", "", f"{len(new)} rows", u["email"])
    st.success(f"นำเข้าสำเร็จ {len(new)} แถว")


def page_users(u):
    st.header("👥 จัดการผู้ใช้")
    users = fresh("Users")
    st.dataframe(users[["email", "display_name", "role", "status", "last_login"]], use_container_width=True)
    br = read("Branches")
    bname = dict(zip(br.branch_id, br.branch_name))
    st.subheader("เพิ่ม / แก้ไขผู้ใช้")
    with st.form("uf"):
        email = st.text_input("อีเมล Google").strip().lower()
        name = st.text_input("ชื่อที่แสดง")
        role = st.selectbox("สิทธิ์", ["user", "viewer", "admin"])
        status = st.selectbox("สถานะ", ["active", "inactive"])
        bs = st.multiselect("สาขาที่เข้าถึงได้ (admin เข้าได้ทุกสาขาอยู่แล้ว)", list(bname), format_func=bname.get)
        ok = st.form_submit_button("บันทึกผู้ใช้", type="primary")
    if ok and email:
        ex = users[users.email.str.lower() == email]
        if ex.empty:
            uid = uuid.uuid4().hex[:10]
            append("Users", [[uid, email, name or email, role, status, now(), ""]])
        else:
            uid = ex.iloc[0].uid
            row = int(ex.index[0]) + 2
            get_book().worksheet("Users").batch_update(
                [{"range": f"C{row}:E{row}", "values": [[name or ex.iloc[0].display_name, role, status]]}],
                value_input_option="RAW")
        acc = fresh("UserBranchAccess")
        acc = acc[acc.uid != uid]
        add = pd.DataFrame([[uid, b, "active"] for b in bs], columns=H["UserBranchAccess"])
        rewrite("UserBranchAccess", pd.concat([acc, add]))
        st.success("บันทึกผู้ใช้แล้ว")
        st.rerun()


PAGES = {
    "📝 บันทึกข้อมูล": page_entry,
    "📊 Dashboard": page_dashboard,
    "🛠️ แก้ไข/ลบข้อมูล": page_manage,
    "📥 นำเข้าข้อมูลเดิม": page_import,
    "👥 ผู้ใช้งาน": page_users,
}
MENU = {
    "admin": list(PAGES),
    "user": ["📝 บันทึกข้อมูล"],
    "viewer": ["📊 Dashboard"],
}


def main():
    u = login()
    menu = MENU.get(u["role"], [])
    if not menu:
        st.error("สิทธิ์ผู้ใช้ไม่ถูกต้อง")
        st.stop()
    with st.sidebar:
        st.write(f"👤 {u['display_name'] or u['email']}")
        st.caption(f"สิทธิ์: {u['role']}")
        page = st.radio("เมนู", menu)
        st.button("ออกจากระบบ", on_click=st.logout)
    PAGES[page](u)


main()
