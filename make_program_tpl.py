#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主日崇拜程序表 Word 生成器（模板版）
以教會原版 Word 做底，只更新文字；版面跟足原檔。
用法：
  python make_program_tpl.py 程序表-2026-09-27.json [-o 輸出.docx]
pip install python-docx requests
模板：
  MOS程序表2026-09-20 (1).docx  （普通主日）
  MOS程序表2026-08-2.docx       （有聖餐、第4主日）
"""
import argparse, copy, json, os, re
import requests
from docx import Document

GB = "https://api.getbible.net/v2/cut"
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def clean(t):
    return re.sub(r"\s+", "", t or "")

def fetch_verses(b, c, v):
    try:
        r = requests.get(f"{GB}/{int(b)}/{int(c)}.json", timeout=25); r.raise_for_status()
        verses = r.json().get("verses", [])
    except Exception:
        return None
    if not verses:
        return None
    nums = []
    if v and str(v).strip():
        for part in str(v).split(","):
            part = part.strip()
            if "-" in part:
                a, z = part.split("-"); nums += list(range(int(a), int(z) + 1))
            else:
                nums.append(int(part))
    else:
        nums = list(range(1, len(verses) + 1))
    return [f"{n}{clean(verses[n-1].get('text',''))}" for n in nums if 1 <= n <= len(verses)]

def set_text(p, text):
    if not p.runs:
        p.add_run(text); return
    p.runs[0].text = text
    for r in p.runs[1:]:
        r._element.getparent().remove(r._element)

def find_p(doc, prefixes):
    for p in doc.paragraphs:
        t = p.text.strip()
        for pre in prefixes:
            if t.startswith(pre):
                return p
    return None

def find_i(doc, prefixes):
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        for pre in prefixes:
            if t.startswith(pre):
                return i
    return -1

def clone_p(template_p, text):
    np_el = copy.deepcopy(template_p._p)
    runs = np_el.findall(W + "r")
    for r in runs[1:]:
        np_el.remove(r)
    if runs:
        r0 = runs[0]
        ts = r0.findall(W + "t")
        for t in ts[1:]:
            r0.remove(t)
        if ts:
            ts[0].text = text
        else:
            t = np_el.makeelement(W + "t", {}); t.text = text; r0.append(t)
    else:
        r = np_el.makeelement(W + "r", {}); t = np_el.makeelement(W + "t", {}); t.text = text
        r.append(t); np_el.append(r)
    return np_el

def delete_paras(body, start_el, stop_el=None):
    children = list(body)
    try:
        si = children.index(start_el)
    except ValueError:
        return
    for i in range(si + 1, len(children)):
        c = children[i]
        if stop_el is not None and c is stop_el:
            break
        if c.tag.split('}')[-1] in ("p", "tbl"):
            body.remove(c)

def strip_sdt(doc):
    """移除所有 Word 內容控制（combo box / drop-down / checkbox sdt）。
       因為每個人名格之後都會由 data 重新填寫，所以直接刪走成個 sdt（連內裡舊預設文字），
       令 Word 開出嚟唔再有下拉掣。"""
    for sdt in list(doc.element.body.iter(W + "sdt")):
        parent = sdt.getparent()
        if parent is None:
            continue
        parent.remove(sdt)

def m(v):
    try:
        n = int(float(v or 0))
        return f"${n:,}" if n else "-"
    except (TypeError, ValueError):
        return "-"

def append_cell_row(tbl, template_row_idx, cell_texts):
    """複製 template_row（通常係「香膏瓶」行）做新一行 append 落表尾，
       再逐一改 cell 文字。cell_texts 係 [上主日格, 月份格] 兩個。"""
    if template_row_idx >= len(tbl.rows):
        return
    template_tr = tbl.rows[template_row_idx]._tr
    new_tr = copy.deepcopy(template_tr)
    tbl._tbl.append(new_tr)
    # 用新 append 咗嘅 row 去 set 文字
    new_row = tbl.rows[-1]
    for ci, text in enumerate(cell_texts):
        if ci < len(new_row.cells):
            set_text(new_row.cells[ci].paragraphs[0], text)

SUNDAY = os.path.join(os.path.dirname(os.path.abspath(__file__)), "MOS程序表2026-09-20 (1).docx")
SACR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "MOS程序表2026-08-2.docx")

def build(data, outpath):
    sach = bool(data.get("sach"))
    doc = Document(SACR if sach else SUNDAY)
    strip_sdt(doc)               # 先移除所有 combo box / dropdown（人名下拉）
    ps = doc.paragraphs
    body = doc.element.body

    # ===== 1) 日期 =====
    from datetime import datetime
    try:
        dt = datetime.strptime(data.get("date",""), "%Y-%m-%d")
        dstr = f"{dt.year} 年 {dt.month} 月 {dt.day} 日" if sach else f"{dt.year} 年 {dt.month} 月 {dt.day} 日 "
    except ValueError:
        dstr = data.get("date","")
    done_date = False
    for i, p in enumerate(ps):
        if re.search(r"(20\d\d)\s*年", p.text) and "日" in p.text:
            set_text(p, dstr); done_date = True; break
    if not done_date:
        # fallback: 揾「主日崇拜」前一非空段
        for i, p in enumerate(ps):
            if p.text.strip() == "主日崇拜" and i > 0:
                set_text(ps[i-1], dstr); break

    # ===== 2) 事奉人員 表（t1 = 樣式 3x2）=====
    # 套用通常佈置：t0聚會, t1事奉, t2下主日, t3人數, t4崇拜奉獻, t5什一+香膏瓶
    def set_staff_table(tbl, src, type):
        rows = tbl.rows
        if type == "today":
            pairs = [("講員", src.get("speaker"), "領會", src.get("leader")),
                     ("司樂", src.get("musician"), "招待", src.get("usher")),
                     ("感恩分享", src.get("thanks"), "", "")]
        else:
            pairs = [("講員", src.get("sp"), "領會", src.get("ld")),
                     ("司樂", src.get("mu"), "招待", src.get("us")),
                     ("感恩分享", src.get("th"), "", "")]
        for ri, (l1, v1, l2, v2) in enumerate(pairs):
            if ri < len(rows):
                c = rows[ri].cells[0]
                set_text(c.paragraphs[0], f"{l1}：{v1 or ''}" if l1 else "")
                if len(rows[ri].cells) > 1:
                    set_text(rows[ri].cells[1].paragraphs[0], f"{l2}：{v2 or ''}" if l2 else "")

    try:
        set_staff_table(doc.tables[1], data, "today")
        set_staff_table(doc.tables[2], data.get("next") or {}, "next")
    except IndexError:
        pass

    # ===== 3) 程序段落 =====
    def set_pf(prefixes, text):
        p = find_p(doc, prefixes)
        if p is not None: set_text(p, text)

    xz = data.get("xz") or {}
    if xz.get("ref"):
        lines = fetch_verses(xz.get("b"), xz.get("c"), xz.get("v")) if xz.get("b") else None
        bodytxt = " ".join(l[1:] for l in lines) if lines else ""
        set_pf(["宣召"], "宣召 " + (bodytxt or xz.get("ref","")) + "    " + xz.get("ref",""))
    names = "、".join([h.get("name","") for h in (data.get("worship") or []) if h.get("name")])
    if names: set_pf(["敬拜"], "敬拜  " + names)
    dj0 = (data.get("dj") or [{}])[0]
    if dj0.get("ref"): set_pf(["讀經"], "讀經  " + dj0.get("ref","") + " " * 30 + "眾立")
    if data.get("sermon"): set_pf(["證道"], "證道  " + data.get("sermon","") + " " * 20 + str(data.get("speaker") or ""))
    rname = (data.get("resp") or {}).get("name","")
    if rname: set_pf(["回應"], "回應  " + rname)
    if sach:
        p = find_p(doc, ["聖餐"])
        if p is not None: set_text(p, "聖餐" + " " * 30 + str(data.get("sachPerson") or ""))
    if data.get("offering"):
        p = find_p(doc, ["奉獻"])
        if p is not None: set_text(p, "奉獻 " + "、".join(data.get("offering")))
    if data.get("welcomePerson"):
        p = find_p(doc, ["歡迎及報告"])
        if p is not None: set_text(p, "歡迎及報告\t\t" + data.get("welcomePerson"))
    if data.get("bless"):
        p = find_p(doc, ["祝福"])
        if p is not None: set_text(p, "祝福\t\t\t" + data.get("bless"))

    # ===== 4) 報告事項（重建）=====
    ri = find_i(doc, ["報 告 事 項"])
    if ri >= 0:
        start_el = ps[ri]._p
        stop_i = find_i(doc, ["下主日崇拜事奉人員"])
        stop_el = ps[stop_i]._p if stop_i >= 0 else None
        template = None
        for p in ps[ri+1:]:
            if p.text.strip() and not p.text.startswith("下主日"):
                template = p; break
        delete_paras(body, start_el, stop_el)
        records = data.get("reports") or []
        anchor = start_el
        if template is not None and records:
            for rep in records:
                np = clone_p(template, rep)
                anchor.addnext(np); anchor = np
        elif records:
            for rep in records:
                p2 = doc.add_paragraph(rep); start_el.addnext(p2._p); anchor = p2._p

    # ===== 5) 下主日（已喺 set_staff_table(doc.tables[2]) 做）=====

    # ===== 6) 出席人數 表3：3x2 格 [label]:value =====
    att = data.get("attendance") or {}
    # 表3（人數）有多款佈置：09-20 同 08-2 都係 3x2 兩欄
    if len(doc.tables) >= 4:
        at = doc.tables[3]
        amap = [("主日崇拜", "sun"), ("主日學", "sch"), ("門徒訓練", "disc"),
                ("每月祈禱會", "pr"), ("學生團契", "fel"), ("Patrick飯局", "pat")]
        # 平鋪所有格（左->右、上->下）
        flat = []
        for row in at.rows:
            for c in row.cells:
                flat.append(c)
        for idx, (label, key) in enumerate(amap):
            if idx < len(flat):
                cur = flat[idx].text.split("：")[0]
                set_text(flat[idx].paragraphs[0], f"{cur}：{att.get(key) if att.get(key) else ''}")

    # ===== 7) 奉獻：由標籤更新 =====
    # 崇拜奉獻: cWeek/cMonth/cMonthAmt ; 什一: tWeek/tMonth/tMonthAmt
    # 香膏瓶: oilWeek/oilMonthAmt ; 感恩: thWeek/thMonthAmt
    def find_cell_row(tbl, label_part):
        for row in tbl.rows:
            t0 = row.cells[0].text.strip()
            if t0.startswith(label_part):
                return row
        return None

    # 崇拜奉獻表（t4 樣式：1x2 上主日 | 月份）
    if len(doc.tables) >= 5:
        t4 = doc.tables[4]
        # 上主日崇拜奉獻 所在表：t4
        if len(t4.rows) >= 1 and len(t4.columns) >= 2:
            set_text(t4.rows[0].cells[0].paragraphs[0], f"上主日：{m(data.get('cWeek'))}")
            set_text(t4.rows[0].cells[1].paragraphs[0],
                     f"{data.get('cMonth') or ''}：{m(data.get('cMonthAmt'))}")
    if len(doc.tables) >= 6:
        t5 = doc.tables[5]
        # t5: 什一（上主日|月份）＋ 香膏瓶（上主日|月份）
        if sach:
            # 08-2: t5 2x2: (上主日什一|七月份), (香膏瓶|六月份)
            if len(t5.rows) >= 1:
                set_text(t5.rows[0].cells[0].paragraphs[0], f"上主日：{m(data.get('tWeek'))}")
                set_text(t5.rows[0].cells[1].paragraphs[0], f"{data.get('tMonth') or ''}：{m(data.get('tMonthAmt'))}")
            if len(t5.rows) >= 2:
                set_text(t5.rows[1].cells[0].paragraphs[0], f"香膏瓶：{m(data.get('oilWeek'))}")
                set_text(t5.rows[1].cells[1].paragraphs[0], f"{data.get('oilMonth') or ''}：{m(data.get('oilMonthAmt'))}")
        else:
            # 09-20: t5 2x2: (上主日|九月份),(香膏瓶|九月份)
            if len(t5.rows) >= 1:
                set_text(t5.rows[0].cells[0].paragraphs[0], f"上主日：{m(data.get('tWeek'))}")
                set_text(t5.rows[0].cells[1].paragraphs[0], f"{data.get('tMonth') or ''}：{m(data.get('tMonthAmt'))}")
            if len(t5.rows) >= 2:
                set_text(t5.rows[1].cells[0].paragraphs[0], f"香膏瓶：{m(data.get('oilWeek'))}")
                set_text(t5.rows[1].cells[1].paragraphs[0], f"{data.get('oilMonth') or ''}：{m(data.get('oilMonthAmt'))}")
    # 感恩奉獻（gOn=true，加第三行）
    if data.get("gOn") and len(doc.tables) >= 6:
        t5g = doc.tables[5]
        texts = [
            f"感恩奉獻：{m(data.get('thWeek'))}",
            f"{data.get('thMonth') or ''}：{m(data.get('thMonthAmt'))}",
        ]
        if len(t5g.rows) >= 3:
            set_text(t5g.rows[2].cells[0].paragraphs[0], texts[0])
            set_text(t5g.rows[2].cells[1].paragraphs[0], texts[1])
        else:
            append_cell_row(t5g, 1, texts)

    doc.save(outpath)
    print("OK 已生成:", outpath)

# ===== 通用 block 重建 helper =====
def rebuild_block(doc, start_prefixes, stop_prefix, items,
                  title_style_from=None, line_style_from=None):
    """由第一個 start_prefix 段落之後開始，刪到 stop_prefix 段落之前（唔包），
       再依次用 line_style 樣式插入 items。items 每項係 str。"""
    ps = doc.paragraphs; body = doc.element.body
    si = find_i(doc, start_prefixes)
    if si < 0:
        return
    start_el = ps[si]._p
    # 搵 stop
    stop_el = None
    if stop_prefix:
        for p in ps[si+1:]:
            if p.text.strip().startswith(stop_prefix):
                stop_el = p._p; break
    # 樣式範本：由 items 前方現有段落揾一個唔同於題目嘅做 line 樣式
    tmpl = None
    for p in ps[si+1:]:
        if stop_el is not None and p._p is stop_el:
            break
        if p.text.strip() and not p.text.strip().startswith(start_prefixes[0]):
            tmpl = p; break
    delete_paras(body, start_el, stop_el)
    anchor = start_el
    if tmpl is not None:
        for it in items:
            np = clone_p(tmpl, it)
            anchor.addnext(np); anchor = np
    else:
        for it in items:
            p2 = doc.add_paragraph(it); start_el.addnext(p2._p); anchor = p2._p

def set_block_text(doc, prefix, items, line_prefix=None):
    """單一段落變長塊：將原段落確實做 anchor，刪其後，再插 items。"""
    ps = doc.paragraphs; body = doc.element.body
    p = find_p(doc, [prefix])
    if p is None: return
    start_el = p._p
    # 刪所有緊接喺後面直至下一張表／唔屬 content 前？直接刪到文件段落尾
    children = list(body)
    si = children.index(start_el)
    tmpl = None
    for i in range(si+1, len(children)):
        c = children[i]
        tag = c.tag.split('}')[-1]
        if tag == "tbl":
            break
        if tag == "p":
            t0 = c.find('.//'+W+'t')
            txt = ''.join(tt.text or '' for tt in c.findall('.//'+W+'t'))
            if txt.strip():
                tmpl = c
    delete_paras(body, start_el)  # 刪晒 start 後全部 p
    anchor = start_el
    # 標題本身保留，之前已經 set_text 咗
    for it in items:
        if tmpl is not None:
            np = clone_p(tmpl, it); anchor.addnext(np); anchor = np
        else:
            p2 = doc.add_paragraph(it); start_el.addnext(p2._p); anchor = p2._p

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("json"); ap.add_argument("-o", "--out", default=None)
    a = ap.parse_args()
    data = json.load(open(a.json, encoding="utf-8"))
    out = a.out or a.json.rsplit(".json", 1)[0] + "_程序表.docx"
    build(data, out)                       # 先由模板生成
    _post_generate_hook(data, out)          # 再補內文（金句/代禱/詩歌/讀經）

# 生成後補充：每月金句正文、代禱、詩歌歌詞、讀經內文
def _post_generate_hook(data, outpath):
    sach = bool(data.get("sach"))
    doc = Document(outpath)  # 載返啱啱生成嘅檔再補
    ps = doc.paragraphs
    body = doc.element.body

    # 1) 每月金句正文（頭版：每月金句下面嗰段「Bla Bla」（ref））
    jy = data.get("jinyue") or {}
    pj = find_p(doc, ["「"])
    if pj is not None:
        txt = "「" + (jy.get("ref","") if jy.get("ref") else "（未揀每月金句）") + "」"
        set_text(pj, txt.replace("（（","（").replace("）」","）"))
        # ref 已喺 jy.ref
    # 每月金句正文改用經文全文字（如有）
    if jy.get("ref"):
        lines = fetch_verses(jy.get("b"), jy.get("c"), jy.get("v")) if jy.get("b") else None
        if lines:
            bodytxt = " ".join(l[1:] for l in lines)
            pj2 = find_p(doc, ["「"])
            if pj2 is not None:
                set_text(pj2, "「" + bodytxt + "」 （" + jy.get("ref","") + "）")

    # 2) 代禱事項（重建）
    daos = data.get("daodaos") or []
    if daos:
        rebuild_block(doc, ["代 禱 事 項"], "基督教宣聖會", daos)

    from docx.oxml import OxmlElement
    # 注意：OxmlElement 要用簡記前綴 (w:p / w:r / w:t)，唔可以用完整 URI
    #    格式：新一頁 + 字體縮細 + 詩歌之間加空位（配合雙欄逼入 1 張 A4）
    hymns = []  # 每項 (text, is_title)
    seen = set()
    for h in (data.get("worship") or []):
        name = (h or {}).get("name") or ""
        lyrics = (h or {}).get("lyrics") or ""
        if (not name and not lyrics.strip()) or (name and name in seen):
            continue
        if name:
            seen.add(name)
        hymns.append((name, True))
        for ln in lyrics.splitlines():
            if ln.strip():
                hymns.append((ln, False))
    rname = (data.get("resp") or {}).get("name")
    if rname and rname not in seen:
        rly = (data.get("resp") or {}).get("lyrics") or ""
        hymns.append((rname, True))
        for ln in rly.splitlines():
            if ln.strip():
                hymns.append((ln, False))
    if hymns:
        ri = find_i(doc, ["讀經："])
        if ri > 0:
            # 詩歌起點 = 最後一張表之後嘅第一段
            bodychildren = list(doc.element.body)
            last_tbl_el = None
            for c in bodychildren:
                if c.tag.split('}')[-1] == "tbl":
                    last_tbl_el = c
            start_el = None
            if last_tbl_el is not None:
                ci = bodychildren.index(last_tbl_el)
                for c in bodychildren[ci+1:]:
                    if c.tag.split('}')[-1] == "p":
                        start_el = c; break
            if start_el is None:
                mi = find_i(doc, ["默禱"])
                start_el = doc.paragraphs[mi+1]._p if mi >= 0 else doc.paragraphs[1]._p
            stop_el = doc.paragraphs[ri]._p
            tmpl = None
            for p in doc.paragraphs:
                if p._p is stop_el:
                    break
                if p.text.strip():
                    tmpl = p
            # ---- A) 喺歌詞前加「分頁符」段落（新一頁）----
            pb = doc.add_paragraph()
            pbr = pb.add_run()
            brk = OxmlElement('w:br'); brk.set(W+'type', 'page')
            pbr._element.append(brk)
            start_el.addprevious(pb._p)
            # ---- B) 刪舊歌詞 ----
            delete_paras(doc.element.body, start_el, stop_el)
            # ---- C) 重建歌詞（縮細字體）----
            anchor = start_el
            body_sz = 26   # 歌詞 13pt（half-points乘以2→sz=26 = 13pt）
            title_sz = 30  # 詩歌標題 15pt
            for txt, is_t in hymns:
                if tmpl is not None:
                    np = clone_p(tmpl, txt)
                else:
                    np = OxmlElement('w:p')
                    r0 = OxmlElement('w:r'); t0 = OxmlElement('w:t'); t0.text = txt
                    r0.append(t0); np.append(r0)
                # 統一改字體大細（所有 run）
                for run in np.findall('.//'+W+'r'):
                    rpr = run.find(W+'rPr')
                    if rpr is None:
                        rpr = OxmlElement('w:rPr'); run.insert(0, rpr)
                    sz = rpr.find(W+'sz')
                    if sz is None:
                        sz = OxmlElement('w:sz'); rpr.insert(0, sz)
                    sz.set(W+'val', str(title_sz if is_t else body_sz))
                anchor.addnext(np); anchor = np
                # 詩歌之間加空位：每首詩歌嘅標題之前加一個細空行
                if is_t:
                    sp = OxmlElement('w:p'); r0 = OxmlElement('w:r'); t0 = OxmlElement('w:t'); t0.text = ""
                    r0.append(t0); sp.append(r0)
                    rpr = OxmlElement('w:rPr'); sz0 = OxmlElement('w:sz'); sz0.set(W+'val', '16'); rpr.append(sz0); r0.insert(0, rpr)
                    spPr = OxmlElement('w:pPr'); spc = OxmlElement('w:spacing'); spc.set(W+'after', '40'); spPr.append(spc)
                    sp.insert(0, spPr)
                    anchor.addnext(sp); anchor = sp

    # 4) 讀經內文（後頁「讀經：」之後）
    adj = (data.get("dj") or [])[0] if data.get("dj") else None
    if adj and adj.get("b"):
        lines = fetch_verses(adj.get("b"), adj.get("c"), adj.get("v"))
        # 標題
        pj = find_p(doc, ["讀經："])
        if pj is not None:
            set_text(pj, "讀經：" + (adj.get("ref","") or ""))
        if lines:
            # 讀經段落由標題之後
            rebuild_block(doc, ["讀經："], None, lines)

    doc.save(outpath)
    print("OK 已生成（含內文）:", outpath)
if __name__ == "__main__":
    main()
