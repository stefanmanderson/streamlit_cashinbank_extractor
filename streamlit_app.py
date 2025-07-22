import streamlit as st
import fitz  # PyMuPDF
import re
import pandas as pd
import tempfile

# --- Regex Patterns ---
NOTE_HEAD = re.compile(r"\n?\s*(\d+)\.\s*Kas dan Setara Kas", re.I)
NEXT_NOTE = re.compile(r"^\s*(\d+)\s*\.\s+", re.M)
DAY_MONTH = r"(?:\d{1,2}\s+(?:Januari|Februari|Maret|April|Mei|Juni|Juli|" \
            r"Agustus|September|Oktober|November|Desember|January|February|" \
            r"March|April|May|June|July|August|September|October|November|" \
            r"December))"
DATE_FULL = re.compile(DAY_MONTH + r"\s+\d{4}", re.I)
DATE_HALF = re.compile(DAY_MONTH + r"(\s*/\s*)?$", re.I)
NUM_RGX = re.compile(r"\d{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?")
BANK_CAPTION = re.compile(r"^(?:PT\s+)?[A-Z0-9&.'\- ()]*\bBANK\b.*", re.I)
TOTAL_RGX = re.compile(r"\b(total|jumlah)\b", re.I)
COMPOSITE_RGX = re.compile(r",\s*[^,]*\bdan\b", re.I)

CURRENCY_HINT = {
    "rupiah": "IDR", "us dollar": "USD", "dolar as": "USD",
    "peso": "PHP", "yen": "JPY", "euro": "EUR"
}

def is_amount(val: int) -> bool:
    return val >= 1_000

def company(doc):
    text = doc[0].get_text()
    m = re.search(r"(PT\s+[A-Z0-9 .,&()\"'-]+?\s+Tbk)", text, re.I)
    return m.group(1).title().strip() if m else "Unknown"

def extract_note_text(doc):
    buf, capturing, note_no = [], False, None
    for pg in doc:
        txt = pg.get_text("text")
        if not capturing:
            m = NOTE_HEAD.search(txt)
            if not m: continue
            note_no, capturing = m.group(1), True
            txt = txt[m.start():]
        nxt = NEXT_NOTE.search(txt)
        if nxt and nxt.group(1) != note_no:
            buf.append(txt[: nxt.start()])
            break
        buf.append(txt)
    return "\n".join(buf)

def build_periods(txt):
    full = [m.group(0).title() for m in DATE_FULL.finditer(txt)]
    if len(full) >= 2:
        return full[:2]
    lines = [l.strip() for l in txt.splitlines() if l.strip()]
    halfs, years = [], []
    for l in lines[:60]:
        if DATE_HALF.match(l):
            halfs.append(DATE_HALF.match(l).group(0).replace("/", "").title())
        if (y := re.search(r"\b(20\d{2})\b", l)):
            years.append(y.group(1))
    if halfs and years:
        p0 = f"{halfs[0]} {years[0]}"
        p1 = f"{(halfs[1] if len(halfs)>1 else halfs[0])} {years[1] if len(years)>1 else years[0]}"
        return [p0, p1]
    return ["Current"]

def parse_note(text):
    periods, currency, place = build_periods(text), "IDR", "cash_in_bank"
    rows, lines = [ln.strip() for ln in text.splitlines() if ln.strip()], []

    i = 0
    results = []
    while i < len(rows):
        ln = rows[i]
        low = ln.lower()

        if "deposito" in low or "time deposit" in low:
            place = "time_deposit"
        if "kas di bank" in low or "cash in banks" in low:
            place = "cash_in_bank"

        for kw, iso in CURRENCY_HINT.items():
            if kw in low:
                currency = iso

        if BANK_CAPTION.match(ln) or TOTAL_RGX.search(ln) or COMPOSITE_RGX.search(ln):
            bank = "TOTAL" if TOTAL_RGX.search(ln) or COMPOSITE_RGX.search(ln) else \
                   re.sub(r"\s*(\(PERSERO\))?\s*TBK", "", ln, flags=re.I).upper()

            amounts = []
            clean = re.sub(r"\([^)]*\)", "", ln)
            for m in NUM_RGX.findall(clean):
                v = int(re.sub(r"[^\d]", "", m))
                if is_amount(v): amounts.append(v)
                if len(amounts) == 2: break

            j = i + 1
            while j < len(rows) and len(amounts) < 2:
                if NUM_RGX.fullmatch(rows[j]):
                    v = int(re.sub(r"[^\d]", "", rows[j]))
                    if is_amount(v): amounts.append(v)
                elif BANK_CAPTION.match(rows[j]) or TOTAL_RGX.search(rows[j]) or COMPOSITE_RGX.search(rows[j]):
                    dup = re.sub(r"\s*(\(PERSERO\))?\s*TBK", "", rows[j], flags=re.I).upper()
                    if dup.strip(",.") == bank.strip(",."): j += 1; continue
                    break
                j += 1

            for idx, amt in enumerate(amounts[:2]):
                results.append(dict(period=periods[idx] if idx < len(periods) else periods[0],
                                    bank=bank, currency=currency,
                                    amount=amt, placement_type=place))
            i = j
            continue
        i += 1

    uniq, seen = [], set()
    for r in results:
        key = (r['period'], r['bank'], r['amount'], r['placement_type'])
        if key not in seen:
            seen.add(key)
            uniq.append(r)
    return uniq

def process(file):
    doc = fitz.open(stream=file.read(), filetype="pdf")
    note = extract_note_text(doc)
    data = parse_note(note)
    comp = company(doc)
    for r in data:
        r["company"] = comp
    return pd.DataFrame(data)

# --- Streamlit App ---
st.set_page_config(page_title="Kas dan Setara Kas Extractor", layout="centered")
st.title("📊 Ekstraksi Kas dan Setara Kas dari Laporan Keuangan PDF")
st.write("Unggah file laporan keuangan dalam format PDF. Sistem akan mengekstrak data dari catatan *Kas dan Setara Kas* dan menyimpannya ke dalam Excel.")

uploaded_file = st.file_uploader("Pilih file PDF", type="pdf")

if uploaded_file is not None:
    with st.spinner("🔍 Memproses file..."):
        df = process(uploaded_file)
        st.success("✅ Ekstraksi selesai!")

        st.dataframe(df)

        # Create and offer download link
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            df.to_excel(tmp.name, index=False)
            st.download_button(
                label="⬇️ Unduh Excel",
                data=open(tmp.name, "rb").read(),
                file_name="kas_dan_setara_kas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
