import csv
import io
import os
import re
import zipfile
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Bank Report to Excel Converter", page_icon="🏦", layout="centered"
)
st.markdown("[આ લિંક સીધી તમારા મોબાઈલમાં ખોલો](https://bank-report-converter-igahw2smogh7r9fdopummj.streamlit.app/)")

st.markdown(
    """
    <style>
    .main-title {
        text-align: center;
        color: #1F4E78;
        font-size: 26px;
        font-weight: bold;
        margin-bottom: 5px;
    }
    .sub-title {
        text-align: center;
        color: #555555;
        font-size: 15px;
        margin-bottom: 20px;
    }
    .stDownloadButton button {
        width: 100%;
        background-color: #1F4E78;
        color: white;
        font-weight: bold;
        height: 50px;
        border-radius: 8px;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="main-title">🏦 Bank Reports to Excel Converter</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="sub-title">તમામ ૮૮+ ટેક્સ્ટ રિપોર્ટ્સનું ૧૦૦% ક્લીન એક્સેલ'
    " કન્વર્ટર (V4 Precision Engine)</div>",
    unsafe_allow_html=True,
)

IGNORE_PATTERNS = [
    "REPORT ID:",
    "PROC DATE:",
    "RUN DATE:",
    "BRANCH :",
    "BRANCH NO.",
    "PAGE NO",
    "PRODUCT TOTAL",
    "ACCOUNT TYPE TOTAL",
    "OVER DRAWN TOTAL",
    "NO OF ACCOUNTS",
    "SUB TOTAL",
    "GRAND TOTAL",
    "BALANCE FORWARD",
    "BHAVNAGAR DISTRICT",
    "TOTAL FOR PRODUCT",
    "TOTAL NO OF",
    "PRODUCT-WISE TOTAL",
    "====>",
    "GL CLASS CODE",
    "GL-CLASS-CODE",
    "AREA:",
]


def clean_dr_amount(val):
  val_str = str(val).strip()
  if not val_str:
    return ""
  if (
      re.search(r"\bDr\b", val_str, re.IGNORECASE)
      or val_str.endswith("Dr")
      or val_str.endswith("dr")
  ):
    num_clean = re.sub(r"(?i)\s*dr\s*", "", val_str).strip()
    return f"-{num_clean}" if not num_clean.startswith("-") else num_clean
  elif re.search(r"\bCr\b", val_str, re.IGNORECASE):
    return re.sub(r"(?i)\s*cr\s*", "", val_str).strip()
  return val_str


# ૧. DEPOSITS BALANCE ફાઇલ માટે પ્રિસિઝન એન્જિન (સ્ક્રીનશૉટની એરરનું કાયમી સોલ્યુશન)
def parse_deposits_balance_report(lines):
  pattern = re.compile(
      r"^\s*(\d{11}-\d|\d{8,16})\s+"  # ૧. ખાતા નંબર
      r"(\S+(?:\s\S+)*)\s{2,}"  # ૨. ખાતાનો પ્રકાર
      r"(.+?)\s{2,}"  # ૩. ગ્રાહકનું નામ
      r"([\d,]+\.\d{2})\s+"  # ૪. Available Balance
      r"([\d,]+\.\d{2})\s+"  # ૫. Uncleared Balance
      r"([\d,]+\.\d{2}(?:\s*Dr)?)\s+"  # ૬. Current Balance
      r"([\d,]+\.\d{2})\s*"  # ૭. Limit
      r"(?:\s+(\d+[DMY]))?"  # ૮. Term (મુદ્દત)
      r"\s+([\d,]+\.\d{2})"  # ૯. Interest Rate
      r"\s+([A-Z]+)"  # ૧૦. Status
      r"\s+([YN])\s*$",  # ૧૧. Joint Flag
      re.IGNORECASE,
  )

  rows = []
  for l in lines:
    m = pattern.match(l)
    if m:
      curr_val = clean_dr_amount(m.group(6))
      rows.append({
          "ACCOUNT NUMBER": m.group(1).strip(),
          "ACCOUNT TYPE (DESCRIPTION)": m.group(2).strip(),
          "CUSTOMER NAME": m.group(3).strip(),
          "AVAILABLE BALANCE": m.group(4).strip(),
          "UNCLEARED BALANCE": m.group(5).strip(),
          "CURRENT BALANCE": curr_val,
          "LIMIT": m.group(7).strip(),
          "TERM": m.group(8).strip() if m.group(8) else "",
          "INT-RATE": m.group(9).strip(),
          "STATUS": m.group(10).strip(),
          "JOINT-HOLD-FLAG": m.group(11).strip(),
      })

  if rows:
    return pd.DataFrame(rows)
  return None


# ૨. પાઇપ (|) સેપરેટેડ ફાઇલો માટે પાર્સર
def parse_pipe_delimited(lines):
  header_cols = []
  data_rows = []
  for l in lines:
    stripped = l.strip()
    if (
        not stripped
        or re.match(r"^-{10,}", stripped)
        or stripped.startswith("^[")
        or stripped.startswith("\x0c")
    ):
      continue
    if any(kw in stripped.upper() for kw in IGNORE_PATTERNS):
      continue
    if "|" in stripped:
      parts = [clean_dr_amount(p.strip()) for p in stripped.split("|")]
      if parts and parts[0] == "":
        parts = parts[1:]
      if parts and parts[-1] == "":
        parts = parts[:-1]
      if not header_cols and any(
          w in stripped.upper()
          for w in ["NAME", "ACCOUNT", "AMOUNT", "LIMIT", "DATE", "PRODUCT"]
      ):
        header_cols = parts
      elif header_cols and len(parts) >= 2:
        if parts != header_cols:
          if len(parts) < len(header_cols):
            parts.extend([""] * (len(header_cols) - len(parts)))
          data_rows.append(parts[: len(header_cols)])
      elif not header_cols and len(parts) >= 3:
        header_cols = [f"Col_{i+1}" for i in range(len(parts))]
        data_rows.append(parts)
  if header_cols and data_rows:
    return pd.DataFrame(data_rows, columns=header_cols)
  return None


# ૩. સામાન્ય CBS બેંકિંગ રિપોર્ટ્સ માટે પાર્સર
def parse_general_cbs_report(lines):
  header_idx = -1
  for i, l in enumerate(lines):
    if re.match(r"^\s*-{15,}", l) and i + 2 < len(lines):
      if re.match(r"^\s*-{15,}", lines[i + 2]):
        header_idx = i + 1
        break
  if header_idx == -1:
    for i, l in enumerate(lines):
      if (
          i + 1 < len(lines)
          and re.match(r"^\s*-{15,}", lines[i + 1])
          and len(l.strip()) > 10
      ):
        header_idx = i
        break
  if header_idx == -1:
    return None

  raw_header_line = lines[header_idx]
  header_matches = list(
      re.finditer(r"(\S+(?:\s(?!\s)\S+)*)", raw_header_line)
  )
  if len(header_matches) < 2:
    return None

  col_headers = [m.group(1).strip() for m in header_matches]
  report_title = lines[header_idx - 2].strip() if header_idx >= 2 else ""

  data_rows = []
  for l in lines[header_idx + 2 :]:
    stripped = l.strip()
    if (
        not stripped
        or re.match(r"^-{10,}", stripped)
        or stripped.startswith("^[")
        or stripped.startswith("\x0c")
    ):
      continue
    if any(kw in stripped.upper() for kw in IGNORE_PATTERNS):
      continue
    if report_title and report_title.upper() in stripped.upper():
      continue
    if stripped == raw_header_line.strip() or (
        "ACCOUNT NUMBER" in stripped and "CUSTOMER NAME" in stripped
    ):
      continue
    if "SLNO" in stripped and "CUSTOMER" in stripped and "ACCOUNT" in stripped:
      continue

    parts = [p.strip() for p in re.split(r"\s{2,}", stripped) if p.strip() != ""]
    if len(parts) == len(col_headers):
      row_dict = {
          col_headers[k]: clean_dr_amount(parts[k])
          for k in range(len(col_headers))
      }
      data_rows.append(row_dict)
    elif len(parts) > 1:
      row_dict = {}
      for k in range(len(col_headers)):
        row_dict[col_headers[k]] = (
            clean_dr_amount(parts[k]) if k < len(parts) else ""
        )
      data_rows.append(row_dict)

  if data_rows:
    return pd.DataFrame(data_rows)
  return None


def parse_csv_generic(content_str):
  try:
    sample = content_str[:4096]
    dialect = csv.Sniffer().sniff(sample, delimiters=[",", "\t", "|", ";"])
    delim = dialect.delimiter
  except Exception:
    delim = "," if "," in content_str else "\t"
  try:
    df = pd.read_csv(
        io.StringIO(content_str), sep=delim, engine="python", on_bad_lines="skip"
    )
  except Exception:
    df = pd.read_csv(
        io.StringIO(content_str),
        sep=delim,
        engine="python",
        encoding="latin1",
        on_bad_lines="skip",
    )
  for col in df.columns:
    if df[col].dtype == object:
      df[col] = df[col].apply(clean_dr_amount)
  return df


def convert_dataframe_to_excel_bytes(df):
  output = io.BytesIO()
  wb = Workbook()
  ws = wb.active

  headers = list(df.columns)
  ws.append(headers)

  header_fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
  data_font = Font(name="Calibri", size=10)
  thin_border = Border(
      left=Side(style="thin", color="D9D9D9"),
      right=Side(style="thin", color="D9D9D9"),
      top=Side(style="thin", color="D9D9D9"),
      bottom=Side(style="thin", color="D9D9D9"),
  )

  for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )

  for row in df.itertuples(index=False):
    ws.append(list(row))

  for row in ws.iter_rows(
      min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column
  ):
    for cell in row:
      cell.font = data_font
      cell.border = thin_border
      val_str = str(cell.value or "").strip()
      if re.match(r"^-?[\d,]+(\.\d+)?$", val_str):
        cell.alignment = Alignment(horizontal="right")
      else:
        cell.alignment = Alignment(horizontal="left")

  for col in ws.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

  ws.freeze_panes = "A2"
  wb.save(output)
  output.seek(0)
  return output.getvalue()


# ----------------- STREAMLIT UI ----------------- #

uploaded_files = st.file_uploader(
    "મોબાઇલ કે કમ્પ્યુટરમાંથી ટેક્સ્ટ ફાઇલો પસંદ કરો (Select Files):",
    accept_multiple_files=True
)

if uploaded_files:
  total_files = len(uploaded_files)
  st.info(f"📁 પસંદ કરેલી ફાઇલોની સંખ્યા: {total_files}")

  if st.button("🚀 Convert to Excel (એક્સેલમાં કન્વર્ટ કરો)"):
    progress_bar = st.progress(0)
    status_text = st.empty()

    converted_files = {}
    failed_files = []

    for idx, file_obj in enumerate(uploaded_files):
      filename = file_obj.name
      base_name, _ = os.path.splitext(filename)
      status_text.text(
          f"પ્રોસેસ થઈ રહી છે: {filename} ({idx+1}/{total_files})..."
      )

      try:
        content_bytes = file_obj.read()
        try:
          content_str = content_bytes.decode("utf-8")
        except UnicodeDecodeError:
          content_str = content_bytes.decode("latin1", errors="ignore")

        lines = content_str.splitlines()

        df = None
        # ૧. Deposits Balance File
        if (
            "DEPOSITS BALANCE FILE" in content_str.upper()
            or "AVAILABLE BALANCE" in content_str.upper()
        ):
          df = parse_deposits_balance_report(lines)

        # ૨. પાઇપ (|) વાળી ફાઇલ
        if df is None or df.empty:
          pipe_count = sum(1 for l in lines[:50] if l.count("|") >= 3)
          if pipe_count >= 3:
            df = parse_pipe_delimited(lines)

        # ૩. સામાન્ય CBS ફાઇલ
        if df is None or df.empty:
          df = parse_general_cbs_report(lines)

        # ૪. CSV / Delimited
        if df is None or df.empty:
          df = parse_csv_generic(content_str)

        if df is not None and not df.empty:
          excel_bytes = convert_dataframe_to_excel_bytes(df)
          converted_files[f"{base_name}.xlsx"] = excel_bytes
        else:
          failed_files.append(filename)
      except Exception as e:
        failed_files.append(f"{filename} ({str(e)})")

      progress_bar.progress((idx + 1) / total_files)

    status_text.empty()
    progress_bar.empty()

    st.success(
        f"✅ સફળતાપૂર્વક કન્વર્ટ થયેલ ફાઇલો: {len(converted_files)} /"
        f" {total_files}"
    )
    if failed_files:
      st.warning(
          f"⚠️ સ્કીપ થયેલ ફાઇલો ({len(failed_files)}): {', '.join(failed_files)}"
      )

    if len(converted_files) == 1:
      only_name, only_data = list(converted_files.items())[0]
      st.download_button(
          label=f"📥 Download {only_name}",
          data=only_data,
          file_name=only_name,
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )
    elif len(converted_files) > 1:
      zip_buffer = io.BytesIO()
      with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for fname, fdata in converted_files.items():
          zip_file.writestr(fname, fdata)
      zip_buffer.seek(0)

      st.download_button(
          label=(
              "📦 Download All Excel Files (.ZIP -"
              f" {len(converted_files)} Files)"
          ),
          data=zip_buffer.getvalue(),
          file_name="Converted_Bank_Reports.zip",
          mime="application/zip",
      )
