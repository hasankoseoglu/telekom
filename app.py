import os
from typing import Optional, Tuple

from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd


app = Flask(__name__)
app.secret_key = "change-me-in-production"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOADED_XLSX_PATH = os.path.join(BASE_DIR, "uploaded.xlsx")
DEFAULT_XLSX_PATH = os.path.join(BASE_DIR, "Abone rehber.abn.xlsx")
DROP_COLUMNS_1_INDEXED = {
    1,2, 4, 5, 7, 8, 10, 11, 13, 16, 19, 20, 21, 22,23,24, 25, 26, 27, 28, 29,
    30, 31, 32, 33, 34, 35, 36,37,38,39,40,41,42,43,44,45,46, 47,48,50,51,52,53,54,55,56, 57, 58, 59, 60,61,63,64, 65, 66, 67, 68, 69,
    # Additional columns to drop
    70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80,
    81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91,
}
TITLE_OVERRIDES_1_INDEXED = {
    3: "Numara",
    6: "Hat Durumu",
    9: "İptal Tarihi",
    12: "Aktiflik Tarihi",
    14: "Ad Soyad",
    15: "T.C No",
    17: "Firma Ünvanı",
    18: "Firma Vergi No",
    49: "İletişim Numarası",
    62: "Adres",
   
}


def read_excel_as_textframe(xlsx_path: str) -> pd.DataFrame:
    """Read an Excel file as all-text DataFrame; split by ';|' if data is a single column."""
    df = pd.read_excel(xlsx_path, header=None, dtype=str)
    if df.shape[1] == 1:
        single_col = df.iloc[:, 0]
        # If delimiter ';|' seems present, split it into columns
        if single_col.str.contains(r";\|", na=False).any():
            df = single_col.str.split(r";\|", expand=True)
    # Normalize all values to string for consistent filtering/rendering
    df = df.astype(str)
    # Remove '|' characters from all cells
    df = df.replace(r"\|", "", regex=True)
    return df


def load_active_dataframe() -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """Load the most relevant DataFrame and return (df, source_path)."""
    if os.path.exists(UPLOADED_XLSX_PATH):
        try:
            return read_excel_as_textframe(UPLOADED_XLSX_PATH), UPLOADED_XLSX_PATH
        except Exception:
            pass
    if os.path.exists(DEFAULT_XLSX_PATH):
        try:
            return read_excel_as_textframe(DEFAULT_XLSX_PATH), DEFAULT_XLSX_PATH
        except Exception:
            pass
    return None, None


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file or file.filename == "":
            flash("Lütfen bir Excel dosyası seçin.", "warning")
            return redirect(url_for("index"))
        filename_lower = file.filename.lower()
        if not (filename_lower.endswith(".xlsx") or filename_lower.endswith(".xls")):
            flash("Sadece .xlsx veya .xls uzantılı dosyalar desteklenir.", "danger")
            return redirect(url_for("index"))
        try:
            file.save(UPLOADED_XLSX_PATH)
            flash("Dosya yüklendi.", "success")
        except Exception as exc:
            flash(f"Dosya kaydedilirken hata oluştu: {exc}", "danger")
        return redirect(url_for("index"))

    # GET request: show search UI
    query = request.args.get("q", "").strip()
    df, source_path = load_active_dataframe()

    file_info = None
    if source_path is not None:
        file_info = os.path.basename(source_path)

    rows = []
    headers = []
    total_count = 0
    filtered_count = 0
    error_message = None

    if df is not None:
        total_count = len(df)
        # Merge column 14 and 15 (1-indexed) -> put result into 14, drop 15
        num_cols = df.shape[1]
        if num_cols >= 15:
            left = df.iloc[:, 13].fillna("")
            right = df.iloc[:, 14].fillna("")
            merged = (left + " " + right).str.strip()
            df.iloc[:, 13] = merged
            # drop original 15th column after merge
            df.drop(df.columns[14], axis=1, inplace=True)
            num_cols = df.shape[1]
        # Drop selected columns by 1-indexed positions
        drop_zero_indexed = {i - 1 for i in DROP_COLUMNS_1_INDEXED if 1 <= i <= num_cols}
        keep_indices = [i for i in range(num_cols) if i not in drop_zero_indexed]
        if keep_indices:
            df = df.iloc[:, keep_indices]
        else:
            # If everything would be dropped, keep original to avoid empty table
            keep_indices = list(range(num_cols))
        headers = [TITLE_OVERRIDES_1_INDEXED.get(i + 1, f"Kolon {i+1}") for i in keep_indices]

        if query:
            try:
                mask = df.apply(lambda s: s.str.contains(query, case=False, na=False))
                mask_any = mask.any(axis=1)
                df_view = df[mask_any]
            except Exception as exc:
                error_message = f"Arama sırasında hata: {exc}"
                df_view = df
        else:
            df_view = df

        filtered_count = len(df_view)
        # Limit rows for display
        df_view = df_view.head(200)
        rows = df_view.values.tolist()

    return render_template(
        "index.html",
        file_info=file_info,
        headers=headers,
        rows=rows,
        total_count=total_count,
        filtered_count=filtered_count,
        query=query,
        error_message=error_message,
    )


# Enable running with: python app.py
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)


