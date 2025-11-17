import streamlit as st
import pandas as pd
import tempfile
import os
import io

from barcode import EAN13, EAN8, Code128
from barcode.writer import ImageWriter
import xlsxwriter


def excel_col_to_index(col: str) -> int:
    """
    Wandelt einen Excel-Spaltenbuchstaben (z.B. 'A', 'B', 'G') in
    einen 0-basierten Index um (A=0, B=1, ...).
    """
    col = col.strip().upper()
    result = 0
    for ch in col:
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Ungültiger Spaltenbuchstabe: {col}")
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1  # 0-basiert


def generate_excel_with_barcodes(
    uploaded_file,
    ean_col_letter: str,
    header_row_excel: int,
    data_start_row_excel: int,
    barcode_col_letter: str,
    row_height: float,
    col_width: float,
) -> io.BytesIO:
    # Excel einlesen (erste Tabelle, ohne Header-Interpretation)
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = xls.sheet_names[0]
    df = xls.parse(sheet_name, header=None)

    # Excel-Zeilen (1-basiert) → 0-basiert für pandas/xlsxwriter
    header_row_idx = header_row_excel - 1
    data_start_idx = data_start_row_excel - 1

    # Spaltenbuchstaben → 0-basierten Index
    ean_col_idx = excel_col_to_index(ean_col_letter)
    barcode_col_idx = excel_col_to_index(barcode_col_letter)

    # Sicherstellen, dass es mindestens bis zur Barcode-Spalte Spalten gibt
    while df.shape[1] <= barcode_col_idx:
        df[df.shape[1]] = None

    # Temporäres Verzeichnis für Barcode-Bilder
    tmpdir = tempfile.mkdtemp()

    # Output im Speicher (BytesIO), damit wir einen Download anbieten können
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    worksheet = workbook.add_worksheet(sheet_name)

    rows, cols = df.shape

    # Alle vorhandenen Werte aus df in das neue Sheet schreiben
    for r in range(rows):
        for c in range(cols):
            val = df.iat[r, c]
            if pd.isna(val):
                continue
            worksheet.write(r, c, val)

    # Überschrift für Barcode-Spalte (z.B. Zeile 13)
    worksheet.write(header_row_idx, barcode_col_idx, "Barcode")

    # Spalte für Barcodes etwas breiter machen
    worksheet.set_column(barcode_col_idx, barcode_col_idx, col_width)

    # Zeilenhöhe ab Datenbeginn erhöhen (für bessere Scanbarkeit)
    for r in range(data_start_idx, rows):
        worksheet.set_row(r, row_height)

    # Zeilenweise Barcodes generieren und in Barcode-Spalte einfügen
    for r in range(data_start_idx, rows):
        ean_val = df.iat[r, ean_col_idx]

        if pd.isna(ean_val):
            continue

        ean_str = str(ean_val).strip()

        # Typischer Fall: Excel hat aus EAN eine Zahl mit .0 gemacht
        if ean_str.endswith(".0"):
            ean_str = ean_str[:-2]

        # Nur Ziffern behalten
        ean_str = "".join(ch for ch in ean_str if ch.isdigit())

        if not ean_str:
            continue

        # Barcode-Typ wählen
        if len(ean_str) == 13:
            bc_class = EAN13
        elif len(ean_str) == 8:
            bc_class = EAN8
        else:
            # Fallback: Code128 für „ungewöhnliche“ Längen
            bc_class = Code128

        try:
            # Dateiname für dieses Bild
            base_path = os.path.join(tmpdir, f"barcode_row_{r}")
            bc_obj = bc_class(ean_str, writer=ImageWriter())
            img_file = bc_obj.save(base_path)  # gibt Pfad inkl. Endung zurück

            # Bild in Zelle (r, Barcode-Spalte) einfügen
            worksheet.insert_image(
                r,
                barcode_col_idx,
                img_file,
                {"x_scale": 0.7, "y_scale": 0.7},
            )
        except Exception as err:
            # Wenn eine Zeile Probleme macht, überspringen
            print(f"Fehler in Zeile {r+1} für EAN {ean_str}: {err}")
            continue

    workbook.close()
    output.seek(0)
    return output


def main():
    st.title("Excel → EAN-Barcodes einfügen")

    st.write(
        """
        Dieses Tool fügt Barcodes in eine Excel-Liste ein.

        Du kannst einstellen:
        - **Spalte mit EAN-Codes** (z. B. `B`)
        - **Zeile mit Überschriften** (z. B. `13`)
        - **Zeile, in der die Werte beginnen** (z. B. `14`)
        - **Spalte für die Barcodes** (z. B. `G`)
        - **Zeilenhöhe** und **Spaltenbreite** für die Barcodes
        """
    )

    uploaded_file = st.file_uploader("Excel-Datei hochladen (.xlsx)", type=["xlsx"])

    st.subheader("Einstellungen für deine Datei")

    col1, col2 = st.columns(2)

    with col1:
        ean_col_letter = st.text_input(
            "Spalte mit EAN-Codes (Buchstabe)",
            value="B",
            help="In deinem Beispiel ist das die Spalte B.",
        )
        header_row_excel = st.number_input(
            "Zeile mit Überschrift",
            min_value=1,
            value=13,
            help="In deinem Beispiel: Zeile 13.",
        )
        data_start_row_excel = st.number_input(
            "Startzeile für Werte",
            min_value=1,
            value=14,
            help="In deinem Beispiel: ab Zeile 14.",
        )

    with col2:
        barcode_col_letter = st.text_input(
            "Spalte für Barcodes (Buchstabe)",
            value="G",
            help="In deinem Beispiel: Spalte G.",
        )
        row_height = st.number_input(
            "Zeilenhöhe für Barcode-Zeilen",
            min_value=10.0,
            max_value=200.0,
            value=40.0,
            step=1.0,
            help="Höher = mehr Abstand, besser scannbar.",
        )
        col_width = st.number_input(
            "Spaltenbreite der Barcode-Spalte",
            min_value=5.0,
            max_value=100.0,
            value=25.0,
            step=1.0,
            help="Breiter = Barcode läuft nicht an den Rand.",
        )

    if uploaded_file is not None:
        st.info("Datei erkannt. Klicke unten auf „Barcodes generieren“.")

        if st.button("Barcodes generieren"):
            try:
                result = generate_excel_with_barcodes(
                    uploaded_file=uploaded_file,
                    ean_col_letter=ean_col_letter,
                    header_row_excel=int(header_row_excel),
                    data_start_row_excel=int(data_start_row_excel),
                    barcode_col_letter=barcode_col_letter,
                    row_height=float(row_height),
                    col_width=float(col_width),
                )

                st.success("Fertige Datei mit Barcodes erzeugt!")

                st.download_button(
                    label="Excel-Datei mit Barcodes herunterladen",
                    data=result,
                    file_name="preisliste_mit_barcodes.xlsx",
                    mime=(
                        "application/"
                        "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    ),
                )
            except Exception as e:
                st.error(f"Es ist ein Fehler aufgetreten: {e}")


if __name__ == "__main__":
    main()
