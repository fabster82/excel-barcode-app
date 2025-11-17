import streamlit as st
import pandas as pd
import tempfile
import os
import io

from barcode import EAN13, EAN8, Code128
from barcode.writer import ImageWriter
import xlsxwriter


HEADER_ROW_IDX = 12   # Zeile 13 (0-basiert)
DATA_START_IDX = 13   # Zeile 14 (0-basiert)
EAN_COL_IDX = 1       # Spalte B (Index 1)
BARCODE_COL_IDX = 6   # Spalte G (Index 6)


def generate_excel_with_barcodes(uploaded_file) -> io.BytesIO:
    # Excel einlesen (erste Tabelle, ohne Header-Interpretation)
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = xls.sheet_names[0]
    df = xls.parse(sheet_name, header=None)

    # Sicherstellen, dass es mindestens bis Spalte G (Index 6) Spalten gibt
    while df.shape[1] <= BARCODE_COL_IDX:
        df[df.shape[1]] = None

    # Temporäres Verzeichnis für Barcode-Bilder
    tmpdir = tempfile.mkdtemp()

    # Output im Speicher (BytesIO), damit wir einen Download anbieten können
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {"in_memory": True})
    worksheet = workbook.add_worksheet(sheet_name)

    # Alle vorhandenen Werte aus df in das neue Sheet schreiben
    rows, cols = df.shape
    for r in range(rows):
        for c in range(cols):
            val = df.iat[r, c]
            if pd.isna(val):
                continue
            worksheet.write(r, c, val)

    # Überschrift für Barcode-Spalte in Zeile 13 (Index 12)
    worksheet.write(HEADER_ROW_IDX, BARCODE_COL_IDX, "Barcode")

    # Zeilenweise Barcodes generieren und in Spalte G einfügen
    for r in range(DATA_START_IDX, rows):
        ean_val = df.iat[r, EAN_COL_IDX]

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

            # Bild in Zelle (r, Spalte G) einfügen
            worksheet.insert_image(
                r,
                BARCODE_COL_IDX,
                img_file,
                {"x_scale": 0.7, "y_scale": 0.7}
            )
        except Exception as err:
            # Wenn eine Zeile Probleme macht, überspringen wir sie einfach
            print(f"Fehler in Zeile {r+1} für EAN {ean_str}: {err}")
            continue

    workbook.close()
    output.seek(0)
    return output


def main():
    st.title("Excel → EAN-Barcodes einfügen")
    st.write(
        """
        Dieses Tool erwartet eine Excel-Datei wie deine Preisliste:

        - Überschriften in **Zeile 13**
        - Daten ab **Zeile 14**
        - **EAN in Spalte B**
        - Barcodes werden als **Bilder in Spalte G** eingefügt
        """
    )

    uploaded_file = st.file_uploader("Excel-Datei hochladen (.xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        st.info("Datei erkannt. Klicke unten auf „Barcodes generieren“.")

        if st.button("Barcodes generieren"):
            try:
                result = generate_excel_with_barcodes(uploaded_file)
                st.success("Fertige Datei mit Barcodes erzeugt!")

                st.download_button(
                    label="Excel-Datei mit Barcodes herunterladen",
                    data=result,
                    file_name="preisliste_mit_barcodes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error(f"Es ist ein Fehler aufgetreten: {e}")


if __name__ == "__main__":
    main()
