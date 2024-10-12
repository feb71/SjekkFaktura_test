import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

# Funksjon for å lese fakturanummer fra PDF
def get_invoice_number(file):
    try:
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                match = re.search(r"Fakturanummer\s*[:\-]?\s*(\d+)", text, re.IGNORECASE)
                if match:
                    return match.group(1)
        return None
    except Exception as e:
        st.error(f"Kunne ikke lese fakturanummer fra PDF: {e}")
        return None

# Funksjon for å lese PDF-filen og hente ut relevante data
def extract_data_from_pdf(file, doc_type, invoice_number=None):
    try:
        with pdfplumber.open(file) as pdf:
            data = []
            start_reading = False

            for page in pdf.pages:
                text = page.extract_text()
                if text is None:
                    st.error(f"Ingen tekst funnet på side {page.page_number} i PDF-filen.")
                    continue
                
                lines = text.split('\n')
                for line in lines:
                    # Sjekk etter kolonneoverskriftene vi er ute etter: "Nummer", "Bransjenr.", "Beskrivelse", "MVA%", "Antall", "Enhet", "Pris", "Beløp"
                    if doc_type == "Faktura" and "Nummer" in line and "Bransjenr." in line:
                        start_reading = True
                        continue

                    if start_reading:
                        columns = line.split()
                        if len(columns) >= 8:
                            item_number = columns[0]  # Nummer = Varenummer
                            if not item_number.isdigit():
                                continue

                            description = " ".join(columns[2:-5])  # Kombiner beskrivelsen
                            try:
                                # Pris = Enhetspris, Beløp = Totalt pris, og antall kan beregnes
                                total_price = float(columns[-1].replace('.', '').replace(',', '.')) if columns[-1].replace('.', '').replace(',', '').isdigit() else None
                                unit_price = float(columns[-3].replace('.', '').replace(',', '.')) if columns[-3].replace('.', '').replace(',', '').isdigit() else None
                                
                                # Antall kan beregnes ved å dele beløp på pris
                                if total_price and unit_price:
                                    quantity = total_price / unit_price
                                else:
                                    quantity = None
                            except ValueError as e:
                                st.error(f"Kunne ikke konvertere til flyttall: {e}")
                                continue

                            unique_id = f"{invoice_number}_{item_number}" if invoice_number else item_number
                            data.append({
                                "UnikID": unique_id,
                                "Varenummer": item_number,
                                "Beskrivelse_Faktura": description,
                                "Antall_Faktura": quantity,
                                "Enhetspris_Faktura": unit_price,
                                "Totalt pris": total_price,
                                "Type": doc_type
                            })
            if len(data) == 0:
                st.error("Ingen data ble funnet i PDF-filen.")
                
            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {e}")
        return pd.DataFrame()

# Funksjon for å sammenligne faktura med tilbud
def compare_invoice_offer(invoice_data, offer_data):
    # Merge faktura og tilbud på varenummer
    merged_data = pd.merge(offer_data, invoice_data, on="Varenummer", how='outer', suffixes=('_Tilbud', '_Faktura'))
    
    # Konverter kolonner til numerisk for å sikre korrekt beregning
    merged_data["Antall_Faktura"] = pd.to_numeric(merged_data["Antall_Faktura"], errors='coerce')
    merged_data["Antall_Tilbud"] = pd.to_numeric(merged_data["Antall_Tilbud"], errors='coerce')
    merged_data["Enhetspris_Faktura"] = pd.to_numeric(merged_data["Enhetspris_Faktura"], errors='coerce')
    merged_data["Enhetspris_Tilbud"] = pd.to_numeric(merged_data["Enhetspris_Tilbud"], errors='coerce')

    # Beregn avvik
    merged_data["Avvik_Antall"] = merged_data["Antall_Faktura"] - merged_data["Antall_Tilbud"]
    merged_data["Avvik_Enhetspris"] = merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]
    
    # Prosentvis økning i enhetspris
    merged_data["Prosentvis_økning"] = ((merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]) / merged_data["Enhetspris_Tilbud"]) * 100

    # Filtrer for å vise kun avvik
    avvik = merged_data[(merged_data["Avvik_Antall"].notna() & (merged_data["Avvik_Antall"] != 0)) |
                        (merged_data["Avvik_Enhetspris"].notna() & (merged_data["Avvik_Enhetspris"] != 0))]
    
    # Varenummer som ikke finnes i tilbud
    only_in_invoice = merged_data[merged_data['Enhetspris_Tilbud'].isna()]
    
    return avvik, only_in_invoice, merged_data

# Hovedfunksjon for Streamlit-appen
def main():
    st.title("Sammenlign Faktura mot Tilbud")

    # Opplasting av filer
    invoice_file = st.file_uploader("Last opp faktura fra Brødrene Dahl", type="pdf")
    offer_file = st.file_uploader("Last opp tilbud fra Brødrene Dahl (Excel)", type="xlsx")

    if invoice_file and offer_file:
        # Hent fakturanummer
        invoice_number = get_invoice_number(invoice_file)

        if invoice_number:
            # Ekstraher data fra PDF-filer
            invoice_data = extract_data_from_pdf(invoice_file, "Faktura", invoice_number)

            # Les tilbudsdata
            offer_data = pd.read_excel(offer_file)
            offer_data.rename(columns={
                'VARENR': 'Varenummer',
                'BESKRIVELSE': 'Beskrivelse_Tilbud',
                'ANTALL': 'Antall_Tilbud',
                'ENHET': 'Enhet_Tilbud',
                'ENHETSPRIS': 'Enhetspris_Tilbud',
                'TOTALPRIS': 'Totalt pris_Tilbud'
            }, inplace=True)

            # Sammenlign faktura med tilbud
            avvik, only_in_invoice, merged_data = compare_invoice_offer(invoice_data, offer_data)

            # Vis avvik
            st.subheader("Avvik mellom Faktura og Tilbud")
            st.dataframe(avvik)

            # Vis varenummer som kun finnes i faktura
            st.subheader("Varenummer som kun finnes i faktura")
            st.dataframe(only_in_invoice)

            # Last ned avviksrapport som Excel
            def convert_df_to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                return output.getvalue()

            st.download_button(
                label="Last ned avviksrapport som Excel",
                data=convert_df_to_excel(avvik),
                file_name="avvik_rapport.xlsx"
            )

            st.download_button(
                label="Last ned varer kun i faktura",
                data=convert_df_to_excel(only_in_invoice),
                file_name="varer_kun_i_faktura.xlsx"
            )
        else:
            st.error("Fakturanummeret ble ikke funnet i PDF-filen.")
    else:
        st.info("Vennligst last opp både faktura og tilbud for sammenligning.")

if __name__ == "__main__":
    main()
