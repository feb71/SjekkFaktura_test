import streamlit as st
import pdfplumber
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Streamlit App", layout="wide", initial_sidebar_state="expanded")

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
                    if doc_type == "Faktura" and "Artikkel" in line:
                        start_reading = True
                        continue

                    if start_reading:
                        columns = line.split()
                        if len(columns) >= 6:
                            item_number = columns[1]
                            if not item_number.isdigit():
                                continue

                            description = " ".join(columns[2:-4])
                            try:
                                quantity = float(columns[-4].replace('.', '').replace(',', '.')) if columns[-4].replace('.', '').replace(',', '').isdigit() else columns[-4]
                                unit_price = float(columns[-3].replace('.', '').replace(',', '.')) if columns[-3].replace('.', '').replace(',', '').isdigit() else columns[-3]
                                discount = columns[-2]
                                total_price = float(columns[-1].replace('.', '').replace(',', '.')) if columns[-1].replace('.', '').replace(',', '').isdigit() else columns[-1]
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
                                "Rabatt_Faktura": discount,
                                "Totalt pris": total_price,
                                "Type": doc_type
                            })
            if len(data) == 0:
                st.error("Ingen data ble funnet i PDF-filen.")
                
            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {e}")
        return pd.DataFrame()

# Funksjon for å konvertere DataFrame til en Excel-fil
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# Hovedfunksjon for Streamlit-appen
def main():
    st.title("Sammenlign Faktura mot Tilbud")

    # Opprett tre kolonner
    col1, col2, col3 = st.columns([1, 5, 1])

    with col1:
        st.header("Last opp filer")
        invoice_file = st.file_uploader("Last opp faktura fra Brødrene Dahl", type="pdf")
        offer_file = st.file_uploader("Last opp tilbud fra Brødrene Dahl (Excel)", type="xlsx")

    if invoice_file and offer_file:
        # Hent fakturanummer
        with col1:
            st.info("Henter fakturanummer fra faktura...")
            invoice_number = get_invoice_number(invoice_file)

        if invoice_number:
            with col1:
                st.success(f"Fakturanummer funnet: {invoice_number}")
            
            # Ekstraher data fra PDF-filer
            with col1:
                st.info("Laster inn faktura...")
            invoice_data = extract_data_from_pdf(invoice_file, "Faktura", invoice_number)

            # Les tilbudet fra Excel-filen
            with col1:
                st.info("Laster inn tilbud fra Excel-filen...")
            offer_data = pd.read_excel(offer_file)

            # Riktige kolonnenavn fra Excel-filen for tilbud
            offer_data.rename(columns={
                'VARENR': 'Varenummer',
                'BESKRIVELSE': 'Beskrivelse_Tilbud',
                'ANTALL': 'Antall_Tilbud',
                'ENHET': 'Enhet_Tilbud',
                'ENHETSPRIS': 'Enhetspris_Tilbud',
                'TOTALPRIS': 'Totalt pris'
            }, inplace=True)

            # Sammenligne faktura mot tilbud for varenummer
            merged_data = pd.merge(offer_data, invoice_data, on="Varenummer", how='outer', suffixes=('_Tilbud', '_Faktura'))

            # Finn avvik mellom faktura og tilbud (kun varer som finnes i begge)
            avvik = merged_data[merged_data['Beskrivelse_Tilbud'].notna() & merged_data['Beskrivelse_Faktura'].notna()]

            # Finn varer som kun finnes i fakturaen (inkluderer rabatt i denne tabellen)
            only_in_invoice = merged_data[merged_data['Beskrivelse_Tilbud'].isna()]

            # Finne avvik
            avvik["Avvik_Antall"] = avvik["Antall_Faktura"] - avvik["Antall_Tilbud"]
            avvik["Avvik_Enhetspris"] = avvik["Enhetspris_Faktura"] - avvik["Enhetspris_Tilbud"]
            avvik["Prosentvis_økning"] = ((avvik["Enhetspris_Faktura"] - avvik["Enhetspris_Tilbud"]) / avvik["Enhetspris_Tilbud"]) * 100

            # Vise avvikstabellen (kun varer som finnes i både faktura og tilbud)
            with col2:
                st.subheader("Avvik mellom Faktura og Tilbud")
                st.dataframe(avvik)

            # Vise varenummer som finnes i faktura men ikke i tilbud (med rabatt inkludert)
            with col2:
                st.subheader("Varenummer som finnes i faktura, men ikke i tilbud (med rabatt)")
                st.dataframe(only_in_invoice)

            # Lagre kun artikkeldataene til XLSX
            all_items = invoice_data[["UnikID", "Varenummer", "Beskrivelse_Faktura", "Antall_Faktura", "Enhetspris_Faktura", "Rabatt_Faktura", "Totalt pris"]]
            
            excel_data = convert_df_to_excel(all_items)

            with col3:
                st.download_button(
                    label="Last ned alle varenummer som Excel",
                    data=excel_data,
                    file_name="faktura_varer.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

                # Lag en Excel-fil med varenummer som finnes i faktura, men ikke i tilbud
                only_in_invoice_data = convert_df_to_excel(only_in_invoice)
                st.download_button(
                    label="Last ned varenummer som ikke eksiterer i tilbudet (med rabatt)",
                    data=only_in_invoice_data,
                    file_name="varer_kun_i_faktura_med_rabatt.xlsx",
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
        else:
            st.error("Fakturanummeret ble ikke funnet i PDF-filen.")

if __name__ == "__main__":
    main()
