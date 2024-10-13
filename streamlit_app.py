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
                        if len(columns) >= 5:
                            item_number = columns[1]
                            if not item_number.isdigit():
                                continue

                            description = " ".join(columns[2:-3])
                            try:
                                quantity = float(columns[-3].replace('.', '').replace(',', '.')) if columns[-3].replace('.', '').replace(',', '').isdigit() else columns[-3]
                                unit_price = float(columns[-2].replace('.', '').replace(',', '.')) if columns[-2].replace('.', '').replace(',', '').isdigit() else columns[-2]
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
                                "Beløp_Faktura": total_price,
                                "Rabatt": 0,  # Rabatt er foreløpig satt til 0 for alle rader
                                "Type": doc_type
                            })
            if len(data) == 0:
                st.error("Ingen data ble funnet i PDF-filen.")
                
            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {e}")
        return pd.DataFrame()

# Funksjon for å dele opp beskrivelsen basert på siste elementer
def split_description(data, doc_type):
    if doc_type == "Faktura":
        data['Enhet_Faktura'] = data['Beskrivelse_Faktura'].str.extract(r'(\bM2|\bM|\bSTK)$', expand=False)
        data['Beskrivelse_Faktura'] = data['Beskrivelse_Faktura'].str.replace(r'\s*\b(M2|M|STK)$', '', regex=True)

    return data

# Funksjon for å trekke ut antall fra beskrivelse_faktura
def extract_quantity_from_description(row):
    match = re.search(r'(\d+)$', row["Beskrivelse_Faktura"])
    if match:
        return float(match.group(1))
    return row["Antall_Faktura"]

# Funksjon for å konvertere DataFrame til en Excel-fil
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def extract_quantity_for_missing_items(data):
    # Denne funksjonen brukes kun for å trekke ut antall fra beskrivelsen for varer som ikke finnes i tilbudet
    for idx, row in data.iterrows():
        description = row["Beskrivelse_Faktura"]
        match = re.search(r'(\d+)\s*$', description)  # Matcher et tall på slutten av beskrivelsen
        if match:
            quantity = match.group(1)
            data.at[idx, "Antall_Faktura"] = float(quantity)  # Oppdaterer Antall_Faktura med verdien
            data.at[idx, "Beskrivelse_Faktura"] = description[:match.start()].strip()  # Fjerner antallet fra beskrivelsen
    return data

# Oppdater hovedfunksjonen for å bare bruke dette på varenummer som ikke finnes i tilbudet
def main():
    st.title("Sammenlign Faktura mot Tilbud")

    col1, col2, col3 = st.columns([1, 5, 1])

    with col1:
        st.header("Last opp filer")
        invoice_file = st.file_uploader("Last opp faktura fra Brødrene Dahl", type="pdf")
        offer_file = st.file_uploader("Last opp tilbud fra Brødrene Dahl (Excel)", type="xlsx")

    if invoice_file and offer_file:
        with col1:
            st.info("Henter fakturanummer fra faktura...")
            invoice_number = get_invoice_number(invoice_file)

        if invoice_number:
            with col1:
                st.success(f"Fakturanummer funnet: {invoice_number}")
            
            # Ekstraher data fra PDF-filen
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

            # Del opp beskrivelsen fra fakturaen
            if not invoice_data.empty:
                invoice_data = split_description(invoice_data, "Faktura")

            if not offer_data.empty:
                with col2:
                    st.write("Sammenligner data...")
                merged_data = pd.merge(offer_data, invoice_data, on="Varenummer", how='outer', suffixes=('_Tilbud', '_Faktura'))

                merged_data["Antall_Faktura"] = pd.to_numeric(merged_data["Antall_Faktura"], errors='coerce')
                merged_data["Antall_Tilbud"] = pd.to_numeric(merged_data["Antall_Tilbud"], errors='coerce')
                merged_data["Enhetspris_Faktura"] = pd.to_numeric(merged_data["Enhetspris_Faktura"], errors='coerce')
                merged_data["Enhetspris_Tilbud"] = pd.to_numeric(merged_data["Enhetspris_Tilbud"], errors='coerce')

                # Finne avvik
                merged_data["Avvik_Antall"] = merged_data["Antall_Faktura"] - merged_data["Antall_Tilbud"]
                merged_data["Avvik_Enhetspris"] = merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]
                merged_data["Prosentvis_økning"] = ((merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]) / merged_data["Enhetspris_Tilbud"]) * 100

                # Filter for varer som finnes i faktura, men ikke i tilbudet
                only_in_invoice = merged_data[merged_data['Enhetspris_Tilbud'].isna()]

                # Anvende extract_quantity_for_missing_items kun for "varenummer som finnes i faktura, men ikke i tilbud"
                only_in_invoice = extract_quantity_for_missing_items(only_in_invoice)

                with col2:
                    st.subheader("Avvik mellom Faktura og Tilbud")
                    st.dataframe(merged_data)

                    st.subheader("Varenummer som finnes i faktura, men ikke i tilbud")
                    st.dataframe(only_in_invoice)

                all_items = invoice_data[["UnikID", "Varenummer", "Beskrivelse_Faktura", "Antall_Faktura", "Enhetspris_Faktura", "Beløp_Faktura", "Rabatt"]]
                excel_data = convert_df_to_excel(all_items)

                with col3:
                    st.download_button(
                        label="Last ned avviksrapport som Excel",
                        data=convert_df_to_excel(merged_data),
                        file_name="avvik_rapport.xlsx"
                    )
                    
                    st.download_button(
                        label="Last ned alle varenummer som Excel",
                        data=excel_data,
                        file_name="faktura_varer.xlsx",
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )

                    only_in_invoice_data = convert_df_to_excel(only_in_invoice)
                    st.download_button(
                        label="Last ned varenummer som ikke eksiterer i tilbudet",
                        data=only_in_invoice_data,
                        file_name="varer_kun_i_faktura.xlsx",
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
            else:
                st.error("Kunne ikke lese tilbudsdata fra Excel-filen.")
        else:
            st.error("Fakturanummeret ble ikke funnet i PDF-filen.")

if __name__ == "__main__":
    main()
