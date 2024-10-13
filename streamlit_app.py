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


# Funksjon for å lese PDF-filen og hente ut relevante data inkludert enhet, rabatt, og pris
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
                        if len(columns) >= 7:  # For å inkludere varenummer, beskrivelse, antall, enhet, enhetspris, rabatt, totalpris
                            item_number = columns[1]
                            if not item_number.isdigit():
                                continue

                            # Beskrivelse kan være sammensatt av flere kolonner
                            description = []
                            for word in columns[2:]:
                                if word.isdigit() or word.replace(',', '').replace('.', '').isdigit():
                                    break  # Stopp når vi møter et tall
                                description.append(word)
                            description = " ".join(description)

                            # Antall, enhet, pris og beløp
                            try:
                                quantity = float(columns[-6].replace('.', '').replace(',', '.')) if columns[-6].replace(',', '').replace('.', '').isdigit() else None
                                unit = columns[-5]  # Enhet skal være tekst
                                
                                # Valutahåndtering: enhetspris, rabatt og totalpris
                                unit_price = float(columns[-4].replace('.', '').replace(',', '.')) if columns[-4].replace(',', '').replace('.', '').isdigit() else None
                                discount = float(columns[-3].replace('.', '').replace(',', '.')) if columns[-3].replace(',', '').replace('.', '').isdigit() else 0  # Rabatt
                                total_price = float(columns[-1].replace('.', '').replace(',', '.')) if columns[-1].replace(',', '').replace('.', '').isdigit() else None
                            except ValueError as e:
                                st.error(f"Kunne ikke konvertere til flyttall: {e}")
                                continue

                            # Lag UnikID basert på fakturanummer og varenummer
                            unique_id = f"{invoice_number}_{item_number}" if invoice_number else item_number
                            data.append({
                                "UnikID": unique_id,
                                "Varenummer": item_number,
                                "Beskrivelse_Faktura": description,
                                "Antall_Faktura": quantity if quantity else None,
                                "Enhet_Faktura": unit if unit else None,
                                "Enhetspris_Faktura": unit_price if unit_price else None,
                                "Rabatt": discount if discount else 0,
                                "Totalt_pris_Faktura": total_price if total_price else None,
                                "Type": doc_type
                            })
            if len(data) == 0:
                st.error("Ingen data ble funnet i PDF-filen.")
                
            return pd.DataFrame(data)
    except Exception as e:
        st.error(f"Kunne ikke lese data fra PDF: {e}")
        return pd.DataFrame()

# Resten av koden bør forbli den samme.


# Resten av koden bør forbli den samme.


# Resten av koden bør forbli den samme.
# Funksjon for å konvertere DataFrame til en Excel-fil
def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# Hovedfunksjon for Streamlit-appen
def main():
    st.title("Sammenlign Faktura mot Tilbud")

    # Opprett tre kolonner for layout
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
                'TOTALPRIS': 'Totalt pris_Tilbud'
            }, inplace=True)

            # Debug: Vis dataene som er lest inn for faktura og tilbud
            st.write("Faktura data etter ekstraksjon:")
            st.dataframe(invoice_data)
            st.write("Tilbud data etter lesing:")
            st.dataframe(offer_data)

            if not invoice_data.empty and not offer_data.empty:
                # Sammenligne faktura mot tilbud
                with col2:
                    st.write("Sammenligner data...")
                    # Rename the total price column in the invoice data to Beløp_Faktura
                    invoice_data.rename(columns={"Totalt pris": "Beløp_Faktura"}, inplace=True)

                # Merge the data using the renamed column
                merged_data = pd.merge(offer_data, invoice_data, on="Varenummer", how='outer', suffixes=('_Tilbud', '_Faktura'))

                # Ensure the correct columns are used when generating the final report
                all_items = invoice_data[["UnikID", "Varenummer", "Beskrivelse_Faktura", "Antall_Faktura", "Enhetspris_Faktura", "Beløp_Faktura", "Rabatt"]]

              
                # Debug: Vis det sammenslåtte datasettet
                st.write("Sammenslåtte data (tilbud og faktura):")
                st.dataframe(merged_data)

                # Konverter kolonner til numerisk
                merged_data["Antall_Faktura"] = pd.to_numeric(merged_data["Antall_Faktura"], errors='coerce')
                merged_data["Antall_Tilbud"] = pd.to_numeric(merged_data["Antall_Tilbud"], errors='coerce')
                merged_data["Enhetspris_Faktura"] = pd.to_numeric(merged_data["Enhetspris_Faktura"], errors='coerce')
                merged_data["Enhetspris_Tilbud"] = pd.to_numeric(merged_data["Enhetspris_Tilbud"], errors='coerce')

                # Finne avvik
                merged_data["Avvik_Antall"] = merged_data["Antall_Faktura"] - merged_data["Antall_Tilbud"]
                merged_data["Avvik_Enhetspris"] = merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]
                merged_data["Prosentvis_økning"] = ((merged_data["Enhetspris_Faktura"] - merged_data["Enhetspris_Tilbud"]) / merged_data["Enhetspris_Tilbud"]) * 100

                # Vis bare avvikene som faktisk har forskjeller
                avvik = merged_data[(merged_data["Avvik_Antall"].notna() & (merged_data["Avvik_Antall"] != 0)) |
                                    (merged_data["Avvik_Enhetspris"].notna() & (merged_data["Avvik_Enhetspris"] != 0))]

                with col2:
                    st.subheader("Avvik mellom Faktura og Tilbud")
                    st.dataframe(avvik)

                # Artikler som finnes i faktura, men ikke i tilbud
                only_in_invoice = merged_data[merged_data['Enhetspris_Tilbud'].isna()]
                with col2:
                    st.subheader("Varenummer som finnes i faktura, men ikke i tilbud")
                    st.dataframe(only_in_invoice)

                # Lagre kun artikkeldataene til XLSX
                all_items = invoice_data[["UnikID", "Varenummer", "Beskrivelse_Faktura", "Antall_Faktura", "Enhetspris_Faktura", "Totalt pris_Faktura"]]
                
                excel_data = convert_df_to_excel(all_items)

                with col3:
                    st.download_button(
                        label="Last ned avviksrapport som Excel",
                        data=convert_df_to_excel(avvik),
                        file_name="avvik_rapport.xlsx"
                    )
                    
                    st.download_button(
                        label="Last ned alle varenummer som Excel",
                        data=excel_data,
                        file_name="faktura_varer.xlsx",
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )

                    # Lag en Excel-fil med varenummer som finnes i faktura, men ikke i tilbud
                    only_in_invoice_data = convert_df_to_excel(only_in_invoice)
                    st.download_button(
                        label="Last ned varenummer som ikke eksiterer i tilbudet",
                        data=only_in_invoice_data,
                        file_name="varer_kun_i_faktura.xlsx",
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
        else:
            st.error("Fakturanummeret ble ikke funnet i PDF-filen.")
    else:
        st.warning("Last opp både faktura og tilbud for å fortsette.")

if __name__ == "__main__":
    main()

