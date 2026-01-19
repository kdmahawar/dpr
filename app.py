import streamlit as st
import re
from io import BytesIO
from openpyxl import load_workbook

# --- पेज सेटिंग ---
st.set_page_config(page_title="DPR Auto-Filler", layout="wide")
st.title("📊 WhatsApp to Excel: DPR Automation (Final V3)")
st.markdown("यह टूल बुलेट (•) हो या न हो, तारीख और डेटा को सही से अपडेट करेगा।")

# --- 1. फाइल अपलोडर ---
uploaded_file = st.file_uploader("अपनी Excel Template यहाँ अपलोड करें (.xlsx)", type=["xlsx"])

# --- 2. टेक्स्ट इनपुट ---
raw_text = st.text_area("WhatsApp Message यहाँ पेस्ट करें:", height=300)

# --- प्रोसेस बटन ---
if st.button("Excel अपडेट करें"):
    if uploaded_file and raw_text:
        try:
            # 1. एक्सेल फाइल लोड करें
            wb = load_workbook(uploaded_file)
            ws = wb.active
            
            # -----------------------------------------------
            # PART A: तारीख (Date) अपडेट करना
            # -----------------------------------------------
            date_pattern = r"Date:\s*([\d]{1,2}[/-][\d]{1,2}[/-][\d]{2,4})"
            date_match = re.search(date_pattern, raw_text, re.IGNORECASE)
            
            date_found = False
            new_date = "Unknown"
            
            if date_match:
                new_date = date_match.group(1)
                
                # एक्सेल की ऊपर की 10 लाइनों में "Date:" शब्द ढूँढें
                for row in ws.iter_rows(min_row=1, max_row=10, max_col=10):
                    for cell in row:
                        if cell.value and isinstance(cell.value, str) and "Date:" in cell.value:
                            # सेल में तारीख अपडेट करें
                            cell.value = f"Date: {new_date}"
                            date_found = True
                            break
                    if date_found:
                        break
            
            # -----------------------------------------------
            # PART B: डेटा (Figures) अपडेट करना (Updated Regex)
            # -----------------------------------------------
            # (?:•\s*)? का मतलब है: बुलेट और स्पेस 'ऑप्शनल' हैं (हो तो ठीक, न हो तो भी ठीक)
            pattern = (
                r"\*(.*?)(?::)?\*\s+"           # Name line (Example: *Silica Sand:*)
                r"(?:•\s*)?Daily:\s*([\d.]+).*?\n\s*"   # Daily line
                r"(?:•\s*)?Monthly:\s*([\d.]+).*?\n\s*" # Monthly line
                r"(?:•\s*)?Yearly:\s*([\d.]+)"          # Yearly line
            )
            
            matches = re.findall(pattern, raw_text, re.MULTILINE)
            
            # डेटा मैप तैयार करना
            data_map = {}
            for match in matches:
                # नाम में से : हटाकर साफ करें
                clean_name = match[0].replace(":", "").strip().lower()
                data_map[clean_name] = {
                    'daily': float(match[1]),
                    'monthly': float(match[2]),
                    'yearly': float(match[3])
                }
            
            updated_count = 0
            
            # एक्सेल की पंक्तियाँ (Rows) स्कैन करें
            for row in ws.iter_rows(min_row=1, max_col=6):
                name_cell = row[1]  # Column B (Name)
                
                if name_cell.value:
                    cell_value = str(name_cell.value).strip().lower()
                    
                    if cell_value in data_map:
                        values = data_map[cell_value]
                        
                        # डेटा अपडेट करें
                        row[3].value = values['daily']   # Column D
                        row[4].value = values['monthly'] # Column E
                        row[5].value = values['yearly']  # Column F
                        
                        updated_count += 1

            # -----------------------------------------------
            # PART C: फाइल सेव और डाउनलोड
            # -----------------------------------------------
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            msg_date = f" (Date updated: {new_date})" if date_found else " (Date not found in Excel)"
            st.success(f"सफलतापूर्वक! कुल {updated_count} एंट्रीज अपडेट की गईं।{msg_date}")
            
            file_name_date = new_date.replace('/', '-') if new_date != "Unknown" else "Updated"
            
            st.download_button(
                label="📥 अपडेटेड Excel फाइल डाउनलोड करें",
                data=output,
                file_name=f"DPR_{file_name_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error: {e}")
            
    else:
        st.warning("⚠️ कृपया पहले Excel फाइल अपलोड करें और मैसेज पेस्ट करें।")
