import streamlit as st
import re
import os
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

# --- पेज सेटिंग ---
st.set_page_config(page_title="DPR Auto-Filler", layout="wide")
st.title("🚀 Quick DPR Generator")
st.markdown("##### Design & Concept : **K D Mahawar**")
st.markdown("---") 

TEMPLATE_FILE = "template.xlsx"
LAST_YEAR_FILE = "last_year_data.xlsx"

# --- HELPER 1: नाम को "नॉर्मल" बनाना ---
def normalize_name(name):
    if not name:
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(name)).lower()

# --- ALIAS MAPPING ---
NAME_ALIASES = {
    "silicaunivlts": "silicasandlts",
    "silicasand": "silicasandlts",
    "cumulativesilica": "cumulativesilicasand"
}

# --- HELPER 2: टेक्स्ट में से सही नंबर निकालना (Trucks को हटाकर) ---
def extract_float(text):
    if not text:
        return 0.0
    
    # 1. सबसे पहले NIL चेक करें
    if "nil" in text.lower():
        return 0.0

    # 2. (NEW LOGIC) ब्रैकेट और उसके अंदर की चीज़ों को हटा दें
    # जैसे: "MT (4 Trucks)" --> "MT " रह जाएगा
    text_no_brackets = re.sub(r'\(.*?\)', '', text)

    # 3. अब बचे हुए हिस्से में नंबर ढूँढें
    match = re.search(r"(\d+(\.\d+)?)", text_no_brackets)
    if match:
        return float(match.group(1))
    
    # अगर ब्रैकेट हटाने के बाद कोई नंबर नहीं बचा, तो 0.0
    return 0.0

raw_text = st.text_area("WhatsApp Message यहाँ पेस्ट करें:", height=300)

if st.button("Excel फाइल बनाएँ"):
    if not os.path.exists(TEMPLATE_FILE):
        st.error(f"⚠️ Error: '{TEMPLATE_FILE}' नहीं मिली!")
    elif not raw_text:
        st.warning("⚠️ कृपया मैसेज पेस्ट करें।")
    else:
        try:
            wb = load_workbook(TEMPLATE_FILE)
            ws = wb.active
            
            # ---------------------------------------------------------
            # PART A: तारीख (Date)
            # ---------------------------------------------------------
            date_pattern = r"Date:.*?(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})"
            date_match = re.search(date_pattern, raw_text, re.IGNORECASE)
            
            final_date_str = "Unknown"
            lookup_date_obj = None
            
            if date_match:
                day, month, year = date_match.groups()
                if len(year) == 2: year = "20" + year
                
                final_date_str = f"{day.zfill(2)}-{month.zfill(2)}-{year}"
                lookup_date_obj = pd.to_datetime(f"{day}-{month}-{int(year)-1}", dayfirst=True)
                
                for row in ws.iter_rows(min_row=1, max_row=10):
                    for cell in row:
                        if cell.value and isinstance(cell.value, str) and "Date:" in cell.value:
                            cell.value = f"Date: {final_date_str}"
                            break

            # ---------------------------------------------------------
            # PART B: पिछले साल का डेटा
            # ---------------------------------------------------------
            if lookup_date_obj and os.path.exists(LAST_YEAR_FILE):
                try:
                    ly_df = pd.read_excel(LAST_YEAR_FILE)
                    ly_df['Date'] = pd.to_datetime(ly_df['Date'], dayfirst=True)
                    target_row = ly_df[ly_df['Date'] == lookup_date_obj]
                    
                    if not target_row.empty:
                        ws['G6'] = target_row['Ball Clay'].values[0]
                        ws['G7'] = target_row['Silica'].values[0]
                        st.info(f"✅ Last Year Data ({lookup_date_obj.strftime('%d-%m-%Y')}) Found!")
                except Exception:
                    pass

            # ---------------------------------------------------------
            # PART C: व्हाट्सएप डेटा (Regex)
            # ---------------------------------------------------------
            pattern = (
                r"(?:^|\n)\s*(?:\*)?([^\n\r*]+?)(?::)?(?:\*)?\s*\n\s*" 
                r"(?:•\s*)?Daily\s*(?::)?\s*(.*?)\n\s*"    
                r"(?:•\s*)?Monthly\s*(?::)?\s*(.*?)\n\s*"  
                r"(?:•\s*)?Yearly\s*(?::)?\s*(.*?)(?:\n|$)"
            )
            
            matches = re.findall(pattern, raw_text, re.MULTILINE | re.IGNORECASE)
            
            data_map = {}
            for match in matches:
                raw_name_norm = normalize_name(match[0])
                final_key = NAME_ALIASES.get(raw_name_norm, raw_name_norm)
                
                # यहाँ extract_float फंक्शन अपना काम करेगा
                data_map[final_key] = {
                    'd': extract_float(match[1]),
                    'm': extract_float(match[2]),
                    'y': extract_float(match[3])
                }

            # ---------------------------------------------------------
            # PART D: Excel अपडेट
            # ---------------------------------------------------------
            updated_count = 0
            
            for row_idx, row in enumerate(ws.iter_rows(min_row=4, max_col=6), 4):
                name_cell = row[1]
                if name_cell.value:
                    excel_name_norm = normalize_name(name_cell.value)
                    
                    # 1. Reset Logic
                    if "description" not in excel_name_norm and "date" not in excel_name_norm:
                        ws.cell(row=row_idx, column=4).value = 0.0
                        ws.cell(row=row_idx, column=5).value = 0.0
                        ws.cell(row=row_idx, column=6).value = 0.0

                    # 2. Update Data
                    if excel_name_norm in data_map:
                        ws.cell(row=row_idx, column=4).value = data_map[excel_name_norm]['d']
                        ws.cell(row=row_idx, column=5).value = data_map[excel_name_norm]['m']
                        ws.cell(row=row_idx, column=6).value = data_map[excel_name_norm]['y']
                        updated_count += 1

            # ---------------------------------------------------------
            # PART E: डाउनलोड
            # ---------------------------------------------------------
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success(f"✅ अपडेटेड! {updated_count} एंट्रीज भरी गईं (Trucks numbers ignored).")
            st.download_button(
                label=f"📥 डाउनलोड DPR_{final_date_str}.xlsx",
                data=output,
                file_name=f"DPR_{final_date_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error: {e}")
            
