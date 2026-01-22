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

# --- ALIAS MAPPING ---
NAME_ALIASES = {
    "silica univ lts": "silica sand lts",
    "silica sand": "silica sand lts",
    "cumulative silica": "cumulative silica sand"
}

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
            # PART A: तारीख (Date) हैंडलिंग
            # ---------------------------------------------------------
            date_pattern = r"Date:.*?(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})"
            date_match = re.search(date_pattern, raw_text, re.IGNORECASE)
            
            final_date_str = "Unknown"
            lookup_date_obj = None
            
            if date_match:
                day, month, year = date_match.groups()
                if len(year) == 2: year = "20" + year
                final_date_str = f"{day.zfill(2)}-{month.zfill(2)}-{year}"
                
                # पिछले साल की तारीख
                lookup_date_obj = pd.to_datetime(f"{day}-{month}-{int(year)-1}", dayfirst=True)
                
                # Excel Header Update
                for row in ws.iter_rows(min_row=1, max_row=10):
                    for cell in row:
                        if cell.value and isinstance(cell.value, str) and "Date:" in cell.value:
                            cell.value = f"Date: {final_date_str}"
                            break

            # ---------------------------------------------------------
            # PART B: पिछले साल का डेटा (Last Year Data)
            # ---------------------------------------------------------
            if lookup_date_obj and os.path.exists(LAST_YEAR_FILE):
                try:
                    ly_df = pd.read_excel(LAST_YEAR_FILE)
                    ly_df['Date'] = pd.to_datetime(ly_df['Date'], dayfirst=True)
                    target_row = ly_df[ly_df['Date'] == lookup_date_obj]
                    
                    if not target_row.empty:
                        ws['G6'] = target_row['Ball Clay'].values[0]
                        ws['G7'] = target_row['Silica'].values[0]
                        st.info(f"✅ पिछले साल का डेटा ({lookup_date_obj.strftime('%d-%m-%Y')}) अपडेटेड!")
                except Exception as ly_e:
                    pass

            # ---------------------------------------------------------
            # PART C: व्हाट्सएप डेटा पार्सिंग
            # ---------------------------------------------------------
            # Regex पैटर्न जो स्टार (*) हो या न हो, दोनों को पकड़ेगा
            pattern = (
                r"(?:^|\n)\s*(?:\*)?([^\n\r*]+?)(?::)?(?:\*)?\s*\n\s*" 
                r"(?:•\s*)?Daily:\s*([\d.]+).*?\n\s*"
                r"(?:•\s*)?Monthly:\s*([\d.]+).*?\n\s*"
                r"(?:•\s*)?Yearly:\s*([\d.]+)"
            )
            matches = re.findall(pattern, raw_text, re.MULTILINE)
            
            data_map = {}
            for match in matches:
                raw_name = match[0].strip().lower()
                clean_name = NAME_ALIASES.get(raw_name, raw_name)
                
                data_map[clean_name] = {
                    'd': float(match[1]),
                    'm': float(match[2]),
                    'y': float(match[3])
                }

            # ---------------------------------------------------------
            # PART D: Excel अपडेट (RESET LOGIC ADDED)
            # ---------------------------------------------------------
            updated_count = 0
            
            # Row 4 से शुरू करें (ताकि हेडर खराब न हो)
            for row_idx, row in enumerate(ws.iter_rows(min_row=4, max_col=6), 4):
                name_cell = row[1] # Column B (Name)
                
                if name_cell.value:
                    excel_name = str(name_cell.value).strip().lower()
                    
                    # --- NEW LOGIC: पहले पुराने डेटा को 0 कर दें ---
                    # (ताकि अगर मैसेज में यह नाम न हो, तो पुराना डेटा न दिखे)
                    if "description" not in excel_name and "date" not in excel_name:
                        ws.cell(row=row_idx, column=4).value = 0.0
                        ws.cell(row=row_idx, column=5).value = 0.0
                        ws.cell(row=row_idx, column=6).value = 0.0

                    # अब अगर मैसेज में डेटा है, तो उसे भरें
                    if excel_name in data_map:
                        ws.cell(row=row_idx, column=4).value = data_map[excel_name]['d']
                        ws.cell(row=row_idx, column=5).value = data_map[excel_name]['m']
                        ws.cell(row=row_idx, column=6).value = data_map[excel_name]['y']
                        updated_count += 1

            # ---------------------------------------------------------
            # PART E: डाउनलोड
            # ---------------------------------------------------------
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success(f"✅ अपडेटेड! {updated_count} एंट्रीज भरी गईं (बाकी सब 0 कर दी गईं)।")
            st.download_button(
                label=f"📥 डाउनलोड DPR_{final_date_str}.xlsx",
                data=output,
                file_name=f"DPR_{final_date_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error: {e}")
            
