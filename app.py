import streamlit as st
import re
import os
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook

# --- पेज सेटिंग और डिजाइन ---
st.set_page_config(page_title="DPR Auto-Filler", layout="wide")
st.title("🚀 Quick DPR Generator")
st.markdown("##### Design & Concept : **K D Mahawar**")
st.markdown("---") 

# --- फाइल पाथ्स (GitHub पर जो आपने अपलोड की हैं) ---
TEMPLATE_FILE = "template.xlsx"
LAST_YEAR_FILE = "last_year_data.xlsx"

st.markdown("बस WhatsApp मैसेज पेस्ट करें, यह पिछले साल का डेटा भी अपने आप उठा लेगा।")

# --- टेक्स्ट इनपुट ---
raw_text = st.text_area("WhatsApp Message यहाँ पेस्ट करें:", height=300)

if st.button("Excel फाइल बनाएँ"):
    if not os.path.exists(TEMPLATE_FILE):
        st.error(f"⚠️ Error: '{TEMPLATE_FILE}' नहीं मिली! इसे GitHub पर अपलोड करें।")
    elif not raw_text:
        st.warning("⚠️ कृपया पहले WhatsApp मैसेज पेस्ट करें।")
    else:
        try:
            # 1. टेम्पलेट लोड करें (Formatting बचाने के लिए openpyxl)
            wb = load_workbook(TEMPLATE_FILE)
            ws = wb.active
            
            # ---------------------------------------------------------
            # PART A: तारीख निकालना और पिछले साल की तारीख बनाना
            # ---------------------------------------------------------
            date_pattern = r"Date:.*?(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})"
            date_match = re.search(date_pattern, raw_text, re.IGNORECASE)
            
            final_date_str = "Unknown"
            lookup_date_str = None
            
            if date_match:
                day, month, year = date_match.groups()
                if len(year) == 2: year = "20" + year
                
                # आज की तारीख (Format: 20-01-2026)
                final_date_str = f"{day.zfill(2)}-{month.zfill(2)}-{year}"
                
                # पिछले साल की तारीख (Format: 20-01-2025)
                last_year = str(int(year) - 1)
                lookup_date_str = f"{day.zfill(2)}-{month.zfill(2)}-{last_year}"
                
                # Excel के हेडर में आज की तारीख अपडेट करें
                for row in ws.iter_rows(min_row=1, max_row=10):
                    for cell in row:
                        if cell.value and isinstance(cell.value, str) and "Date:" in cell.value:
                            cell.value = f"Date: {final_date_str}"
                            break

            # ---------------------------------------------------------
            # PART B: पिछले साल की फाइल से डेटा उठाना (G6, G7)
            # ---------------------------------------------------------
            if lookup_date_str and os.path.exists(LAST_YEAR_FILE):
                try:
                    # पिछले साल की फाइल पढ़ें
                    ly_df = pd.read_excel(LAST_YEAR_FILE)
                    
                    # सुनिश्चित करें कि 'Date' कॉलम सही फॉर्मेट में हो
                    ly_df['Date'] = pd.to_datetime(ly_df['Date']).dt.strftime('%d-%m-%Y')
                    
                    # मैचिंग रो ढूँढें
                    target_row = ly_df[ly_df['Date'] == lookup_date_str]
                    
                    if not target_row.empty:
                        # G6 में Ball Clay की वैल्यू (मान लीजिए कॉलम का नाम 'Ball Clay' है)
                        ws['G6'] = target_row['Ball Clay'].values[0]
                        # G7 में Silica की वैल्यू (मान लीजिए कॉलम का नाम 'Silica' है)
                        ws['G7'] = target_row['Silica'].values[0]
                        st.info(f"✅ पिछले साल का डेटा ({lookup_date_str}) G6 और G7 में भर दिया गया है।")
                    else:
                        st.warning(f"⚠️ पिछले साल की फाइल में {lookup_date_str} की तारीख नहीं मिली।")
                except Exception as ly_e:
                    st.error(f"Last Year File Error: {ly_e}. कृपया कॉलम के नाम 'Date', 'Ball Clay', 'Silica' रखें।")

            # ---------------------------------------------------------
            # PART C: व्हाट्सएप मैसेज से आज का डेटा भरना
            # ---------------------------------------------------------
            pattern = (
                r"\*(.*?)(?::)?\*\s+"
                r"(?:•\s*)?Daily:\s*([\d.]+).*?\n\s*"
                r"(?:•\s*)?Monthly:\s*([\d.]+).*?\n\s*"
                r"(?:•\s*)?Yearly:\s*([\d.]+)"
            )
            matches = re.findall(pattern, raw_text, re.MULTILINE)
            data_map = {m[0].replace(":","").strip().lower(): {'d':float(m[1]),'m':float(m[2]),'y':float(m[3])} for m in matches}

            updated_count = 0
            for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_col=6), 1):
                name_cell = row[1] # Column B
                if name_cell.value:
                    val = str(name_cell.value).strip().lower()
                    if val in data_map:
                        ws.cell(row=row_idx, column=4).value = data_map[val]['d'] # Col D
                        ws.cell(row=row_idx, column=5).value = data_map[val]['m'] # Col E
                        ws.cell(row=row_idx, column=6).value = data_map[val]['y'] # Col F
                        updated_count += 1

            # ---------------------------------------------------------
            # PART D: डाउनलोड
            # ---------------------------------------------------------
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success(f"✅ फाइल तैयार! {updated_count} एंट्रीज अपडेट की गईं।")
            st.download_button(
                label=f"📥 डाउनलोड DPR_{final_date_str}.xlsx",
                data=output,
                file_name=f"DPR_{final_date_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error: {e}")
