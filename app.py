import streamlit as st
import re
import os
from io import BytesIO
from openpyxl import load_workbook

# --- पेज सेटिंग ---
st.set_page_config(page_title="DPR Auto-Filler", layout="wide")

# --- टाइटल और आपका नाम ---
st.title("🚀 Quick DPR Generator")
st.markdown("##### Design & Concept : **K D Mahawar**")
st.markdown("---") 

st.markdown("बस WhatsApp मैसेज पेस्ट करें और फाइल तैयार! (Template ऑटोमेटिक लोड होगा)")

# --- फाइल का नाम (जो GitHub पर है) ---
TEMPLATE_FILE = "template.xlsx"

# --- टेक्स्ट इनपुट ---
raw_text = st.text_area("WhatsApp Message यहाँ पेस्ट करें:", height=300)

# --- प्रोसेस बटन ---
if st.button("Excel फाइल बनाएँ"):
    if not os.path.exists(TEMPLATE_FILE):
        st.error("⚠️ Error: 'template.xlsx' फाइल नहीं मिली! कृपया GitHub पर फाइल चेक करें।")
    elif not raw_text:
        st.warning("⚠️ कृपया पहले WhatsApp मैसेज पेस्ट करें।")
    else:
        try:
            wb = load_workbook(TEMPLATE_FILE)
            ws = wb.active
            
            # -----------------------------------------------
            # PART A: स्मार्ट डेट लॉजिक (Smart Date Logic)
            # -----------------------------------------------
            # यह Regex तारीख के टुकड़ों (Day, Month, Year) को अलग-अलग पकड़ेगा
            # चाहे बीच में / हो या -
            date_pattern = r"Date:.*?(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})"
            date_match = re.search(date_pattern, raw_text, re.IGNORECASE)
            
            final_date_str = "Unknown"
            file_date_str = "Updated"
            
            if date_match:
                day, month, year = date_match.groups()
                
                # अगर साल सिर्फ 2 अंकों का है (जैसे 26), तो उसे 2026 बनाएं
                if len(year) == 2:
                    year = "20" + year
                
                # दिन और महीने को 2 अंकों का बनाएं (जैसे 1 को 01)
                day = day.zfill(2)
                month = month.zfill(2)
                
                # फाइनल फॉर्मेट: DD-MM-YYYY (20-01-2026)
                final_date_str = f"{day}-{month}-{year}"
                file_date_str = final_date_str # फाइल नाम के लिए भी यही इस्तेमाल होगा
                
                # Excel में अपडेट करें
                date_found_in_excel = False
                for row in ws.iter_rows(min_row=1, max_row=10, max_col=10):
                    for cell in row:
                        if cell.value and isinstance(cell.value, str) and "Date:" in cell.value:
                            cell.value = f"Date: {final_date_str}"
                            date_found_in_excel = True
                            break
                    if date_found_in_excel:
                        break
            
            # -----------------------------------------------
            # PART B: डेटा अपडेट करना (Robust Regex)
            # -----------------------------------------------
            pattern = (
                r"\*(.*?)(?::)?\*\s+"                   # Name line
                r"(?:•\s*)?Daily:\s*([\d.]+).*?\n\s*"   # Daily line
                r"(?:•\s*)?Monthly:\s*([\d.]+).*?\n\s*" # Monthly line
                r"(?:•\s*)?Yearly:\s*([\d.]+)"          # Yearly line
            )
            
            matches = re.findall(pattern, raw_text, re.MULTILINE)
            
            data_map = {}
            for match in matches:
                clean_name = match[0].replace(":", "").strip().lower()
                data_map[clean_name] = {
                    'daily': float(match[1]),
                    'monthly': float(match[2]),
                    'yearly': float(match[3])
                }
            
            updated_count = 0
            
            for row in ws.iter_rows(min_row=1, max_col=6):
                name_cell = row[1]
                if name_cell.value:
                    cell_value = str(name_cell.value).strip().lower()
                    if cell_value in data_map:
                        values = data_map[cell_value]
                        row[3].value = values['daily']
                        row[4].value = values['monthly']
                        row[5].value = values['yearly']
                        updated_count += 1

            # -----------------------------------------------
            # PART C: फाइल सेव और डाउनलोड
            # -----------------------------------------------
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            msg_date = f" (Date: {final_date_str})" if date_match else " (Date not found in Msg)"
            st.success(f"✅ काम हो गया! {updated_count} एंट्रीज अपडेट हुईं।{msg_date}")
            
            # फाइल का नाम सेट करें
            final_filename = f"DPR_{file_date_str}.xlsx"
            
            st.download_button(
                label=f"📥 डाउनलोड {final_filename}",
                data=output,
                file_name=final_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error: {e}")
