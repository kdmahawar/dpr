import streamlit as st
import re
import os
from io import BytesIO
from openpyxl import load_workbook

# --- पेज सेटिंग ---
st.set_page_config(page_title="DPR Auto-Filler", layout="wide")
st.title("🚀 Quick DPR Generator")
st.markdown("बस WhatsApp मैसेज पेस्ट करें और फाइल तैयार! (Template ऑटोमेटिक लोड होगा)")

# --- फाइल का नाम (जो आपने GitHub पर अपलोड की है) ---
TEMPLATE_FILE = "template.xlsx"

# --- टेक्स्ट इनपुट ---
raw_text = st.text_area("WhatsApp Message यहाँ पेस्ट करें:", height=300)

# --- प्रोसेस बटन ---
if st.button("Excel फाइल बनाएँ"):
    # चेक करें कि GitHub पर template फाइल है या नहीं
    if not os.path.exists(TEMPLATE_FILE):
        st.error("⚠️ Error: 'template.xlsx' फाइल नहीं मिली! कृपया इसे GitHub पर अपलोड करें।")
    elif not raw_text:
        st.warning("⚠️ कृपया पहले WhatsApp मैसेज पेस्ट करें।")
    else:
        try:
            # 1. GitHub (सर्वर) से सीधे टेम्पलेट फाइल लोड करें
            wb = load_workbook(TEMPLATE_FILE)
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
                
                # Excel की पहली 10 लाइनों में "Date:" ढूंढकर अपडेट करें
                for row in ws.iter_rows(min_row=1, max_row=10, max_col=10):
                    for cell in row:
                        if cell.value and isinstance(cell.value, str) and "Date:" in cell.value:
                            cell.value = f"Date: {new_date}"
                            date_found = True
                            break
                    if date_found:
                        break
            
            # -----------------------------------------------
            # PART B: डेटा (Figures) अपडेट करना
            # -----------------------------------------------
            # यह पैटर्न बुलेट (•), स्पेस और कॉलन (:) की सभी गलतियों को संभाल लेगा
            pattern = (
                r"\*(.*?)(?::)?\*\s+"                   # Name line
                r"(?:•\s*)?Daily:\s*([\d.]+).*?\n\s*"   # Daily line
                r"(?:•\s*)?Monthly:\s*([\d.]+).*?\n\s*" # Monthly line
                r"(?:•\s*)?Yearly:\s*([\d.]+)"          # Yearly line
            )
            
            matches = re.findall(pattern, raw_text, re.MULTILINE)
            
            # डेटा मैप तैयार करना
            data_map = {}
            for match in matches:
                clean_name = match[0].replace(":", "").strip().lower()
                data_map[clean_name] = {
                    'daily': float(match[1]),
                    'monthly': float(match[2]),
                    'yearly': float(match[3])
                }
            
            updated_count = 0
            
            # Excel की पंक्तियाँ स्कैन करें
            for row in ws.iter_rows(min_row=1, max_col=6):
                name_cell = row[1]  # Column B
                
                if name_cell.value:
                    cell_value = str(name_cell.value).strip().lower()
                    
                    if cell_value in data_map:
                        values = data_map[cell_value]
                        
                        row[3].value = values['daily']   # Column D
                        row[4].value = values['monthly'] # Column E
                        row[5].value = values['yearly']  # Column F
                        
                        updated_count += 1

            # -----------------------------------------------
            # PART C: फाइल डाउनलोड के लिए तैयार करना
            # -----------------------------------------------
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            msg_date = f" (Date: {new_date})" if date_found else ""
            st.success(f"✅ काम हो गया! {updated_count} एंट्रीज अपडेट हुईं।{msg_date}")
            
            file_name_date = new_date.replace('/', '-') if new_date != "Unknown" else "Updated"
            
            st.download_button(
                label="📥 डाउनलोड Excel फाइल",
                data=output,
                file_name=f"DPR_{file_name_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error: {e}")
