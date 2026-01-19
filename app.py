import streamlit as st
import re
from io import BytesIO
from openpyxl import load_workbook

# --- पेज सेटिंग ---
st.set_page_config(page_title="DPR Auto-Filler", layout="wide")
st.title("📊 WhatsApp to Excel: DPR Automation")
st.markdown("यह टूल स्पेस (Space) की गलतियों को भी अपने आप ठीक कर लेगा।")

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
            
            # --- 2. डेटा निकालने का लॉजिक (UPDATED) ---
            # बदलाव: \s* जोड़ दिया गया है ताकि स्पेस हो या न हो, दोनों चलेगा।
            pattern = r"\*(.*?):\*\s*\n•\s*Daily:\s*([\d.]+).*?\n•\s*Monthly:\s*([\d.]+).*?\n•\s*Yearly:\s*([\d.]+)"
            
            matches = re.findall(pattern, raw_text, re.MULTILINE)
            
            # डेटा मैप तैयार करना
            data_map = {}
            for match in matches:
                clean_name = match[0].strip().lower()
                data_map[clean_name] = {
                    'daily': float(match[1]),
                    'monthly': float(match[2]),
                    'yearly': float(match[3])
                }
            
            # 3. एक्सेल अपडेट करना
            updated_count = 0
            
            for row in ws.iter_rows(min_row=1, max_col=6):
                name_cell = row[1]  # Column B
                
                if name_cell.value:
                    cell_value = str(name_cell.value).strip().lower()
                    
                    if cell_value in data_map:
                        values = data_map[cell_value]
                        
                        # डेटा अपडेट करें
                        row[3].value = values['daily']   # Column D
                        row[4].value = values['monthly'] # Column E
                        row[5].value = values['yearly']  # Column F
                        
                        updated_count += 1

            # 4. फाइल सेव करें
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.success(f"सफलतापूर्वक! कुल {updated_count} एंट्रीज अपडेट की गईं। (Abhiraj वाली एंट्री भी चेक कर लें!)")
            
            st.download_button(
                label="📥 अपडेटेड Excel फाइल डाउनलोड करें",
                data=output,
                file_name="Updated_DPR_19_Jan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error: {e}")
            
    else:
        st.warning("⚠️ कृपया पहले Excel फाइल अपलोड करें और मैसेज पेस्ट करें।")
