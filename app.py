import streamlit as st
import re
from io import BytesIO
from openpyxl import load_workbook

# --- पेज सेटिंग ---
st.set_page_config(page_title="DPR Auto-Filler", layout="wide")
st.title("📊 WhatsApp to Excel: DPR Automation (Format Preserved)")
st.markdown("यह टूल आपकी एक्सेल शीट का फॉर्मेट (रंग, बॉर्डर) खराब नहीं करेगा।")

# --- 1. फाइल अपलोडर ---
uploaded_file = st.file_uploader("अपनी Excel Template यहाँ अपलोड करें (.xlsx)", type=["xlsx"])

# --- 2. टेक्स्ट इनपुट ---
raw_text = st.text_area("WhatsApp Message यहाँ पेस्ट करें:", height=300)

# --- प्रोसेस बटन ---
if st.button("Excel अपडेट करें"):
    if uploaded_file and raw_text:
        try:
            # 1. एक्सेल फाइल को openpyxl से लोड करें (ताकि फॉर्मेट सुरक्षित रहे)
            wb = load_workbook(uploaded_file)
            ws = wb.active  # पहली शीट को सेलेक्ट करें
            
            # 2. डेटा निकालने का लॉजिक (Parsing Logic - Same as before)
            pattern = r"\*(.*?):\*\s*\n• Daily:\s*([\d.]+).*?\n• Monthly:\s*([\d.]+).*?\n• Yearly:\s*([\d.]+)"
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
            
            # 3. एक्सेल की हर लाइन को स्कैन करें और डेटा भरें
            updated_count = 0
            
            # हम मानकर चल रहे हैं:
            # Column B (2) = Material Name
            # Column D (4) = Daily
            # Column E (5) = Monthly
            # Column F (6) = Yearly
            
            # Row 1 से लेकर आखिरी तक चेक करें
            for row in ws.iter_rows(min_row=1, max_col=6):
                # Column B (index 1 in 0-based tuple) में नाम चेक करें
                name_cell = row[1]  
                
                if name_cell.value:
                    cell_value = str(name_cell.value).strip().lower()
                    
                    # अगर नाम हमारे डेटा में है
                    if cell_value in data_map:
                        values = data_map[cell_value]
                        
                        # डेटा अपडेट करें (सीधे सेल्स में लिखें)
                        # row[3] -> Column D
                        # row[4] -> Column E
                        # row[5] -> Column F
                        
                        row[3].value = values['daily']
                        row[4].value = values['monthly']
                        row[5].value = values['yearly']
                        
                        updated_count += 1

            # 4. फाइल सेव करें
            output = BytesIO()
            wb.save(output)
            output.seek(0)  # पॉइंटर को शुरू में लाएं
            
            st.success(f"सफलतापूर्वक! कुल {updated_count} एंट्रीज अपडेट की गईं और फॉर्मेट सुरक्षित है!")
            
            st.download_button(
                label="📥 सही फॉर्मेट वाली फाइल डाउनलोड करें",
                data=output,
                file_name="Updated_DPR_Formatted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error: {e}")
            
    else:
        st.warning("⚠️ कृपया पहले Excel फाइल अपलोड करें और WhatsApp मैसेज पेस्ट करें।")
