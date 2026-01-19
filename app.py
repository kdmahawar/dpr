import streamlit as st
import pandas as pd
import re
from io import BytesIO

# --- पेज सेटिंग ---
st.set_page_config(page_title="DPR Auto-Filler", layout="wide")
st.title("📊 WhatsApp to Excel: DPR Automation")
st.markdown("अपना व्हाट्सएप मैसेज पेस्ट करें और ऑटो-अपडेटेड एक्सेल फाइल डाउनलोड करें।")

# --- 1. फाइल अपलोडर ---
uploaded_file = st.file_uploader("अपनी Excel Template यहाँ अपलोड करें (.xlsx)", type=["xlsx"])

# --- 2. टेक्स्ट इनपुट ---
raw_text = st.text_area("WhatsApp Message यहाँ पेस्ट करें:", height=300)

# --- बटन और लॉजिक (सुधारा गया हिस्सा) ---
if st.button("Excel अपडेट करें"):  # बटन अब सिर्फ एक बार है
    if uploaded_file and raw_text:
        try:
            # एक्सेल फाइल लोड करें
            df = pd.read_excel(uploaded_file, header=None)
            
            # --- डेटा निकालने का लॉजिक (Parsing Logic) ---
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
                
            # --- एक्सेल में डेटा भरना ---
            updated_count = 0
            
            for index, row in df.iterrows():
                cell_value = str(row[1]) # कॉलम B (नाम)
                
                if pd.notna(cell_value):
                    excel_name_clean = cell_value.strip().lower()
                    
                    if excel_name_clean in data_map:
                        values = data_map[excel_name_clean]
                        
                        # डेटा अपडेट करें (Columns D, E, F -> Index 3, 4, 5)
                        df.at[index, 3] = values['daily']
                        df.at[index, 4] = values['monthly']
                        df.at[index, 5] = values['yearly']
                        
                        updated_count += 1

            # --- फाइल सेव और डाउनलोड ---
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
                
            st.success(f"सफलतापूर्वक! कुल {updated_count} एंट्रीज अपडेट की गईं!")
            
            st.download_button(
                label="📥 नई Excel फाइल डाउनलोड करें",
                data=output.getvalue(),
                file_name="Updated_DPR.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error: {e}")
            
    else:
        # अगर फाइल या टेक्स्ट नहीं है तो यह मैसेज दिखेगा
        st.warning("⚠️ कृपया पहले Excel फाइल अपलोड करें और WhatsApp मैसेज पेस्ट करें।")
        
