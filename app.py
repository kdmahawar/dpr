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

# --- प्रोसेस करने का बटन ---
if st.button("Excel अपडेट करें") and uploaded_file and raw_text:
    try:
        # एक्सेल फाइल लोड करें (Header को skip करते हुए क्योंकि आपका फॉर्मेट कॉम्प्लेक्स है)
        # हम सीधे इंडेक्स के आधार पर काम करेंगे
        df = pd.read_excel(uploaded_file, header=None)
        
        # --- डेटा निकालने का लॉजिक (Parsing Logic) ---
        # यह पैटर्न आपके मैसेज फॉर्मेट के हिसाब से बनाया गया है
        # Group 1: Name, Group 2: Daily, Group 3: Monthly, Group 4: Yearly
        pattern = r"\*(.*?):\*\s*\n• Daily:\s*([\d.]+).*?\n• Monthly:\s*([\d.]+).*?\n• Yearly:\s*([\d.]+)"
        
        matches = re.findall(pattern, raw_text, re.MULTILINE)
        
        # मैसेज के डेटा को एक डिक्शनरी में सेव करें ताकि खोजने में आसानी हो
        data_map = {}
        for match in matches:
            clean_name = match[0].strip().lower()  # नाम को छोटा (lowercase) करें मैचिंग के लिए
            data_map[clean_name] = {
                'daily': float(match[1]),
                'monthly': float(match[2]),
                'yearly': float(match[3])
            }
            
        # --- एक्सेल में डेटा भरना ---
        # हम एक्सेल की हर लाइन चेक करेंगे
        # मान रहे हैं: Col B (Index 1) में नाम है, Col D (3) Daily, Col E (4) Monthly, Col F (5) Yearly
        
        updated_count = 0
        
        for index, row in df.iterrows():
            cell_value = str(row[1]) # कॉलम B (नाम)
            
            if pd.notna(cell_value):
                # एक्सेल के नाम को भी साफ और छोटा करें
                excel_name_clean = cell_value.strip().lower()
                
                # चेक करें कि क्या यह नाम हमारे व्हाट्सएप डेटा में है?
                if excel_name_clean in data_map:
                    values = data_map[excel_name_clean]
                    
                    # डेटा अपडेट करें
                    df.at[index, 3] = values['daily']   # Column D (Daily)
                    df.at[index, 4] = values['monthly'] # Column E (Monthly)
                    df.at[index, 5] = values['yearly']  # Column F (Yearly)
                    
                    updated_count += 1

        # --- फाइल सेव और डाउनलोड ---
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # हेडर नहीं लिख रहे क्योंकि हम पूरी शीट को वैसे का वैसा वापस दे रहे हैं
            df.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
            
        st.success(f"कुल {updated_count} एंट्रीज अपडेट की गईं!")
        
        st.download_button(
            label="📥 नई Excel फाइल डाउनलोड करें",
            data=output.getvalue(),
            file_name="Updated_DPR.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")

elif st.button("Excel अपडेट करें"):
    st.warning("कृपया पहले फाइल अपलोड करें और मैसेज पेस्ट करें।")
