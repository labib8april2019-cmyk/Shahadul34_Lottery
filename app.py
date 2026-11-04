import streamlit as st
import pandas as pd
import random
import io
import time
from datetime import datetime, timedelta
import requests
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="আমন ধান সংগ্রহ লটারি অ্যাপ-২০২৬", page_icon="🌾", layout="wide")
st.title("🌾 আমন ধান সংগ্রহ কার্যক্রম-২০২৬")
st.subheader("🎯 ইউনিয়ন ভিত্তিক কৃষক নির্বাচন লটারি অ্যাপ")

# ----------------------------------------------------
# 👤 সাইডবারের সবার উপরে আপনার পরিচিতি
# ----------------------------------------------------
st.sidebar.markdown("### 💻 অ্যাপ ক্রিয়েটর")

try:
    image_url = "https://raw.githubusercontent.com/labib8april2019-cmyk/Shahadul34_Lottery/main/profile.jpg"
    st.sidebar.image(image_url, caption="মোঃ শাহাদুল ইসলাম", width=150)
except Exception:
    pass

st.sidebar.info(
    "**মোঃ শাহাদুল ইসলাম**\n\n"
    "পরিচিতি নম্বর: ৩০১৩\n\n"
    "৩৪তম বিসিএস (কৃষি)\n\n"
    "**উপজেলা কৃষি অফিসার**"
)

st.sidebar.markdown("---")

# 📍 লটারি ফিল্টার ও সেটিং
st.sidebar.header("📍 লটারি ফিল্টার ও সেটিং")
selected_upazila = st.sidebar.text_input("আপনার উপজেলার নাম লিখুন", value=" ").strip()
selected_dist = st.sidebar.text_input("আপনার জেলার নাম লিখুন", value="হбиগঞ্জ").strip()

# ----------------------------------------------------
# 💾 ফায়ারবেস রিয়েলটাইম ডাটাবেজ ইন্টিগ্রেশন (২ বছরের মেয়াদের জন্য)
# ----------------------------------------------------
FIREBASE_URL = "https://amiikrishok-default-rtdb.firebaseio.com/"

def get_upazila_data(upazila_name):
    if not upazila_name:
        return None
    try:
        url = f"{FIREBASE_URL}unlocked_upazilas/{upazila_name}.json"
        response = requests.get(url, timeout=5)
        if response.status_code == 200 and response.json() is not None:
            return response.json()
    except Exception:
        pass
    return None

def save_or_renew_upazila(upazila_name, trx_id, method, is_renew=False):
    if not upazila_name:
        return False
    try:
        url = f"{FIREBASE_URL}unlocked_upazilas/{upazila_name}.json"
        current_time = datetime.now()
        
        # ২ বছর মেয়াদ যোগ (৭৩০ দিন)
        expiry_time = current_time + timedelta(days=730)
        
        data = {
            "status": "unlocked",
            "trx_id": trx_id,
            "payment_method": method,
            "activation_date": current_time.strftime("%Y-%m-%d %H:%M:%S"),
            "expiry_date": expiry_time.strftime("%Y-%m-%d %H:%M:%S"),
            "expiry_timestamp": expiry_time.timestamp(),
            "type": "renew" if is_renew else "new_activation"
        }
        response = requests.put(url, json=data, timeout=5)
        if response.status_code == 200:
            return True
    except Exception:
        pass
    return False

# ----------------------------------------------------
# 🔑 লাইসেন্স ও সাবস্ক্রিপশন চেক লজিক
# ----------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("🔑 লাইসেন্স স্ট্যাটাস")

is_authorized = False
upazila_info = get_upazila_data(selected_upazila)

if upazila_info:
    expiry_timestamp = upazila_info.get("expiry_timestamp", 0)
    current_timestamp = time.time()
    
    # মেয়াদ চেক করা হচ্ছে
    if current_timestamp < expiry_timestamp:
        expiry_dt = datetime.fromtimestamp(expiry_timestamp)
        days_left = (expiry_dt - datetime.now()).days
        
        st.sidebar.success(f"✅ {selected_upazila} উপজেলা আনলকড!")
        st.sidebar.info(f"⏳ লাইসেন্সের মেয়াদ বাকি: {days_left} দিন\n(মেয়াদের শেষ তারিখ: {expiry_dt.strftime('%d-%b-%Y')})")
        is_authorized = True
    else:
        # ২ বছর পার হয়ে গেলে লক হবে এবং ২০০০ টাকা রিনিউ অপশন দেখাবে
        st.sidebar.error(f"❌ {selected_upazila} উপজেলার লাইসেন্সের মেয়াদ শেষ হয়ে গেছে!")
        st.sidebar.markdown("### 🔄 লাইসেন্স রিনিউ করুন")
        
        st.sidebar.markdown(
            f"<div style='background-color:#f8d7da; padding:10px; border-radius:5px; border-left:5px solid #dc3545; color:#721c24; font-size:13px;'> "
            f"<b>রিনিউয়াল নির্দেশিকা:</b><br>"
            f"পরবর্তী ২ বছরের জন্য অ্যাপটি সচল করতে নিচে দেওয়া নম্বরে <b>২,০০০ টাকা</b> Send Money করুন।<br>"
            f"📱 <b>নগদ/রকেট:</b> 01725314028"
            f"</div>", 
            unsafe_allow_html=True
        )
        
        renew_method = st.sidebar.selectbox("পেমেন্ট মাধ্যম", ["নগদ (Nagad)", "রকেট (Rocket)"], key="renew_m")
        renew_trx = st.sidebar.text_input("রিনিউ ট্রানজেকশন আইডি (TrxID) দিন", key="renew_t").strip()
        
        if st.sidebar.button("🔄 লাইসেন্স রিনিউ নিশ্চিত করুন"):
            if len(renew_trx) < 8:
                st.sidebar.error("❌ সঠিক ট্রানজেকশন আইডি দিন (সর্বনিম্ন ৮ অক্ষর)।")
            else:
                with st.sidebar.spinner("লাইসেন্স নবায়ন করা হচ্ছে..."):
                    if save_or_renew_upazila(selected_upazila, renew_trx, renew_method, is_renew=True):
                        st.sidebar.success("🎉 লাইসেন্স সফলভাবে পরবর্তী ২ বছরের জন্য রিনিউ হয়েছে!")
                        time.sleep(1)
                        st.rerun()
else:
    # একদম নতুন অ্যাক্টিভেশন (৪,৯৯৯ টাকা)
    st.sidebar.warning(f"🛑 {selected_upazila} উপজেলাটি এখনো একটিভেট করা হয়নি।")
    
    show_payment_gateway = st.sidebar.checkbox("💳 নতুন লাইসেন্স একটিভ করুন (৪,৯৯৯ টাকা)")
    if show_payment_gateway:
        st.sidebar.markdown(
            f"<div style='background-color:#fff3cd; padding:10px; border-radius:5px; border-left:5px solid #ffc107; color:#856404; font-size:13px;'>"
            f"<b>অ্যাক্টিভেশন নির্দেশিকা (মেয়াদ ২ বছর):</b><br>"
            f"আপনার উপজেলা আনলক করতে নিচের নম্বরে <b>৪,৯৯৯ টাকা</b> Send Money করুন।<br>"
            f"📱 <b>নগদ/রকেট:</b> 01725314028"
            f"</div>", 
            unsafe_allow_html=True
        )
        
        method = st.sidebar.selectbox("পেমেন্ট মাধ্যম", ["নগদ (Nagad)", "রকেট (Rocket)"], key="act_m")
        trx_input = st.sidebar.text_input("ট্রানজেকশন আইডি (TrxID) দিন", key="act_t").strip()
        
        if st.sidebar.button("✅ অ্যাক্টিভেশন নিশ্চিত করুন"):
            if not trx_input or len(trx_input) < 8:
                st.sidebar.error("❌ অনুগ্রহ করে সঠিক ট্রানজেকশন আইডি দিন।")
            else:
                with st.sidebar.spinner("ডাটাবেজে তথ্য সেভ করা হচ্ছে..."):
                    if save_or_renew_upazila(selected_upazila, trx_input, method, is_renew=False):
                        st.sidebar.success(f"🎉 {selected_upazila} উপজেলা ২ বছরের জন্য সফলভাবে অ্যাক্টিভেট হয়েছে!")
                        time.sleep(1)
                        st.rerun()

# ----------------------------------------------------
# 🌾 মূল অ্যাপের কার্যকারিতা (লকড মেকানিজম)
# ----------------------------------------------------
if is_authorized:
    uploaded_file = st.file_uploader("কৃষকদের ডাটাবেজ (Excel File) আপলোড করুন", type=["xlsx", "xls"])

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        st.success("✅ ডাটাবেজ সফলভাবে আপলোড হয়েছে!")
        df.columns = df.columns.str.strip()

        union_col = None
        possible_names = ['ইউনিয়ন', 'ইউনিয়ন', 'ইউনিয়নের নাম', 'ইউনিয়নেরনাম', 'union', 'union name', 'ইউনি', 'ইউন']
        for col in df.columns:
            col_cleaned = str(col).strip().lower()
            if any(name in col_cleaned for name in possible_names):
                union_col = col
                break

        if union_col is None:
            st.error("❌ আপলোডকৃত এক্সেল ফাইলে ইউনিয়ন সংক্রান্ত কোনো কলাম পাওয়া যায়নি!")
        else:
            union_list = sorted(df[union_col].dropna().unique())
            selected_union = st.sidebar.selectbox("লটারির জন্য ইউনিয়ন নির্বাচন করুন", union_list)
            df_union = df[df[union_col] == selected_union].reset_index(drop=True)
            total_applicants = len(df_union)

            st.info(f"📍 **নির্বাচিত ইউনিয়ন:** {selected_union} | 📝 **মোট আবেদনকারী কৃষক:** {total_applicants} জন")

            with st.expander(f"👀 {selected_union} ইউনিয়নের আবেদনকারীদের তালিকা"):
                st.dataframe(df_union)

            winner_count = st.sidebar.number_input(
                f"কতজন বিজয়ী কৃষক নির্বাচন করবেন? (সর্বোচ্চ {total_applicants} জন)",
                min_value=1,
                max_value=max(1, total_applicants),
                value=min(5, max(1, total_applicants))
            )

            if 'lottery_done' not in st.session_state:
                st.session_state['lottery_done'] = False
                st.session_state['df_winners'] = None
                st.session_state['current_union'] = ""
                st.session_state['total_applicants'] = 0
                st.session_state['winner_count'] = 0

            if st.button(f"🎰 {selected_union} ইউনিয়নের লটারি শুরু করুন"):
                if total_applicants == 0:
                    st.error("এই ইউনিয়নে কোনো আবেদনকারী নেই!")
                else:
                    with st.spinner("🎰 লটারির ড্র চলছে..."):
                        time.sleep(1.2)
                    winner_indices = random.sample(range(total_applicants), winner_count)
                    df_winners = df_union.iloc[winner_indices].reset_index(drop=True)

                    st.session_state['df_winners'] = df_winners
                    st.session_state['current_union'] = selected_union
                    st.session_state['total_applicants'] = total_applicants
                    st.session_state['winner_count'] = winner_count
                    st.session_state['lottery_done'] = True

            if st.session_state['lottery_done'] and st.session_state['current_union'] == selected_union:
                df_winners = st.session_state['df_winners']
                st.balloons()
                st.subheader(f"🎉 লটারিতে নির্বাচিত {selected_union} ইউনিয়নের বিজয়ীদের তালিকা")
                st.dataframe(df_winners)

                # ১. এক্সেল ফাইল জেনারেটর (Buffer)
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_winners.to_excel(writer, index=False, sheet_name='বিজয়ী তালিকা')
                excel_data = excel_buffer.getvalue()

                # ২. ওয়ার্ড ফাইল (.docx) জেনারেটর ফাংশন
                def generate_word_file(winners_dataframe, u_name, total_app, win_num, upazila, district):
                    doc = Document()

                    for section in doc.sections:
                        section.top_margin = Inches(0.75)
                        section.bottom_margin = Inches(0.75)
                        section.left_margin = Inches(0.75)
                        section.right_margin = Inches(0.75)

                    p1 = doc.add_paragraph()
                    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    run1 = p1.add_run(f"আমন ধান সংগ্রহ কার্যক্রম-২০২৬\nউপজেলা: {upazila}, জেলা: {district}।")
                    run1.font.size = Pt(16)
                    run1.font.bold = True
                    p1.add_run().add_break() 

                    run2 = p1.add_run("লটারির মাধ্যমে নির্বাচিত কৃষক বিজয়ীদের অফিসিয়াল তালিকা")
                    run2.font.size = Pt(14)
                    run2.font.bold = True
                    p1.add_run().add_break() 

                    run3 = p1.add_run(f"ইউনিয়ন: {u_name} | মোট আবেদনকারী: {total_app} জন | নির্বাচিত বিজয়ী: {win_num} জন")
                    run3.font.size = Pt(11)

                    doc.add_paragraph()

                    cols_count = len(winners_dataframe.columns) + 1
                    table = doc.add_table(rows=1, cols=cols_count)
                    table.style = 'Table Grid'

                    hdr_cells = table.rows[0].cells
                    hdr_cells[0].text = 'ক্রমিক'
                    hdr_cells[0].paragraphs[0].runs[0].font.bold = True

                    for i, col_name in enumerate(winners_dataframe.columns):
                        hdr_cells[i+1].text = str(col_name)
                        hdr_cells[i+1].paragraphs[0].runs[0].font.bold = True

                    for idx, row in winners_dataframe.iterrows():
                        row_cells = table.add_row().cells
                        row_cells[0].text = str(idx + 1)
                        for i, col_name in enumerate(winners_dataframe.columns):
                            val = row[col_name]
                            row_cells[i+1].text = str(val) if pd.notna(val) else ""

                    p_space = doc.add_paragraph()
                    for _ in range(5):
                        p_space.add_run().add_break()

                    footer_table = doc.add_table(rows=1, cols=3)
                    footer_table.autofit = True
                    
                    cell_0 = footer_table.rows[0].cells[0].paragraphs[0]
                    cell_0.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run_sig1 = cell_0.add_run(f"___________________\nউপজেলা খাদ্য নিয়ন্ত্রক\n{upazila}, {district}।")
                    run_sig1.font.size = Pt(10)
                    run_sig1.font.bold = True

                    cell_1 = footer_table.rows[0].cells[1].paragraphs[0]
                    cell_1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run_sig2 = cell_1.add_run(f"___________________\nউপজেলা কৃষি অফিসার\n{upazila}, {district}।")
                    run_sig2.font.size = Pt(10)
                    run_sig2.font.bold = True

                    cell_2 = footer_table.rows[0].cells[2].paragraphs[0]
                    cell_2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run_sig3 = cell_2.add_run(f"___________________\nউপজেলা নির্বাহী অফিসার\n{upazila}, {district}।")
                    run_sig3.font.size = Pt(10)
                    run_sig3.font.bold = True

                    word_buffer = io.BytesIO()
                    doc.save(word_buffer)
                    return word_buffer.getvalue()

                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label=f"🟢 {selected_union} এর এক্সেল ফাইল ডাউনলোড",
                        data=excel_data,
                        file_name=f"বিজয়ী_তালিকা_{selected_union}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                with col2:
                    with st.spinner('ওয়ার্ড ফাইল তৈরি হচ্ছে...'):
                        word_data = generate_word_file(
                            df_winners,
                            st.session_state['current_union'],
                            st.session_state['total_applicants'],
                            st.session_state['winner_count'],
                            selected_upazila,
                            selected_dist
                        )

                    st.download_button(
                        label=f"🔵 {selected_union} এর অফিসিয়াল ওয়ার্ড ফাইল ডাউনলোড",
                        data=word_data,
                        file_name=f"বিজয়ী_তালিকা_{selected_union}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

                st.success("🎉 আপনার ফাইলগুলো ডাউনলোডের জন্য প্রস্তুত!")
    else:
        st.info("ℹ️ অনুগ্রহ করে প্রথমে কৃষকদের মূল এক্সেল ফাইলটি আপলোড করুন। ফাইলে অবশ্যই ইউনিয়ন কলামটি থাকতে হবে।")
else:
    st.info("☎️ লাইসেন্স একটিভ/নবায়ন করতে উপরে দেওয়া রকেট/নগদ নাম্বারে ফি প্রদান করে TrxID সাবমিট করুন।")

# 🏛️ মূল অ্যাপের নিচে সুন্দর ফুটার ডিজাইন
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray; font-size: 14px;'>"
    "⚙️ কারিগরি সহায়তায়: <b>মোঃ শাহাদুল ইসলাম</b> (পরিচিতি: ৩০১৩), ৩৪তম বিসিএস (কৃষি), উপজেলা কৃষি অফিসার।"
    "</div>", 
    unsafe_allow_html=True
)
