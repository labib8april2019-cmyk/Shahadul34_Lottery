import os code_content = """
import streamlit as st
import pandas as pd
import random
import io
import time
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="আমন ধান সংগ্রহ লটারি অ্যাপ-২০২৬", page_icon="🌾", layout="wide")
st.title("🌾 আমন ধান সংগ্রহ কার্যক্রম-২০২৬")
st.subheader("🎯 ইউনিয়ন ভিত্তিক কৃষক নির্বাচন লটারি অ্যাপ")

st.sidebar.markdown("### 💻 অ্যাপ ক্রিয়েটর")

try:
    repo_url = st.context.headers.get("Referer", "")
    if "github.com" in repo_url:
        username = repo_url.split("://://github.com").split("/")
    else:
        username = "labib8april2019-cmyk" 
    
    image_url = f"https://githubusercontent.com{username}/Shahadul34_Lottery/main/profile.jpg"
    st.sidebar.image(image_url, caption="মোঃ শাহাদুল ইসলাম", width=150)
except Exception:
    st.sidebar.warning("📸 প্রোফাইল ছবি লোড হচ্ছে...")

st.sidebar.info(
    "**মোঃ শাহাদুল ইসলাম**\\n\\n"
    "পরিচিতি নম্বর: ৩০১৩\\n\\n"
    "৩৪তম বিসিএস (কৃষি)\\n\\n"
    "**উপজেলা কৃষি অফিসার**"
)

st.sidebar.markdown("---")
st.sidebar.header("📍 লটারি ফিল্টার ও সেटिंग")

selected_upazila = st.sidebar.text_input("আপনার উপজেলার নাম লিখুন", value="লাখাই").strip()
selected_dist = st.sidebar.text_input("আপনার জেলার নাম লিখুন", value="হбиগঞ্জ").strip()

VALID_TRANSACTIONS = ["txnid12345", "nagad01725314028", "rocket01725314028"]
LICENSE_DATABASE = {
    "লাখাই": "lakhai-2026-ok",
    "নবীগঞ্জ": "nabiganj-5542",
    "বানিয়াচং": "baniachong-9981",
    "ماধবপুর": "madhabpur-4432",
}

is_authorized = False
st.sidebar.markdown("---")
st.sidebar.markdown("### 💳 অ্যাপ আনলক ও পেমেন্ট ভেরিফিকেশন")
verification_method = st.sidebar.radio("ভেরিফিকেশন পদ্ধতি সিলেক্ট করুন", ["ট্রানজেকশন আইডি (TxnID)", "গোপন লাইসেন্স কি"])

if verification_method == "ট্রানজেকশন আইডি (TxnID)":
    user_txnid = st.sidebar.text_input("📝 পেমেন্টের Transaction ID (TxnID) দিন").strip().lower()
    if user_txnid in VALID_TRANSACTIONS:
        is_authorized = True
        st.sidebar.success("✅ পেমেন্ট সফলভাবে ভেরিফাইড! অ্যাপ আনলক হয়েছে।")
    elif user_txnid != "":
        st.sidebar.error("❌ এই ট্রানজেকশন আইডিটি সঠিক নয় বা এখনও অনুমোদিত হয়নি।")
else:
    license_key = st.sidebar.text_input("🔑 উপজেলার গোপন লাইসেন্স কি দিন", type="password")
    if selected_upazila in LICENSE_DATABASE and license_key == LICENSE_DATABASE[selected_upazila]:
        is_authorized = True
        st.sidebar.success("✅ লাইসেন্স সফলভাবে ভেরিফাইড!")
    elif license_key != "":
        st.sidebar.error("❌ ভুল লাইসেন্স কি!")

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

                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_winners.to_excel(writer, index=False, sheet_name='বিজয়ী তালিকা')
                excel_data = excel_buffer.getvalue()

                def generate_word_file(winners_dataframe, u_name, total_app, win_num, upazila, district):
                    doc = Document()

                    for section in doc.sections:
                        section.top_margin = Inches(0.75)
                        section.bottom_margin = Inches(0.75)
                        section.left_margin = Inches(0.75)
                        section.right_margin = Inches(0.75)

                    p1 = doc.add_paragraph()
                    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                    run1 = p1.add_run(f"আমন ধান সংগ্রহ কার্যক্রম-২০২৬\\nউপজেলা: {upazila}, জেলা: {district}।")
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
                    run_sig1 = cell_0.add_run(f"___________________\\nউপজেলা খাদ্য নিয়ন্ত্রক\\n{upazila}, {district}।")
                    run_sig1.font.size = Pt(10)
                    run_sig1.font.bold = True

                    cell_1 = footer_table.rows[0].cells[1].paragraphs[0]
                    cell_1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run_sig2 = cell_1.add_run(f"___________________\\nউপজেলা কৃষি অফিসার\\n{upazila}, {district}।")
                    run_sig2.font.size = Pt(10)
                    run_sig2.font.bold = True

                    cell_2 = footer_table.rows[0].cells[2].paragraphs[0]
                    cell_2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run_sig3 = cell_2.add_run(f"___________________\\nউপজেলা নির্বাহী অফিসার\\n{upazila}, {district}।")
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
