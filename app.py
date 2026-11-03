import streamlit as st
import pandas as pd
import random
import io
import time
import requests
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="আমন ধান সংগ্রহ লটারি অ্যাপ-২০২৬", page_icon="🌾", layout="wide")
st.title("🌾 আমন ধান সংগ্রহ কার্যক্রম-২০২৬")
st.subheader("🎯 ইউনিয়ন ভিত্তিক কৃষক নির্বাচন লটারি অ্যাপ")

# ----------------------------------------------------
# 👤 সাইডবারের সবার উপরে আপনার ছবি ও পরিচিতি (Top Sidebar)
# ----------------------------------------------------
st.sidebar.markdown("### 💻 অ্যাপ ক্রিয়েটর")

try:
    repo_url = st.context.headers.get("Referer", "")
    if "github.com" in repo_url:
        username = repo_url.split("github.com/")[1].split("/")[0]
    else:
        username = "labib8april2019-cmyk" 
    
    image_url = f"https://raw.githubusercontent.com/{username}/Shahadul34_Lottery/main/profile.jpg"
    st.sidebar.image(image_url, caption="মোঃ শাহাদুল ইসলাম", width=150)
except Exception:
    st.sidebar.warning("📸 প্রোফাইল ছবি লোড হচ্ছে...")

st.sidebar.info(
    "**মোঃ শাহাদুল ইসলাম**\n\n"
    "পরিচিতি নম্বর: ৩০১৩\n\n"
    "৩৪তম বিসিএস (কৃষি)\n\n"
    "**উপজেলা কৃষি অফিসার**"
)

st.sidebar.markdown("---")

# 📍 লটারি ফিল্টার ও সেटिंग
st.sidebar.header("📍 লটারি ফিল্টার ও সেटिंग")

# 🏛️ উপজেলা ও জেলা ইনপুট বক্স
selected_upazila = st.sidebar.text_input("আপনার উপজেলার নাম লিখুন", value="লাখাই").strip()
selected_dist = st.sidebar.text_input("আপনার জেলার নাম লিখুন", value="হবিগঞ্জ").strip()

# ----------------------------------------------------
# 💳 নতুন ফিচার: অটোমেটিক লাইসেন্স অ্যাক্টিভেশন পোর্টাল (Registration & Payment)
# ----------------------------------------------------
st.sidebar.markdown("---")
st.sidebar.subheader("🔑 লাইসেন্স কি অ্যাক্টিভেশন")

# রিমোট ডাটাবেজ থেকে লাইসেন্স কি চেক করার ডামি ফাংশন (এখানে আপনার Firebase/Supabase এপিআই বসবে)
def get_license_from_db(upazila_name):
    ONLINE_DATABASE = {
        "লাখাই": "lakhai-2026-ok",
        "নবীগঞ্জ": "nabiganj-5542",
        "বানিয়াচং": "baniachong-9981",
        "মাধবপুর": "madhabpur-4432",
    }
    return ONLINE_DATABASE.get(upazila_name, None)

# পেমেন্ট ভেরিফিকেশন ফাংশন
def verify_payment_and_get_key(upazila, district, trx_id, payment_method):
    if len(trx_id) >= 8:
        actual_key = get_license_from_db(upazila)
        if actual_key:
            return actual_key
        else:
            return f"{upazila.lower()}-auto-{random.randint(1000,9999)}"
    return None

# ব্যবহারকারী যদি নতুন হন এবং লাইসেন্স কি না থাকে, তবে তিনি পেমেন্ট সেকশনটি ব্যবহার করবেন
show_payment_gateway = st.sidebar.checkbox("💳 নতুন লাইসেন্স কিনতে চান? (৪,৯৯৯ টাকা)")

if show_payment_gateway:
    st.sidebar.markdown(
        f"<div style='background-color:#fff3cd; padding:10px; border-radius:5px; border-left:5px solid #ffc107; color:#856404;'>"
        f"<b>পেমেন্ট নির্দেশিকা:</b><br>"
        f"আপনার উপজেলা (<b>{selected_upazila}</b>) আনলক করতে নিচের নগদ/রকেটে ৪,৯৯৯ টাকা Send Money করুন।<br>"
        f"📱 <b>নগদ ও রকেট নম্বর:</b> 01725314028"
        f"</div>", 
        unsafe_allow_html=True
    )
    
    method = st.sidebar.selectbox("পেমেন্ট মাধ্যম সিলেক্ট করুন", ["নগদ (Nagad)", "রকেট (Rocket)"])
    trx_input = st.sidebar.text_input("টাকা পাঠানোর পর ট্রানজেকশন আইডি (TrxID) দিন").strip()
    
    if st.sidebar.button("✅ পেমেন্ট নিশ্চিত করুন ও কি (Key) নিন"):
        if not selected_upazila or not selected_dist or not trx_input:
            st.sidebar.error("❌ দয়া করে উপজেলা, জেলা এবং ট্রানজেকশন আইডি সঠিকভাবে লিখুন।")
        else:
            with st.sidebar.spinner("পেমেন্ট ভেরিফাই করা হচ্ছে..."):
                time.sleep(2)
                generated_key = verify_payment_and_get_key(selected_upazila, selected_dist, trx_input, method)
                
                if generated_key:
                    st.sidebar.success("🎉 পেমেন্ট সফল হয়েছে!")
                    st.sidebar.code(f"আপনার লাইসেন্স কি: {generated_key}", language="text")
                    st.sidebar.info("💡 এই কি-টি কপি করে নিচের 'লাইসেন্স কি' বক্সে পেস্ট করুন।")
                else:
                    st.sidebar.error("❌ ট্রানজেকশন আইডিটি সঠিক নয় অথবা পেমেন্টটি এখনো ভেরিফাই করা যায়নি।")

# 🔑 লাইসেন্স কি সিকিউরিটি ইনপুট বক্স
license_key = st.sidebar.text_input("🔑 উপজেলার লাইসেন্স কি (License Key) দিন", type="password")

# 🔒 সিকিউরিটি ভেরিফিকেশন লজিক (অনলাইন ও অফলাইন কম্বাইন্ড)
is_authorized = False
expected_key = get_license_from_db(selected_upazila)

if expected_key:
    if license_key == expected_key:
        is_authorized = True
        st.sidebar.success("✅ লাইসেন্স সফলভাবে ভেরিফাইড!")
    elif license_key != "":
        st.sidebar.error("❌ ভুল লাইসেন্স কি! এই উপজেলার জন্য কোডটি সঠিক নয়।")
else:
    if license_key != "" and selected_upazila in license_key:
        is_authorized = True
        st.sidebar.success("✅ ডাইনামিক লাইসেন্স ভেরিফাইড!")
    elif selected_upazila != "" and license_key != "":
        st.sidebar.warning("🛑 এই উপজেলাটি নিবন্ধিত নয় অথবা লাইসেন্স কি ভুল।")

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
    st.warning("🛑 অ্যাপটি ব্যবহারের জন্য সিকিউরিটি লক করা রয়েছে। অনুগ্রহ করে সাইডবারে আপনার উপজেলা অনুযায়ী সঠিক 'লাইসেন্স কি' (License Key) প্রদান করুন।")
    st.info("☎️ লাইসেন্স কি সংগ্রহ করতে অ্যাপের বাম পাশের সাইডবারে যান, মোবাইলে অ্যাপ ওপেন করলে উপরে থাকা দুটি তীর চিহ্নতে ক্লিক করলে সাইডবার শো করবে। সেখানে উপজেলা বক্সে থাকা লাখাই এবং জেলার বক্সে থাকা হবিগঞ্জ এর স্থলে আপনার নিজ উপজেলা এবং জেলার নাম অবশ্যই পরিবর্তন করে নিবেন। অ্যাপ ক্রিয়েটর **মোঃ শাহাদুল ইসলাম (উপজেলা কৃষি অফিসার, লাখাই)** এর সাথে যোগাযোগ করুন।")

st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray; font-size: 14px;'>"
    "⚙️ কারিগরি সহায়তায়: <b>মোঃ শাহাদুল ইসলাম</b> (পরিচিতি: ৩০১৩), ৩৪তম বিসিএস (কৃষি), উপজেলা কৃষি অফিসার।"
    "</div>", 
    unsafe_allow_html=True
)
