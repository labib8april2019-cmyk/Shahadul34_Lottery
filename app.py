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
    st.warning("🛑 আমন ধান সংগ্রহ লটারি অ্যাপটি সিকিউরিটি লক করা রয়েছে। ফি প্রদানের জন্য নির্দেশাবলী অনুসরণ করুন।")
    st.caption("☎️ যেকোনো সমস্যা বা লাইসেন্স সংক্রান্ত সহযোগিতার জন্য ক্রিয়েটর এর সাথে যোগাযোগ করুন।")

st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray; font-size: 14px;'>"
    "⚙️ কারিগরি সহায়তায়: উপজেলা কৃষি অফিসার।"
    "</div>", 
    unsafe_allow_html=True
)
