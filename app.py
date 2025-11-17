import streamlit as st

DEFAULT_USERS = {"admin": "1234"}  # fallback محلي للاختبار

def get_users():
    try:
        creds = st.secrets["credentials"]
        return dict(creds)
    except Exception:
        return DEFAULT_USERS

USERS = get_users()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""

if not st.session_state.logged_in:
    st.title("تسجيل الدخول")

    username = st.text_input("اسم المستخدم")
    password = st.text_input("كلمة المرور", type="password")

    if st.button("دخول"):
        if username in USERS and password == USERS[username]:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.success(f"مرحبا {username}!")
            st.experimental_rerun()
        else:
            st.error("اسم المستخدم أو كلمة المرور غير صحيحة")
    st.stop()

st.sidebar.write(f"متصل كـ: {st.session_state.username}")
st.title("تطبيق Streamlit - داخل بعد تسجيل الدخول")
st.write("تحكم، اعرض بيانات، ارفع ملفات... أضف اللي تحتاجه هنا.")

uploaded = st.file_uploader("ارفع ملف CSV لاختبار التحليل", type=["csv"])
if uploaded:
    import pandas as pd
    df = pd.read_csv(uploaded)
    st.write("Preview:")
    st.dataframe(df.head())
    st.write("معلومات عامة:")
    st.write(df.describe(include='all'))
