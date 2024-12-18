import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# 이메일 전송 함수
def send_email(to_email, subject, body, from_email, password, smtp_server, smtp_port):
    try:
        # SMTP 서버 연결
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # TLS 보안 연결
        server.login(from_email, password)

        # 이메일 내용 설정
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # 이메일 전송
        server.sendmail(from_email, to_email, msg.as_string())
        server.quit()
        return f"✅ {to_email}에게 이메일 전송 성공!"
    except Exception as e:
        return f"❌ {to_email}에게 이메일 전송 실패! 오류: {e}"

# Streamlit UI
def main():
    st.title("📧 Excel 기반 조건부 이메일 자동 발송")

    # 발송자 정보 설정
    st.header("🔐 발송자 정보 입력")
    from_email = "sa5353@hanyang.ac.kr"  # 발송자 이메일 고정
    st.write(f"발송자 이메일: **{from_email}**")
    password = st.text_input("비밀번호를 입력하세요", type="password")

    # Gmail SMTP 서버 설정 (고정)
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    # 엑셀 파일 업로드
    st.header("📂 엑셀 파일 업로드")
    uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # 엑셀 파일 읽기
        df = pd.read_excel(uploaded_file, engine="openpyxl")

