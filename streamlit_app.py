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
        server.starttls()  # TLS 보안 연결 시작
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

    # SMTP 설정 입력
    st.header("🔐 이메일 설정")
    from_email = st.text_input("이메일 주소 (발송자)", placeholder="example@naver.com")
    password = st.text_input("비밀번호
