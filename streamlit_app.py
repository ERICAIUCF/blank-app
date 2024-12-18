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
    password = st.text_input("비밀번호", type="password")
    smtp_server = st.text_input("SMTP 서버 주소", placeholder="smtp.naver.com")
    smtp_port = st.number_input("SMTP 포트", min_value=1, value=587)

    # 엑셀 파일 업로드
    st.header("📂 엑셀 파일 업로드")
    uploaded_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # 엑셀 파일 읽기
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        st.write("📊 업로드된 데이터:")
        st.dataframe(df)

        # 조건 설정
        st.header("🔎 필터링 조건")
        filter_column = st.selectbox("조건을 설정할 열을 선택하세요", df.columns)
        filter_value = st.text_input("조건 값 입력", placeholder="예: 완료")

        # 조건 적용
        if st.button("조건 적용"):
            filtered_df = df[df[filter_column] == filter_value]
            st.write("✅ 조건에 맞는 데이터:")
            st.dataframe(filtered_df)

            # 이메일 발송
            st.header("📨 이메일 발송")
            email_column = st.selectbox("이메일 주소가 포함된 열을 선택하세요", df.columns)
            subject = st.text_input("이메일 제목", placeholder="이메일 제목 입력")
            body = st.text_area("이메일 내용", placeholder="이메일 내용을 입력하세요.")

            if st.button("이메일 발송 시작"):
                if from_email and password and smtp_server:
                    results = []
                    for index, row in filtered_df.iterrows():
                        to_email = row[email_column]
                        result = send_email(to_email, subject, body, from_email, password, smtp_server, smtp_port)
                        results.append(result)

                    # 발송 결과 출력
                    st.write("📋 발송 결과:")
                    for res in results:
                        st.write(res)
                else:
                    st.warning("이메일 주소, 비밀번호, SMTP 서버 정보를 입력해주세요.")

if __name__ == "__main__":
    main()
