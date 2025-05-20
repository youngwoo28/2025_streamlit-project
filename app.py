import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from wordcloud import WordCloud
import os, glob, re, warnings

# 경고 무시
warnings.filterwarnings("ignore", message="Workbook contains no default style")

# 페이지 설정
st.set_page_config(page_title="사이버 해킹 데이터 대시보드", layout="wide")

# 한글 폰트 직접 설정 (Nanum Gothic)
try:
    base_dir = os.getcwd()
    font_dir = os.path.join(base_dir, "fonts", "Nanum_Gothic")
    font_path = os.path.join(font_dir, "NanumGothic-Regular.ttf")

    fm.fontManager.addfont(font_path)  # 폰트 등록
    font_prop = fm.FontProperties(fname=font_path)
    plt.rcParams['font.family'] = font_prop.get_name()
    print("설정된 폰트:", font_prop.get_name())
except Exception as e:
    print("폰트 설정 에러:", e)

def load_news_data():
    data_dir = os.path.join(os.getcwd(), "data")
    # 영어 파일명에 맞게 glob 패턴 변경
    files = glob.glob(os.path.join(data_dir, "news_hackingx_*.xlsx"))
    print("검색된 뉴스기사 파일:")
    for f in files:
        print(f)
    news = {}
    for path in files:
        filename = os.path.basename(path)
        m = re.search(r"(\d{4})", filename)
        if m:
            year = m.group(1)
            try:
                df = pd.read_excel(path)
                news[year] = df
            except Exception as e:
                print(f"파일 읽기 실패 {filename}: {e}")
    return news

def load_hacking_data():
    data_dir = os.path.join(os.getcwd(), "data")
    path = os.path.join(data_dir, "hacking.xlsx")
    if not os.path.isfile(path):
        return None
    df0 = pd.read_excel(path, sheet_name=0)
    df = df0.iloc[2:5].copy().reset_index(drop=True)
    df.rename(columns={df0.columns[0]: '사고유형'}, inplace=True)
    df.set_index('사고유형', inplace=True)
    years = [str(y) for y in range(2015, 2025)]
    df.columns = years
    df = df.applymap(lambda x: int(str(x).replace(',', '')))
    return df

# 데이터 로딩
news_data = load_news_data()
hacking_df = load_hacking_data()

# 사이드바 UI
st.sidebar.title("분석 옵션")
analysis_type = st.sidebar.selectbox("분석 유형 선택", ["뉴스 기사 시각화", "해킹 사고 분석", "원본 데이터"])
show_raw = st.sidebar.checkbox("🔍 원본 데이터 보기")

# 메인 타이틀
st.title("사이버 해킹 데이터 대시보드")

if analysis_type == "뉴스 기사 시각화":
    st.header("해킹 뉴스 기사 수 (연도별)")
    if news_data:
        years = sorted(news_data.keys())
        counts = {yr: len(df) for yr, df in news_data.items()}
        count_df = pd.DataFrame.from_dict(counts, orient='index', columns=['기사 수']).sort_index()
        fig, ax = plt.subplots(figsize=(10,5))
        ax.plot(count_df.index, count_df['기사 수'], marker='o')
        ax.set_xlabel('연도')
        ax.set_ylabel('기사 수')
        ax.set_title('연도별 해킹 뉴스 기사 수')
        ax.grid(True)

        # x축 tick label 폰트 적용
        for label in ax.get_xticklabels():
            label.set_fontproperties(font_prop)

        st.pyplot(fig)

        st.subheader("제목 워드클라우드")
        year_sel = st.selectbox("연도 선택", years, key="wc_year")
        if st.button("생성", key="wc_btn"):
            texts = news_data[year_sel]['제목'].dropna().astype(str).tolist()
            words = re.findall(r"[가-힣]{2,}|[A-Za-z0-9]{2,}", " ".join(texts))
            freq = pd.Series(words).value_counts().head(100).to_dict()
            wc = WordCloud(
                font_path=font_path,
                width=800,
                height=400,
                background_color='white'
            ).generate_from_frequencies(freq)
            st.image(wc.to_array(), use_column_width=True)
    else:
        st.warning("뉴스 데이터가 없습니다.")

elif analysis_type == "해킹 사고 분석":
    st.header("해킹 사고 유형별 건수 추이 (2020~2024)")
    if hacking_df is not None:
        df_plot = hacking_df.loc[:, '2020':'2024']
        fig, ax = plt.subplots(figsize=(10,5))
        for idx in df_plot.index:
            ax.plot(df_plot.columns, df_plot.loc[idx], marker='o', label=idx)
        ax.set_xlabel('연도')
        ax.set_ylabel('건수')
        ax.set_title('사고 유형별 연도별 건수')
        ax.legend(bbox_to_anchor=(1.05,1), loc='upper left')
        ax.grid(True)

        # x축 tick label 폰트 및 회전 적용
        plt.xticks(rotation=45, ha='right')
        for label in ax.get_xticklabels():
            label.set_fontproperties(font_prop)

        st.pyplot(fig)

        st.subheader("코로나 전·후 총건수 비교")
        covid = hacking_df.loc[:, '2020':'2022'].sum(axis=1)
        post = hacking_df.loc[:, '2023':'2024'].sum(axis=1)
        comp = pd.DataFrame({'코로나(20-22)': covid, '완화(23-24)': post})
        fig2, ax2 = plt.subplots(figsize=(8,4))
        comp.plot(kind='bar', ax=ax2)

        plt.xticks(rotation=45, ha='right')
        for label in ax2.get_xticklabels():
            label.set_fontproperties(font_prop)

        ax2.set_ylabel('총건수')
        ax2.set_title('시기별 해킹 사고 총건수 비교')
        st.pyplot(fig2)
    else:
        st.warning("hacking.xlsx 파일을 찾을 수 없습니다.")

elif analysis_type == "원본 데이터":
    st.header("원본 데이터")
    st.subheader("뉴스 데이터")
    if news_data:
        for yr, df in news_data.items():
            st.write(f"### {yr}년 뉴스")
            st.dataframe(df)
    else:
        st.warning("뉴스 데이터가 없습니다.")
    st.subheader("해킹 사고 데이터")
    if hacking_df is not None:
        st.dataframe(hacking_df)
    else:
        st.warning("해킹 사고 데이터가 없습니다.")

# 원본 데이터 보기 토글에 따른 상세 출력
if show_raw:
    st.markdown("---")
    st.header("원본 데이터 상세 보기")

    st.subheader("뉴스 데이터 상세")
    if news_data:
        for yr, df in news_data.items():
            st.write(f"### {yr}년 뉴스")
            st.dataframe(df)
    else:
        st.warning("뉴스 데이터가 없습니다.")

    st.subheader("해킹 사고 데이터 상세")
    if hacking_df is not None:
        st.dataframe(hacking_df)
    else:
        st.warning("해킹 사고 데이터가 없습니다.")
