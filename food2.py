import streamlit as st
import pandas as pd
import chinese_calendar as calendar
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl import load_workbook


def get_holidays(year):
    holidays = []
    start_date = datetime(year, 1, 1)
    end_date = datetime(year, 12, 31)
    current_date = start_date
    while current_date <= end_date:
        if calendar.is_holiday(current_date):
            holidays.append(current_date)
        current_date += timedelta(days=1)
    return pd.to_datetime(holidays)


def get_meal_period(time):
    if pd.to_datetime("07:20:00").time() <= time <= pd.to_datetime("09:00:00").time():
        return "早餐"
    elif pd.to_datetime("11:00:00").time() <= time <= pd.to_datetime("14:00:00").time():
        return "午餐"
    elif pd.to_datetime("17:00:00").time() <= time <= pd.to_datetime("20:00:00").time():
        return "晚餐"
    else:
        return "其他"


def classify_meal(row, total_amount, holiday_days):
    date = row["交易时间"].date()
    weekday = row["交易时间"].weekday()
    meal_period = row["餐费时间段"]
    is_holiday = date in holiday_days or weekday >= 5
    subsidy_limit = 0

    if row["身份"] == "职工":
        if meal_period == "午餐" and weekday < 5:
            subsidy_limit = 25
        elif meal_period in ["午餐", "晚餐"] and is_holiday:
            subsidy_limit = 29
        elif meal_period == "晚餐":
            subsidy_limit = 29
    elif row["身份"] == "学生":
        if meal_period == "早餐" and weekday < 5:
            subsidy_limit = 2
        elif meal_period == "午餐" and weekday < 5:
            subsidy_limit = 25
        elif meal_period in ["午餐", "晚餐"] and is_holiday:
            subsidy_limit = 29
        elif meal_period == "晚餐":
            subsidy_limit = 29

    extra_payment = max(0, total_amount - subsidy_limit) if row["是否最后一笔"] else 0
    return subsidy_limit, extra_payment


def process_data(df, high_temp_days):
    holidays_2024 = get_holidays(2024)
    holidays_2025 = get_holidays(2025)
    holiday_days = pd.to_datetime(high_temp_days).union(pd.to_datetime(holidays_2024)).union(
        pd.to_datetime(holidays_2025))

    df["交易金额"] = df["交易金额"].abs()
    df["交易时间"] = pd.to_datetime(df["交易时间"], errors="coerce")
    df["身份"] = df["帐号"].astype(str).apply(lambda x: "职工" if len(x) == 4 else "学生" if len(x) == 8 else "未知")
    df["餐费时间段"] = df["交易时间"].apply(lambda x: get_meal_period(x.time()))
    df["日期"] = df["交易时间"].dt.date
    df["是否最后一笔"] = df.duplicated(subset=["姓名", "个人编号", "日期", "餐费时间段"], keep="last") == False
    df["总交易金额"] = df.groupby(["姓名", "个人编号", "日期", "餐费时间段"])["交易金额"].transform("sum")
    df[["补贴上限", "自付（元）"]] = df.apply(lambda row: pd.Series(classify_meal(row, row["总交易金额"], holiday_days)), axis=1)

    df_final = df[["姓名", "个人编号", "卡片类型", "交易地点", "卡户部门", "交易时间", "交易金额", "补贴上限", "自付（元）"]]
    return df_final


def save_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output


st.title("餐补计算小程序")

uploaded_file = st.file_uploader("上传 CSV 文件", type=["csv"])
start_date = st.date_input("选择高温假开始日期", value=datetime(2024, 7, 27))
end_date = st.date_input("选择高温假结束日期", value=datetime(2024, 8, 4))

if uploaded_file is not None:
    df = pd.read_csv(uploaded_file, encoding="gbk")
    high_temp_days = pd.date_range(start=start_date, end=end_date)
    df_final = process_data(df, high_temp_days)
    st.write("✅ 数据处理完成！")
    st.dataframe(df_final.head())

    excel_file = save_excel(df_final)
    st.download_button(
        label="📥 下载 Excel 文件",
        data=excel_file,
        file_name="结果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
