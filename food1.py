import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
import chinese_calendar as calendar
from openpyxl.utils import get_column_letter


# 获取指定年份的法定节假日
def get_holidays(year):
    holidays = []
    current_date = datetime(year, 1, 1)
    end_date = datetime(year, 12, 31)

    while current_date <= end_date:
        if calendar.is_holiday(current_date):
            holidays.append(current_date.date())
        current_date += timedelta(days=1)

    return set(holidays)


# 定义餐补时间段
def get_meal_period(t):
    if t is None or pd.isna(t):
        return "其他"

    breakfast_start = pd.to_datetime("07:20:00").time()
    breakfast_end = pd.to_datetime("09:00:00").time()
    lunch_start = pd.to_datetime("11:00:00").time()
    lunch_end = pd.to_datetime("14:00:00").time()
    dinner_start = pd.to_datetime("17:00:00").time()
    dinner_end = pd.to_datetime("20:00:00").time()

    if breakfast_start <= t <= breakfast_end:
        return "早餐"
    elif lunch_start <= t <= lunch_end:
        return "午餐"
    elif dinner_start <= t <= dinner_end:
        return "晚餐"
    else:
        return "其他"


# 解析日期输入
def parse_date_set(text):
    dates = set()
    invalid = []

    for part in (text or "").split(","):
        part = part.strip()
        if not part:
            continue
        try:
            dates.add(datetime.strptime(part, "%Y-%m-%d").date())
        except ValueError:
            invalid.append(part)

    return dates, invalid


# 根据人员类别、时段、是否工作日/节假日，确定补贴上限
def get_max_subsidy(person_type, meal_period, workday, is_holiday):
    if person_type == "职工":
        if meal_period == "早餐":
            return 0.0
        elif meal_period in ["午餐", "晚餐"] and is_holiday:
            return 29.0
        elif meal_period == "午餐" and workday:
            return 25.0
        elif meal_period in ["午餐", "晚餐"]:
            return 29.0
        else:
            return 0.0

    elif person_type == "研究生":
        if meal_period == "早餐" and workday:
            return 2.0
        elif meal_period in ["午餐", "晚餐"] and is_holiday:
            return 29.0
        elif meal_period == "午餐" and workday:
            return 25.0
        elif meal_period in ["午餐", "晚餐"]:
            return 29.0
        else:
            return 0.0

    return 0.0


# 计算单个分组的餐补
def calculate_subsidy_group(group, overtime_dates, holiday_and_high_temp_days):
    group = group.copy()
    subsidy_used = 0.0

    # 因为已经按“餐费时间段”分组，所以组内取第一条即可
    meal_period = group["餐费时间段"].iloc[0]

    for idx, row in group.iterrows():
        amount = float(row["交易金额"])

        # 超市不参与餐补
        if row["交易地点"] == "超市":
            group.at[idx, "餐补金额"] = 0.0
            group.at[idx, "自付（元）"] = round(amount, 2)
            group.at[idx, "早餐（元）"] = 0.0
            group.at[idx, "工作餐（元）"] = 0.0
            group.at[idx, "加班餐（元）"] = 0.0
            continue

        trade_time = row["交易时间"]
        date = trade_time.date()
        weekday = trade_time.weekday()

        workday = (weekday < 5) or (date in overtime_dates)
        is_holiday = (date in holiday_and_high_temp_days) and (date not in overtime_dates)

        max_subsidy = get_max_subsidy(
            person_type=row["人员类别"],
            meal_period=meal_period,
            workday=workday,
            is_holiday=is_holiday,
        )

        available_subsidy = max(0.0, max_subsidy - subsidy_used)
        subsidy_given = min(amount, available_subsidy)
        subsidy_used += subsidy_given

        group.at[idx, "餐补金额"] = round(subsidy_given, 2)
        group.at[idx, "自付（元）"] = round(amount - subsidy_given, 2)

        # 先清零，再按类别写入
        group.at[idx, "早餐（元）"] = 0.0
        group.at[idx, "工作餐（元）"] = 0.0
        group.at[idx, "加班餐（元）"] = 0.0

        if meal_period == "早餐":
            group.at[idx, "早餐（元）"] = round(subsidy_given, 2)
        elif meal_period == "午餐" and (not is_holiday) and workday:
            group.at[idx, "工作餐（元）"] = round(subsidy_given, 2)
        elif meal_period in ["午餐", "晚餐"]:
            group.at[idx, "加班餐（元）"] = round(subsidy_given, 2)

    return group


# 读取 CSV，自动尝试编码
def read_csv_with_fallback(uploaded_file):
    last_error = None
    for enc in ("gbk", "gb18030", "utf-8-sig", "utf-8"):
        try:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding=enc, low_memory=False)
            return df, enc
        except Exception as e:
            last_error = e

    raise last_error


# 核心处理逻辑
def process_dataframe(raw_df, holiday_year, overtime_dates, high_temp_days):
    required_columns = [
        "人员类别", "姓名", "个人编号", "卡片类型",
        "交易地点", "交易金额", "交易时间", "卡户部门", "交易类型"
    ]
    missing = [col for col in required_columns if col not in raw_df.columns]
    if missing:
        raise ValueError(f"缺少必要列：{', '.join(missing)}")

    df = raw_df[required_columns].copy()

    # 清理字符串列里的空格 / 制表符
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].map(lambda x: x.strip() if isinstance(x, str) else x)

    # 删除收费冲正
    df = df[df["交易类型"] != "收费冲正"].copy()

    # 金额转为正数
    df["交易金额"] = pd.to_numeric(df["交易金额"], errors="coerce").fillna(0).abs()

    # 解析交易时间
    df["交易时间"] = pd.to_datetime(df["交易时间"], errors="coerce")
    invalid_time_count = int(df["交易时间"].isna().sum())

    # 删除无法解析时间的记录
    df = df[df["交易时间"].notna()].copy()
    if df.empty:
        raise ValueError("清洗后没有可用数据，请检查“交易时间”列格式。")

    # 生成餐费时间段
    df["餐费时间段"] = df["交易时间"].dt.time.map(get_meal_period)

    # 初始化新列
    for col in ["餐补金额", "自付（元）", "早餐（元）", "工作餐（元）", "加班餐（元）"]:
        df[col] = 0.0

    df["交易日期"] = df["交易时间"].dt.date
    df = df.sort_values(by=["姓名", "个人编号", "交易日期", "交易时间"]).copy()

    holiday_and_high_temp_days = get_holidays(holiday_year).union(high_temp_days)

    # 不再使用 groupby.apply，直接显式遍历分组，兼容性最好
    group_cols = ["姓名", "个人编号", "交易日期", "餐费时间段"]
    result_groups = []

    for _, group in df.groupby(group_cols, sort=False, observed=True):
        result_groups.append(
            calculate_subsidy_group(
                group=group,
                overtime_dates=overtime_dates,
                holiday_and_high_temp_days=holiday_and_high_temp_days
            )
        )

    if result_groups:
        df_result = pd.concat(result_groups, axis=0)
        df_result = df_result.sort_values(by=["姓名", "个人编号", "交易日期", "交易时间"])
    else:
        df_result = df.copy()

    df_final = df_result[
        ["人员类别", "姓名", "个人编号", "卡片类型", "交易地点", "卡户部门",
         "交易时间", "交易金额", "早餐（元）", "工作餐（元）", "加班餐（元）", "自付（元）"]
    ].copy()

    return df_final, invalid_time_count


# 输出为 Excel 字节流
def build_excel_bytes(df_final):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_final.to_excel(writer, index=False, sheet_name="餐补计算结果")
        ws = writer.book["餐补计算结果"]

        # 全表自动筛选
        ws.auto_filter.ref = ws.dimensions

        width_map = {
            "人员类别": 10,
            "姓名": 12,
            "个人编号": 14,
            "卡片类型": 12,
            "交易地点": 16,
            "卡户部门": 28,
            "交易时间": 20,
            "交易金额": 12,
            "早餐（元）": 12,
            "工作餐（元）": 12,
            "加班餐（元）": 12,
            "自付（元）": 12,
        }

        for i, header in enumerate(df_final.columns, start=1):
            col_letter = get_column_letter(i)
            ws.column_dimensions[col_letter].width = width_map.get(header, 13)

    output.seek(0)
    return output


# ---------------- Streamlit 页面 ----------------
st.title("餐补计算小程序")

uploaded_file = st.file_uploader("上传 CSV 文件", type=["csv"])

if uploaded_file is not None:
    try:
        raw_df, used_encoding = read_csv_with_fallback(uploaded_file)
    except Exception as e:
        st.error(f"CSV 读取失败：{e}")
        st.stop()

    # 尝试从数据里自动识别年份，作为默认节假日年份
    default_year = datetime.now().year
    if "交易时间" in raw_df.columns:
        trade_time_preview = pd.to_datetime(
            raw_df["交易时间"].astype(str).str.strip(),
            errors="coerce"
        )
        valid_years = trade_time_preview.dropna().dt.year
        if not valid_years.empty:
            default_year = int(valid_years.mode().iloc[0])

    st.caption(f"CSV 编码识别：{used_encoding} ｜ pandas 版本：{pd.__version__}")

    holiday_year = st.number_input(
        "请输入节假日年份",
        min_value=2000,
        max_value=2100,
        value=default_year,
        step=1
    )

    overtime_dates_input = st.text_input(
        "请输入加班调休日期（格式：YYYY-MM-DD,YYYY-MM-DD,...）",
        value=""
    )

    high_temp_days_input = st.text_input(
        "请输入高温假日期（格式：YYYY-MM-DD,YYYY-MM-DD,...）",
        value=""
    )

    overtime_dates, overtime_invalid = parse_date_set(overtime_dates_input)
    high_temp_days, high_temp_invalid = parse_date_set(high_temp_days_input)

    if overtime_invalid:
        st.warning(f"以下加班调休日期格式无效，已忽略：{', '.join(overtime_invalid)}")
    if high_temp_invalid:
        st.warning(f"以下高温假日期格式无效，已忽略：{', '.join(high_temp_invalid)}")

    try:
        df_final, invalid_time_count = process_dataframe(
            raw_df=raw_df,
            holiday_year=holiday_year,
            overtime_dates=overtime_dates,
            high_temp_days=high_temp_days
        )
    except Exception as e:
        st.error(f"处理失败：{e}")
        st.stop()

    if invalid_time_count > 0:
        st.warning(f"有 {invalid_time_count} 条记录的“交易时间”无法解析，已自动跳过。")

    excel_data = build_excel_bytes(df_final)

    st.success("✅ 数据处理完成！")
    st.dataframe(df_final, use_container_width=True)

    st.download_button(
        "📥 下载 Excel 文件",
        data=excel_data,
        file_name="餐补计算结果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )