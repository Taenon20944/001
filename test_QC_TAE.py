import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import gspread
from google.oauth2.service_account import Credentials

permanent_fixed_upper = {}
permanent_fixed_lower = {}
permanent_yellow_upper = {}
permanent_yellow_lower = {}

permanent_lock_upper = set()
permanent_lock_lower = set()

permanent_fixed_upper = st.session_state.get("permanent_fixed_upper", {})
permanent_yellow_upper = st.session_state.get("permanent_yellow_upper", {})

st.set_page_config(page_title="Brush Dashboard", layout="wide")

page = st.sidebar.radio("📂 เลือกหน้า", [
    "📊 หน้าแสดงผล rate และ ชั่วโมงที่เหลือ",
    "📝 กรอกข้อมูลแปลงถ่านเพิ่มเติม",
    "📈 พล็อตกราฟตามเวลา (แยก Upper และ Lower)"])

# https://docs.google.com/spreadsheets/d/1PUi4SXo4b_Zu7LO9mm4-EaYpPBnILSG41Jxr7a0Yaaw/edit?usp=sharing

# ------------------ PAGE 1 ------------------
if page == "📊 หน้าแสดงผล rate และ ชั่วโมงที่เหลือ":
    st.title("🛠️ วิเคราะห์อัตราสึกหรอและชั่วโมงที่เหลือของ Brush")

    # Setup credentials and spreadsheet access
    service_account_info = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(service_account_info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(creds)
    sheet_url = "https://docs.google.com/spreadsheets/d/1PUi4SXo4b_Zu7LO9mm4-EaYpPBnILSG41Jxr7a0Yaaw/edit?usp=sharing"
    sh = gc.open_by_url(sheet_url)

    sheet_names = [ws.title for ws in sh.worksheets()]
    if "Sheet1" in sheet_names:
        sheet_names.remove("Sheet1")
        sheet_names = ["Sheet1"] + sheet_names

    sheet_count = st.number_input("📌 เลือกจำนวน Sheet ที่ต้องใช้", min_value=1, max_value=len(sheet_names), value=7)
    selected_sheets = sheet_names[:sheet_count]
    


    import requests
    from io import BytesIO

    sheet_id = "1PUi4SXo4b_Zu7LO9mm4-EaYpPBnILSG41Jxr7a0Yaaw"
    sheet_url_export = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"

    response = requests.get(sheet_url_export)
    xls = pd.ExcelFile(BytesIO(response.content), engine="openpyxl")



    brush_numbers = list(range(1, 33))
    upper_rates, lower_rates = {n:{} for n in brush_numbers}, {n:{} for n in brush_numbers}
    rate_fixed_upper = set()
    rate_fixed_lower = set()
    yellow_mark_upper = {}
    yellow_mark_lower = {}

    # Step 1: Calculate rates per sheet
    for sheet in selected_sheets:
        df_raw = xls.parse(sheet, header=None)
        try:
            hours = float(df_raw.iloc[0, 7])
        except:
            continue
        df = xls.parse(sheet, skiprows=2, header=None)

        lower_df = df.iloc[:, 0:3]
        lower_df.columns = ["No_Lower", "Lower_Previous", "Lower_Current"]
        lower_df = lower_df.apply(pd.to_numeric, errors='coerce').dropna()

        upper_df = df.iloc[:, 4:6]
        upper_df.columns = ["Upper_Current", "Upper_Previous"]
        upper_df = upper_df.apply(pd.to_numeric, errors='coerce').dropna()
        upper_df["No_Upper"] = range(1, len(upper_df) + 1)

        for n in brush_numbers:
            u_row = upper_df[upper_df["No_Upper"] == n]
            if not u_row.empty:
                diff = u_row.iloc[0]["Upper_Current"] - u_row.iloc[0]["Upper_Previous"]
                rate = diff / hours if hours > 0 else 0
                upper_rates[n][f"Upper_{sheet}"] = rate if rate > 0 else 0

            l_row = lower_df[lower_df["No_Lower"] == n]
            if not l_row.empty:
                diff = l_row.iloc[0]["Lower_Previous"] - l_row.iloc[0]["Lower_Current"]
                rate = diff / hours if hours > 0 else 0
                lower_rates[n][f"Lower_{sheet}"] = rate if rate > 0 else 0


   # 🔧 ให้ผู้ใช้กรอกจำนวนรอบขั้นต่ำ และเปอร์เซ็นต์ threshold
 # ใช้ text_input แทน number_input เพื่อไม่ให้มี +/-
    min_required_str = st.text_input("🔢 จำนวนรอบขั้นต่ำที่ทำให้อัตราคงที่", value="5")
    threshold_percent_str = st.text_input("📉 เปอร์เซ็นต์ที่ยอมให้ (%)", value="5.0")

    # แปลง string เป็นตัวเลข (ระวัง error)
    try:
        min_required = int(min_required_str)
    except:
        min_required = 5  # fallback

    try:
        threshold_percent = float(threshold_percent_str)
    except:
        threshold_percent = 5.0

    threshold = threshold_percent / 100  # แปลงเป็นค่าเชิงทศนิยม

    def determine_final_rate(previous_rates, new_rate, row_index, sheet_name, mark_dict, min_required, threshold):
        previous_rates = [r for r in previous_rates if pd.notna(r) and r > 0]
        if len(previous_rates) >= min_required:
            avg_rate = sum(previous_rates) / len(previous_rates)
            percent_diff = abs(new_rate - avg_rate) / avg_rate
            if percent_diff <= threshold:
                mark_dict[row_index] = sheet_name
                return round(avg_rate, 6), True
        combined = previous_rates + [new_rate] if new_rate > 0 else previous_rates
        final_avg = sum(combined) / len(combined) if combined else 0
        return round(final_avg, 6), False


    # 3. แก้ calc_avg_with_flag ให้ใช้ permanent_* ให้ค่าคงที่ตลอด

    # เก็บค่าคงที่เดิมไว้ก่อน เพื่อไม่ให้รีเซ็ตทุกครั้ง
    if "permanent_fixed_upper" not in st.session_state:
        st.session_state.permanent_fixed_upper = {}
    if "permanent_yellow_upper" not in st.session_state:
        st.session_state.permanent_yellow_upper = {}
    if "permanent_fixed_lower" not in st.session_state:
        st.session_state.permanent_fixed_lower = {}
    if "permanent_yellow_lower" not in st.session_state:
        st.session_state.permanent_yellow_lower = {}

    
    sheet_index_map = {name: idx + 1 for idx, name in enumerate(selected_sheets)}
    
    # กำหนด จำนวนรอบที่ทำให้ rate คงที่ และ เปอร์เซ็นไม่เกินเท่าไร


    def calc_avg_with_flag(
        rates_dict, rate_fixed_set, mark_dict, permanent_fixed_rates,
        permanent_yellow_dict, sheet_index_map, min_required , threshold):
        df = pd.DataFrame.from_dict(rates_dict, orient='index')
        df = df.reindex(range(1, 33)).fillna(0)
        avg_col = []

        for i, row in df.iterrows():
            if i in permanent_fixed_rates:
                avg_col.append(permanent_fixed_rates[i])
                mark_dict[i] = permanent_yellow_dict.get(i, "")
                continue

            sheet_names = list(row[row > 0].index)
            values = row[row > 0].tolist()

            if len(values) >= min_required:
                for j in range(min_required, len(values) + 1):
                    prev = values[:j - 1]
                    new = values[j - 1]
                    sheet_name = sheet_names[j - 1]
                    avg = sum(prev) / len(prev) if prev else 0
                    percent_diff = abs(new - avg) / avg if avg > 0 else 1

                    if percent_diff <= threshold:
                        final_avg = round(avg, 6)
                        avg_col.append(final_avg)
                        rate_fixed_set.add(i)
                        permanent_fixed_rates[i] = final_avg
                        permanent_yellow_dict[i] = sheet_name
                        break
                else:
                    avg_col.append(round(sum(values) / len(values), 6))

            else:
                avg_col.append(round(sum(values) / len(values), 6) if values else 0.000000)

        return df, avg_col



    
 

    # 4. เรียกใช้แบบใหม่ (ตัวอย่าง Upper)

    upper_df, upper_avg = calc_avg_with_flag(
    upper_rates, rate_fixed_upper, yellow_mark_upper,
    permanent_fixed_upper, permanent_yellow_upper,sheet_index_map,min_required, threshold)

    lower_df, lower_avg = calc_avg_with_flag(
        lower_rates, rate_fixed_lower, yellow_mark_lower,
        permanent_fixed_lower, permanent_yellow_lower,sheet_index_map,min_required, threshold)
    

    st.session_state.permanent_fixed_upper = permanent_fixed_upper
    st.session_state.permanent_yellow_upper = permanent_yellow_upper
    st.session_state.permanent_fixed_lower = permanent_fixed_lower
    st.session_state.permanent_yellow_lower = permanent_yellow_lower

    
 


    upper_df["Avg Rate (Upper)"] = upper_avg
    lower_df["Avg Rate (Lower)"] = lower_avg

    # Step 3: Styling output
    def highlight_fixed_rate_row(row, column_name, permanent_fixed_rates, yellow_mark_dict):
        styles = []
        for col in row.index:
            if col == column_name:
                if row.name in permanent_fixed_rates:
                    correct_fixed_value = permanent_fixed_rates[row.name]
                    if abs(row[col] - correct_fixed_value) < 1e-6:
                        styles.append("background-color: green; color: black; font-weight: bold")
                    else:
                        styles.append("color: red; font-weight: bold")
                else:
                    styles.append("color: red; font-weight: bold")
            elif yellow_mark_dict.get(row.name, "") == col:
                styles.append("color: yellow; font-weight: bold")
            else:
                styles.append("")
        return styles
    
    round_show = min_required
    percent_show = threshold * 100
    
    
    #
    st.markdown(f"จำนวนรอบขั้นต่ำที่ทำให้อัตราการลดลงคงที่คงที่เท่ากับ {round_show} รอบ")
    st.markdown(f"จำนวนเปอร์เซ็นสูงที่สุดที่ทำให้คิดเป็นอัตราการลดลงคงที่ ไม่เกิน {percent_show} %")




    st.subheader("📋 ตาราง Avg Rate - Upper")
    styled_upper = upper_df.style.apply(
    lambda row: highlight_fixed_rate_row(row, "Avg Rate (Upper)", permanent_fixed_upper, permanent_yellow_upper),
    axis=1).format("{:.6f}")
    st.write(styled_upper)



    st.subheader("📋 ตาราง Avg Rate - Lower")
    styled_lower = lower_df.style.apply(
    lambda row: highlight_fixed_rate_row(row, "Avg Rate (Lower)", permanent_fixed_lower, permanent_yellow_lower),
    axis=1).format("{:.6f}")
    st.write(styled_lower)

    st.markdown("🟩 **สีเขียว** = ค่าคงที่ที่นำไปใช้ในกราฟ")
    st.markdown("🟨 **ตัวอักษรสีเหลือง** = ค่า Rate ที่ทำให้ค่าเฉลี่ยกลายเป็น 'คงที่'")
    st.markdown("🔴 **สีแดง** = ค่า Rate ยังไม่คงที่")



    avg_rate_upper = upper_avg
    avg_rate_lower = lower_avg
    

    if "Sheet7" in xls.sheet_names:
            df_sheet7 = xls.parse("Sheet7", header=None)
            upper_current = pd.to_numeric(df_sheet7.iloc[2:34, 5], errors='coerce').values
            lower_current = pd.to_numeric(df_sheet7.iloc[2:34, 2], errors='coerce').values

    def calculate_hours_safe(current, rate):
            return [(c - 35) / r if pd.notna(c) and r and r > 0 and c > 35 else 0 for c, r in zip(current, rate)]

    hour_upper = calculate_hours_safe(upper_current, avg_rate_upper)
    hour_lower = calculate_hours_safe(lower_current, avg_rate_lower)



    st.subheader("📊 กราฟรวม Avg Rate")
    fig_combined = go.Figure()
    fig_combined.add_trace(go.Scatter(x=brush_numbers, y=avg_rate_upper, mode='lines+markers+text', name='Upper Avg Rate', line=dict(color='red'), text=[str(i) for i in brush_numbers], textposition='top center'))
    fig_combined.add_trace(go.Scatter(x=brush_numbers, y=avg_rate_lower, mode='lines+markers+text', name='Lower Avg Rate', line=dict(color='deepskyblue'), text=[str(i) for i in brush_numbers], textposition='top center'))
    fig_combined.update_layout(xaxis_title='Brush Number', yaxis_title='Wear Rate (mm/hour)', template='plotly_white')
    st.plotly_chart(fig_combined, use_container_width=True)



    st.subheader("🔺 กราฟ Avg Rate - Upper")
    fig_upper = go.Figure()
    fig_upper.add_trace(go.Scatter(x=brush_numbers, y=avg_rate_upper, mode='lines+markers+text', name='Upper Avg Rate', line=dict(color='red'), text=[str(i) for i in brush_numbers], textposition='top center'))
    fig_upper.update_layout(xaxis_title='Brush Number', yaxis_title='Wear Rate (mm/hour)', template='plotly_white')
    st.plotly_chart(fig_upper, use_container_width=True)

    st.subheader("🔻 กราฟ Avg Rate - Lower")
    fig_lower = go.Figure()
    fig_lower.add_trace(go.Scatter(x=brush_numbers, y=avg_rate_lower, mode='lines+markers+text', name='Lower Avg Rate', line=dict(color='deepskyblue'), text=[str(i) for i in brush_numbers], textposition='top center'))
    fig_lower.update_layout(xaxis_title='Brush Number', yaxis_title='Wear Rate (mm/hour)', template='plotly_white')
    st.plotly_chart(fig_lower, use_container_width=True)


    #sheet_names = [ws.title for ws in sh.worksheets() if ws.title.lower().startswith("sheet")]
    #sheet_count = st.number_input("📌 กรอกจำนวนชีตย้อนหลังที่ต้องใช้", min_value=1, max_value=len(sheet_names), value=6)
    try:
        
        xls = pd.ExcelFile(sheet_url_export, engine='openpyxl')
        
        selected_sheet_names = sheet_names[:sheet_count]
        brush_numbers = list(range(1, 33))
        upper_rates, lower_rates = {n: {} for n in brush_numbers}, {n: {} for n in brush_numbers}

        for sheet in selected_sheet_names:
            df_raw = xls.parse(sheet, header=None)
            try:
                hours = float(df_raw.iloc[0, 7])
            except:
                continue
            df = xls.parse(sheet, skiprows=2, header=None)

            lower_df = df.iloc[:, 0:3]
            lower_df.columns = ["No_Lower", "Lower_Previous", "Lower_Current"]
            lower_df = lower_df.dropna().apply(pd.to_numeric, errors='coerce')

            upper_df = df.iloc[:, 4:6]
            upper_df.columns = ["Upper_Current", "Upper_Previous"]
            upper_df = upper_df.dropna().apply(pd.to_numeric, errors='coerce')
            upper_df["No_Upper"] = range(1, len(upper_df) + 1)

            for n in brush_numbers:
                u_row = upper_df[upper_df["No_Upper"] == n]
                if not u_row.empty:
                    diff = u_row.iloc[0]["Upper_Current"] - u_row.iloc[0]["Upper_Previous"]
                    rate = diff / hours if hours > 0 else np.nan
                    upper_rates[n][f"Upper_{sheet}"] = rate if rate > 0 else np.nan

                l_row = lower_df[lower_df["No_Lower"] == n]
                if not l_row.empty:
                    diff = l_row.iloc[0]["Lower_Previous"] - l_row.iloc[0]["Lower_Current"]
                    rate = diff / hours if hours > 0 else np.nan
                    lower_rates[n][f"Lower_{sheet}"] = rate if rate > 0 else np.nan

        def avg_positive(row):
            valid = row[row > 0]
            return valid.sum() / len(valid) if len(valid) > 0 else np.nan

            # ใช้ค่าที่คำนวณแล้วก่อนหน้า
        avg_rate_upper = upper_avg
        avg_rate_lower = lower_avg

        df_current = xls.parse(f"Sheet{sheet_count}", header=None, skiprows=2)
        upper_current = pd.to_numeric(df_current.iloc[0:32, 5], errors='coerce').values
        lower_current = pd.to_numeric(df_current.iloc[0:32, 2], errors='coerce').values

        def calculate_hours_safe(current, rate):
            return [(c - 35) / r if pd.notna(c) and r and r > 0 and c > 35 else 0 for c, r in zip(current, rate)]

        hour_upper = calculate_hours_safe(upper_current, avg_rate_upper)
        hour_lower = calculate_hours_safe(lower_current, avg_rate_lower)

        st.subheader("📋 ตารางผลการคำนวณ")
        result_df = pd.DataFrame({
            "Brush #": brush_numbers,
            "Upper Current (F)": upper_current,
            "Lower Current (C)": lower_current,
            "Avg Rate Upper": avg_rate_upper,
            "Avg Rate Lower": avg_rate_lower,
            "Remaining Hours Upper": hour_upper,
            "Remaining Hours Lower": hour_lower,
        })
        st.dataframe(result_df, use_container_width=True)

        st.subheader("📊 กราฟ Remaining Hours ถึง 35mm")
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 8))

        color_upper = ['black' if h < 500 else 'red' for h in hour_upper]
        bars1 = ax1.bar(brush_numbers, hour_upper, color=color_upper)
        ax1.set_title("Remaining Hours to Reach 35mm - Upper")
        ax1.set_ylabel("Hours")
        ax1.set_xticks(brush_numbers)
        for bar, val in zip(bars1, hour_upper):
            ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 10, f"{int(val)}", ha='center', fontsize=8)

        color_lower = ['black' if h < 500 else 'deepskyblue' for h in hour_lower]
        bars2 = ax2.bar(brush_numbers, hour_lower, color=color_lower)
        ax2.set_title("Remaining Hours to Reach 35mm - Lower")
        ax2.set_ylabel("Hours")
        ax2.set_xticks(brush_numbers)
        for bar, val in zip(bars2, hour_lower):
            ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 10, f"{int(val)}", ha='center', fontsize=8)

        plt.tight_layout()
        st.pyplot(fig)

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")
        
    st.session_state.upper_avg = upper_avg
    st.session_state.lower_avg = lower_avg

# --------------------------------------------------- PAGE 2 -------------------------------------------------


elif page == "📝 กรอกข้อมูลแปลงถ่านเพิ่มเติม":
    st.title("📝 กรอกข้อมูลแปรงถ่าน + ชั่วโมง")
    
    from io import BytesIO
    import requests

    sheet_id = "1PUi4SXo4b_Zu7LO9mm4-EaYpPBnILSG41Jxr7a0Yaaw"
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    response = requests.get(url)

    xls = pd.ExcelFile(BytesIO(response.content), engine="openpyxl")



    service_account_info = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(service_account_info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(creds)
    sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1PUi4SXo4b_Zu7LO9mm4-EaYpPBnILSG41Jxr7a0Yaaw/edit?usp=sharing")

# ✅ ดึงเฉพาะชีตที่ชื่อขึ้นต้นด้วย Sheet (หรือเปลี่ยนเป็นตาม pattern ของคุณ เช่น "Sheet1", "Sheet2", ...)
    # ✅ 1. เตรียมรายชื่อชีตทั้งหมดแบบ normalize (รองรับ sheet ชื่อเล็ก/ใหญ่)


    sheet_names_all = [ws.title for ws in sh.worksheets()]

    def extract_sheet_number(name):
        try:
            return int(name.lower().replace("sheet", ""))
        except:
            return float("inf")

    sheet_names = [s for s in sheet_names_all if s.lower().startswith("sheet")]
    sheet_names_sorted = sorted(sheet_names, key=extract_sheet_number)
    if "Sheet1" in sheet_names_sorted:
        sheet_names_sorted.remove("Sheet1")
        sheet_names_sorted = ["Sheet1"] + sheet_names_sorted

    sheet_names = sheet_names_sorted

    filtered_sheet_names = [s for s in sheet_names_all if s.lower().startswith("sheet") and s.lower() != "sheet1"]

    # ✅ 2. ดึงตัวเลขของ SheetN
    sheet_numbers = []
    for name in filtered_sheet_names:
        suffix = name.lower().replace("sheet", "")
        if suffix.isdigit():
            sheet_numbers.append(int(suffix))

    sheet_numbers.sort()
    next_sheet_number = sheet_numbers[-1] + 1 if sheet_numbers else 2
    next_sheet_name = f"Sheet{next_sheet_number}"

    # ทำให้ sheet มีการเรียงกัน
    def extract_sheet_number(name):
        try:
            return int(name.lower().replace("sheet", ""))
        except:
            return float('inf')  # สำหรับกรณีชื่อไม่ใช่ตัวเลข

    sheet_names = [s for s in sheet_names_all if s.lower().startswith("sheet")]
    sheet_names_sorted = sorted(sheet_names, key=extract_sheet_number)

    # ถ้าอยากให้ Sheet1 อยู่บนสุดเสมอ:
    if "Sheet1" in sheet_names_sorted:
        sheet_names_sorted.remove("Sheet1")
        sheet_names_sorted = ["Sheet1"] + sheet_names_sorted

    sheet_names = sheet_names_sorted
    
    
    filtered_sheet_names = [s for s in sheet_names if s.lower() != "sheet1"]
    sheet_numbers = [
        int(s.lower().replace("sheet", "")) 
        for s in filtered_sheet_names if s.lower().replace("sheet", "").isdigit()
    ]
    sheet_numbers.sort()

    next_sheet_number = sheet_numbers[-1] + 1 if sheet_numbers else 2
    next_sheet_name = f"Sheet{next_sheet_number}"
    
    selected_sheet_auto = st.session_state.get("selected_sheet_auto", "Sheet1")
    if selected_sheet_auto not in sheet_names:
        selected_sheet_auto = sheet_names[0]  # fallback เผื่อ sheet ใหม่ยังไม่เจอทัน

    selected_sheet = st.selectbox("📄 เลือก Sheet ที่ต้องการกรอกข้อมูล", sheet_names_sorted)

    #st.write(f"🧪 Selected (auto): {selected_sheet_auto}")
    #st.write(f"🧪 Dropdown Options: {sheet_names}")
   

        # ✅ เตรียมชื่อชีตถัดไป (เช่น Sheet13)
    
    
    next_sheet_number = sheet_numbers[-1] + 1 if sheet_numbers else 2
    next_sheet_name = f"Sheet{next_sheet_number}"

    


        # ดึงเลขชีตล่าสุดก่อนแสดงปุ่ม
    filtered_sheet_names = [s for s in sheet_names if s.lower().startswith("sheet") and s.lower() != "sheet1"]
    sheet_numbers = [int(s.lower().replace("sheet", "")) for s in filtered_sheet_names if s.lower().replace("sheet", "").isdigit()]
    sheet_numbers.sort()
    next_sheet_name = f"Sheet{sheet_numbers[-1] + 1}" if sheet_numbers else "Sheet2"

    # 📌 คำนวณชื่อชีตใหม่ (SheetN+1)
    filtered_sheet_names = [s for s in sheet_names if s.lower() != "sheet1"]
    sheet_numbers = [int(s.lower().replace("sheet", "")) for s in filtered_sheet_names if s.lower().replace("sheet", "").isdigit()]
    sheet_numbers.sort()
    next_sheet_number = sheet_numbers[-1] + 1 if sheet_numbers else 2
    next_sheet_name = f"Sheet{next_sheet_number}"

    # 📦 ปุ่มสร้างชีตใหม่
    if st.button(f"➕ สร้างชีตที่ {next_sheet_name} "):
        try:
            # ใช้ sheet ล่าสุดเป็นต้นแบบ
            last_sheet = f"Sheet{sheet_numbers[-1]}"
            source_ws = sh.worksheet(last_sheet)
            df_prev = source_ws.get_all_values()

            # คัดลอกค่า current
            lower_previous_formulas = [[f"={last_sheet}!C{i+3}"] for i in range(32)]
            upper_previous_formulas = [[f"={last_sheet}!F{i+3}"] for i in range(32)]
            

            # ตรวจว่าชีตนี้มีอยู่แล้วหรือไม่
            if next_sheet_name.lower() in [ws.title.lower() for ws in sh.worksheets()]:
                st.warning(f"⚠️ Sheet '{next_sheet_name}' มีอยู่แล้ว")
                st.stop()

            # สร้างชีตใหม่
            new_ws = sh.duplicate_sheet(source_sheet_id=source_ws.id, new_sheet_name=next_sheet_name)
            
            sheets = sh.worksheets()
            new_ws = sh.worksheet(next_sheet_name)
            # ย้าย sheet ไปท้ายสุด
            sheets = [ws for ws in sheets if ws.title != next_sheet_name]
            sheets.append(new_ws)
            sh.reorder_worksheets(sheets)

            
                       
                        
            # วางสูตร (ระบุ USER_ENTERED เพื่อให้เป็นสูตร)
            new_ws.update("B3:B34", lower_previous_formulas, value_input_option="USER_ENTERED")
            new_ws.update("E3:E34", upper_previous_formulas, value_input_option="USER_ENTERED")
            
            
            try:
                new_ws.update("B3:B34", lower_previous_formulas, value_input_option="USER_ENTERED")
                new_ws.update("E3:E34", upper_previous_formulas, value_input_option="USER_ENTERED")
            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาดขณะใส่สูตร: {e}")


            from gspread.utils import rowcol_to_a1
            
            import time

            for i in range(32):

                if i % 10 == 0:
                    time.sleep(2)



            st.session_state["selected_sheet_auto"] = next_sheet_name  # ✅ เพิ่มบรรทัดนี้
            st.success(f"✅ สร้างชีต '{next_sheet_name}' สำเร็จแล้ว 🎉")
            st.rerun()
        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {e}")






    # โหลดค่าทันทีจาก selected_sheet
    ws = sh.worksheet(selected_sheet)
    df_prev = ws.get_all_values()

    lower_current = [row[2] if len(row) > 2 else "" for row in df_prev[2:34]]
    upper_current = [row[5] if len(row) > 5 else "" for row in df_prev[2:34]]

    # โหลดชั่วโมง/วัน
    try:
        default_hours = float(ws.acell("H1").value or 0)
    except:
        default_hours = 0.0
    default_prev_date = ws.acell("A2").value or ""
    default_curr_date = ws.acell("B2").value or ""


    hours = st.number_input("⏱️ ชั่วโมง", min_value=0.0, step=0.1, value=float(default_hours))
    
    prev_date = st.text_input("📅 วันที่ตรวจก่อนหน้า", placeholder="DD/MM/YYYY", value=default_prev_date)
    curr_date = st.text_input("📅 วันที่ตรวจล่าสุด", placeholder="DD/MM/YYYY", value=default_curr_date)

 
    
    

    

    st.markdown("### 🔧 แปลงถ่านส่วน LOWER")
    lower = []
    cols = st.columns(8)
    for i in range(32):
        col = cols[i % 8]
        with col:
            st.markdown(f"<div style='text-align: center;'>แปลงถ่านที่ {i+1}</div>", unsafe_allow_html=True)
            value = st.text_input(
                label="",  # 👈 ใส่ label เป็นค่าว่าง
                key=f"lower_input_{i}",
                value=str(lower_current[i]),
                label_visibility="collapsed",  # 👈 ซ่อน label แบบสมบูรณ์
                )

            #value = st.text_input(
            #f"lower_{i+1}",                     # ชื่อ label
            #key=f"lower_input_{i}",             # key ไม่ซ้ำ
            #value=str(lower_current[i]),       # ดึงค่าปัจจุบันมาแสดง
        #)
            try:
                lower.append(float(value))
            except:
                lower.append(0.0)
                
    st.markdown("### 🔧 แปลงถ่านส่วน UPPER")
    upper = []
    cols = st.columns(8)
    for i in range(32):
        col = cols[i % 8]
        with col:
            st.markdown(f"<div style='text-align: center;'>แปลงถ่านที่ {i+1}</div>", unsafe_allow_html=True)
            value = st.text_input(
                label="",  # 👈 ใส่ label เป็นค่าว่าง
                key=f"upper_input_{i}",
                value=str(upper_current[i]),
                label_visibility="collapsed",  # 👈 ซ่อน label แบบสมบูรณ์
                )

            try:
                upper.append(float(value))
            except:
                upper.append(0.0)

    if st.button("📤 บันทึก"):
        try:
            ws.update("A2", [[prev_date]])
            ws.update("B2", [[curr_date]])
            ws.update("H1", [[hours]])

            st.success(f"✅ บันทึกลง {selected_sheet} แล้วเรียบร้อย")
        except Exception as e:
            st.error(f"❌ {e}")

    # ------------------ แสดงตารางรวม ------------------
    st.subheader("📄 ตารางรวม Upper + Lower (Current / Previous)")
    
    # เชื่อมต่อ Google Sheet
    sheet_id = "1PUi4SXo4b_Zu7LO9mm4-EaYpPBnILSG41Jxr7a0Yaaw"
    sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    xls = pd.ExcelFile(sheet_url)
   
    # 📌 เลือกชีตที่ต้องการดู
    sheet_options = [ws.title for ws in sh.worksheets() if ws.title.lower().startswith("sheet")]
    selected_view_sheet = st.selectbox("📌 เลือกชีตที่ต้องการดู", sheet_options)

    try:
        #คำหนดคำสั่ง
        selected_ws = sh.worksheet(selected_view_sheet)
        
        #ดึงค่ามาจาก google sheet
        date_prev = selected_ws.acell("A2").value
        date_curr = selected_ws.acell("B2").value        
        hour_val = selected_ws.acell("H1").value
        
        #เอาไปกรอกใน web
        st.markdown(f"📆 วันที่ Previous: **{date_prev}** | วันที่ Current: **{date_curr}**")
        st.markdown(f"#### ⏱️ ชั่วโมงจาก {selected_view_sheet}: {hour_val} ชั่วโมง")

        df = xls.parse(selected_view_sheet, skiprows=1, header=None)
        
        upper_df = df.iloc[:, 4:6]
        upper_df.columns = ["Upper_Previous", "Upper_Current"]
        lower_df = df.iloc[:, 1:3]
        lower_df.columns = ["Lower_Previous", "Lower_Current"]
        
        #ลองสลับค่า
        
        # กรองเฉพาะค่าตัวเลข (drop non-numeric row)
        lower_df = lower_df[pd.to_numeric(lower_df["Lower_Current"], errors="coerce").notna()]
        upper_df = upper_df[pd.to_numeric(upper_df["Upper_Current"], errors="coerce").notna()]

        #ลองแก้หน่อย
        #combined_df = pd.concat([upper_df.reset_index(drop=True), lower_df.reset_index(drop=True)], axis=1)
        #st.dataframe(combined_df, use_container_width=True)
        
        combined_df = pd.concat([lower_df.reset_index(drop=True), upper_df.reset_index(drop=True)], axis=1)
        combined_df.insert(0, "Brush No", range(1, len(combined_df) + 1))
        combined_df.set_index("Brush No", inplace=True)
        st.dataframe(combined_df, use_container_width=True, height=700)



        st.markdown("### 📊 กราฟรวม Upper และ Lower (Current vs Previous)")
        brush_labels = [f"Brush {i+1}" for i in range(len(combined_df))]

        fig = go.Figure()
        fig.add_trace(go.Scatter(
            y=combined_df["Upper_Current"], x=brush_labels,
            mode='lines+markers', name='Upper Current'))
        
        fig.add_trace(go.Scatter(
            y=combined_df["Upper_Previous"], x=brush_labels,
            mode='lines+markers', name='Upper Previous'))
        
        fig.add_trace(go.Scatter(
            y=combined_df["Lower_Current"], x=brush_labels,
            mode='lines+markers', name='Lower Current', line=dict(dash='dot')))
        
        fig.add_trace(go.Scatter(
            y=combined_df["Lower_Previous"], x=brush_labels,
            mode='lines+markers', name='Lower Previous', line=dict(dash='dot')))
        
        fig.update_layout(
            xaxis_title='Brush Number',
            yaxis_title='mm',
            height=600,
            width=1400,  # ✅ เพิ่มความกว้างให้กราฟแสดงเต็มแนวนอน
            xaxis=dict(
                tickmode='linear',
                tick0=1,
                dtick=1,
                type='category'),  # ✅ ให้ Plotly จัด category label brush ให้ดีขึ้น
            
            margin=dict(l=40, r=40, t=40, b=40))

        st.plotly_chart(fig, use_container_width=True)

    except Exception as e:
        st.error(f"❌ ไม่สามารถโหลดข้อมูลจากชีตนี้ได้: {e}")
        
        
        
        
        
        
        
        
        
        
        
        
# ------------------ PAGE 3 ------------------





elif page == "📈 พล็อตกราฟตามเวลา (แยก Upper และ Lower)":
    st.title("📈 พล็อตกราฟตามเวลา (แยก Upper และ Lower)")

    # เชื่อมต่อ Google Sheet
    sheet_id = "1PUi4SXo4b_Zu7LO9mm4-EaYpPBnILSG41Jxr7a0Yaaw"
    sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    xls = pd.ExcelFile(sheet_url)
    
    

    sheet_count = st.number_input("📌 กรอกจำนวนชีตย้อนหลังที่ต้องใช้ (1-7)", min_value=1, max_value=7, value=6)
    # ดึงชื่อชีตจริงจากไฟล์
    all_sheet_names = xls.sheet_names
    sheet_names = [s for s in all_sheet_names if s.lower().startswith("sheet")][:sheet_count]

    brush_numbers = list(range(1, 33))
    upper_rates, lower_rates = {n: {} for n in brush_numbers}, {n: {} for n in brush_numbers}

    for sheet in sheet_names:
        df_raw = xls.parse(sheet, header=None)
        try:
            hours = float(df_raw.iloc[0, 7])
        except:
            continue
        df = xls.parse(sheet, skiprows=2, header=None)

        lower_df = df.iloc[:, 0:3]
        lower_df.columns = ["No_Lower", "Lower_Previous", "Lower_Current"]
        lower_df = lower_df.dropna().apply(pd.to_numeric, errors='coerce')

        upper_df = df.iloc[:, 4:6]
        upper_df.columns = ["Upper_Current", "Upper_Previous"]
        upper_df = upper_df.dropna().apply(pd.to_numeric, errors='coerce')
        upper_df["No_Upper"] = range(1, len(upper_df) + 1)

        for n in brush_numbers:
            u_row = upper_df[upper_df["No_Upper"] == n]
            if not u_row.empty:
                diff = u_row.iloc[0]["Upper_Current"] - u_row.iloc[0]["Upper_Previous"]
                rate = diff / hours if hours > 0 else np.nan
                upper_rates[n][f"Upper_{sheet}"] = rate if rate > 0 else np.nan

            l_row = lower_df[lower_df["No_Lower"] == n]
            if not l_row.empty:
                diff = l_row.iloc[0]["Lower_Previous"] - l_row.iloc[0]["Lower_Current"]
                rate = diff / hours if hours > 0 else np.nan
                lower_rates[n][f"Lower_{sheet}"] = rate if rate > 0 else np.nan

    def avg_positive(row_dict):
        values = [v for v in row_dict.values() if pd.notna(v) and v > 0]
        return sum(values) / len(values) if values else np.nan
    
    def determine_final_rate(previous_rates, new_rate, row_index, sheet_name, mark_dict, min_required=5, threshold=0.1):
        previous_rates = [r for r in previous_rates if pd.notna(r) and r > 0]
        if len(previous_rates) >= min_required:
            avg_rate = sum(previous_rates) / len(previous_rates)
            percent_diff = abs(new_rate - avg_rate) / avg_rate
            if percent_diff <= threshold:
                mark_dict[row_index] = sheet_name
                return round(avg_rate, 6), True
        combined = previous_rates + [new_rate] if new_rate > 0 else previous_rates
        final_avg = sum(combined) / len(combined) if combined else 0
        return round(final_avg, 6), False

    def calc_avg_with_flag(rates_dict, rate_fixed_set, mark_dict):
        df = pd.DataFrame.from_dict(rates_dict, orient='index')
        df = df.reindex(range(1, 33)).fillna(0)
        avg_col = []
        for i, row in df.iterrows():
            values = row[row > 0].tolist()
            if len(values) >= 6:
                prev = values[:-1]
                new = values[-1]
                sheet_name = row[row > 0].index[-1] if len(row[row > 0].index) > 0 else ""
                avg, fixed = determine_final_rate(prev, new, i, sheet_name, mark_dict)
                avg_col.append(avg)
                if fixed:
                    rate_fixed_set.add(i)
            else:
                avg_col.append(round(np.mean(values), 6) if values else 0.000000)
        return df, avg_col
    

    # ใช้ calc_avg_with_flag ที่คุณมีอยู่แล้ว
    rate_fixed_upper = set()
    rate_fixed_lower = set()
    yellow_mark_upper = {}
    yellow_mark_lower = {}

    upper_df, avg_rate_upper = calc_avg_with_flag(upper_rates, rate_fixed_upper, yellow_mark_upper)
    lower_df, avg_rate_lower = calc_avg_with_flag(lower_rates, rate_fixed_lower, yellow_mark_lower)



 

    # ใช้ current จาก sheet ล่าสุด เช่น Sheet{sheet_count}
    df_current = xls.parse(f"Sheet{sheet_count}", header=None, skiprows=2)
    upper_current = pd.to_numeric(df_current.iloc[0:32, 5], errors='coerce').values
    lower_current = pd.to_numeric(df_current.iloc[0:32, 2], errors='coerce').values

    time_hours = np.arange(0, 201, 10)

    # UPPER
    fig_upper = go.Figure()
    for i, (start, rate) in enumerate(zip(upper_current, avg_rate_upper)):
        if pd.notna(start) and pd.notna(rate) and rate > 0:
            y = [start - rate*t for t in time_hours]
            fig_upper.add_trace(go.Scatter(x=time_hours, y=y, name=f"Upper {i+1}", mode='lines'))

    fig_upper.add_shape(type="line", x0=0, x1=200, y0=35, y1=35, line=dict(color="firebrick", width=2, dash="dash"))
    fig_upper.add_annotation(x=5, y=35, text="⚠️ 35 mm", showarrow=False, font=dict(color="firebrick", size=12), bgcolor="white")

    fig_upper.update_layout(title="🔺 ความยาว Upper ตามเวลา", xaxis_title="ชั่วโมง", yaxis_title="mm",
                            xaxis=dict(dtick=10, range=[0, 200]), yaxis=dict(range=[30, 65]))
    st.plotly_chart(fig_upper, use_container_width=True)

    # LOWER
    fig_lower = go.Figure()
    for i, (start, rate) in enumerate(zip(lower_current, avg_rate_lower)):
        if pd.notna(start) and pd.notna(rate) and rate > 0:
            y = [start - rate*t for t in time_hours]
            fig_lower.add_trace(go.Scatter(x=time_hours, y=y, name=f"Lower {i+1}", mode='lines', line=dict(dash='dot')))

    fig_lower.add_shape(type="line", x0=0, x1=200, y0=35, y1=35, line=dict(color="firebrick", width=2, dash="dash"))
    fig_lower.add_annotation(x=5, y=35, text="⚠️  35 mm", showarrow=False, font=dict(color="firebrick", size=12), bgcolor="white")

    fig_lower.update_layout(title="🔻 ความยาว Lower ตามเวลา", xaxis_title="ชั่วโมง", yaxis_title="mm",
                            xaxis=dict(dtick=10, range=[0, 200]), yaxis=dict(range=[30, 65]))
    st.plotly_chart(fig_lower, use_container_width=True)
