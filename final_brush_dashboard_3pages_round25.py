import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import matplotlib.pyplot as plt
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Brush Dashboard", layout="wide")

page = st.sidebar.radio("📂 เลือกหน้า", [
    "📊 หน้าแสดงผล rate และ ชั่วโมงที่เหลือ",
    "📝 กรอกข้อมูลแปลงถ่านเพิ่มเติม",
    "📈 พล็อตกราฟตามเวลา (แยก Upper และ Lower)"
])


# ------------------ PAGE 1 ------------------
if page == "📊 หน้าแสดงผล rate และ ชั่วโมงที่เหลือ":
    st.title("🛠️ วิเคราะห์อัตราสึกหรอและชั่วโมงที่เหลือของ Brush")

    sheet_id = "1SOkIH9jchaJi_0eck5UeyUR8sTn2arndQofmXv5pTdQ"
    sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    xls = pd.ExcelFile(sheet_url)
    
    service_account_info = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(service_account_info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(creds)
    sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1SOkIH9jchaJi_0eck5UeyUR8sTn2arndQofmXv5pTdQ")

    
    sheet_names = [ws.title for ws in sh.worksheets() if ws.title.lower().startswith("sheet")]
    num_sheets = st.number_input("📌 เลือกจำนวน Sheet ที่ต้องการใช้ (สำหรับคำนวณ Avg Rate)", min_value=1, max_value=len(sheet_names), value=7)
    selected_sheets = sheet_names[:num_sheets]

    brush_numbers = list(range(1, 33))

    upper_rates, lower_rates = {n:{} for n in brush_numbers}, {n:{} for n in brush_numbers}

    for sheet in selected_sheets:
        df_raw = xls.parse(sheet, header=None)
        try:
            hours = float(df_raw.iloc[0, 7])
        except:
            continue
        df = xls.parse(sheet, skiprows=1, header=None)

        lower_df_sheet = df.iloc[:, 0:3]
        lower_df_sheet.columns = ["No_Lower", "Lower_Previous", "Lower_Current"]
        lower_df_sheet = lower_df_sheet.dropna().apply(pd.to_numeric, errors='coerce')

        upper_df_sheet = df.iloc[:, 4:6]
        upper_df_sheet.columns = ["Upper_Current", "Upper_Previous"]
        upper_df_sheet = upper_df_sheet.dropna().apply(pd.to_numeric, errors='coerce')
        upper_df_sheet["No_Upper"] = range(1, len(upper_df_sheet) + 1)

        for n in brush_numbers:
            u_row = upper_df_sheet[upper_df_sheet["No_Upper"] == n]
            if not u_row.empty:
                diff = u_row.iloc[0]["Upper_Current"] - u_row.iloc[0]["Upper_Previous"]
                rate = diff / hours if hours > 0 else np.nan
                upper_rates[n][f"Upper_{sheet}"] = rate if rate > 0 else np.nan

            l_row = lower_df_sheet[lower_df_sheet["No_Lower"] == n]
            if not l_row.empty:
                diff = l_row.iloc[0]["Lower_Previous"] - l_row.iloc[0]["Lower_Current"]
                rate = diff / hours if hours > 0 else np.nan
                lower_rates[n][f"Lower_{sheet}"] = rate if rate > 0 else np.nan

    def avg_positive(row):
        valid = row[row > 0]
        return valid.sum() / len(valid) if len(valid) > 0 else np.nan

    upper_df = pd.DataFrame.from_dict(upper_rates, orient='index').fillna(0)
    lower_df = pd.DataFrame.from_dict(lower_rates, orient='index').fillna(0)
    upper_df["Avg Rate (Upper)"] = upper_df.apply(avg_positive, axis=1)
    lower_df["Avg Rate (Lower)"] = lower_df.apply(avg_positive, axis=1)

    avg_rate_upper = upper_df["Avg Rate (Upper)"].tolist()[:32]
    avg_rate_lower = lower_df["Avg Rate (Lower)"].tolist()[:32]

    if "Sheet7" in xls.sheet_names:
        df_sheet7 = xls.parse("Sheet7", header=None)
        upper_current = pd.to_numeric(df_sheet7.iloc[2:34, 5], errors='coerce').values
        lower_current = pd.to_numeric(df_sheet7.iloc[2:34, 2], errors='coerce').values
    else:
        st.error("❌ ไม่พบ Sheet7 สำหรับค่าสภาพปัจจุบัน")
        st.stop()

    def calculate_hours_safe(current, rate):
        return [(c - 35) / r if pd.notna(c) and r and r > 0 and c > 35 else 0 for c, r in zip(current, rate)]

    hour_upper = calculate_hours_safe(upper_current, avg_rate_upper)
    hour_lower = calculate_hours_safe(lower_current, avg_rate_lower)

    st.subheader("📋 ตาราง Avg Rate - Upper")
    def style_upper(val):
        return 'color: red; font-weight: bold' if isinstance(val, float) and val > 0 else ''
    st.dataframe(upper_df.style.applymap(style_upper, subset=["Avg Rate (Upper)"]).format("{:.6f}"), use_container_width=True)

    st.subheader("📋 ตาราง Avg Rate - Lower")
    def style_lower(val):
        return 'color: red; font-weight: bold' if isinstance(val, float) and val > 0 else ''
    st.dataframe(lower_df.style.applymap(style_lower, subset=["Avg Rate (Lower)"]).format("{:.6f}"), use_container_width=True)

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


#----------------ลองแก้ดู----------------------------------------
    sheet_names = [ws.title for ws in sh.worksheets() if ws.title.lower().startswith("sheet")]
    sheet_count = st.number_input("📌 กรอกจำนวนชีตย้อนหลังที่ต้องใช้", min_value=1, max_value=len(sheet_names), value=6)
    try:
        xls = pd.ExcelFile(sheet_url)
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

        upper_df = pd.DataFrame.from_dict(upper_rates, orient='index')
        lower_df = pd.DataFrame.from_dict(lower_rates, orient='index')
        upper_df["Avg Rate (Upper)"] = upper_df.apply(avg_positive, axis=1)
        lower_df["Avg Rate (Lower)"] = lower_df.apply(avg_positive, axis=1)

        avg_rate_upper = upper_df["Avg Rate (Upper)"].tolist()[:32]
        avg_rate_lower = lower_df["Avg Rate (Lower)"].tolist()[:32]

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


        
    
    
    
    
    
    
    
    

# --------------------------------------------------- PAGE 2 -------------------------------------------------
elif page == "📝 กรอกข้อมูลแปลงถ่านเพิ่มเติม":
    st.title("📝 กรอกข้อมูลแปรงถ่าน + ชั่วโมง")

    service_account_info = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(service_account_info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    gc = gspread.authorize(creds)
    sh = gc.open_by_url("https://docs.google.com/spreadsheets/d/1SOkIH9jchaJi_0eck5UeyUR8sTn2arndQofmXv5pTdQ")

# ✅ ดึงเฉพาะชีตที่ชื่อขึ้นต้นด้วย Sheet (หรือเปลี่ยนเป็นตาม pattern ของคุณ เช่น "Sheet1", "Sheet2", ...)
    sheet_names = [ws.title for ws in sh.worksheets() if ws.title.lower().startswith("sheet")]
    selected_sheet = st.selectbox("📄 เลือก Sheet ที่ต้องการกรอกข้อมูล", sheet_names)

    ws = sh.worksheet(selected_sheet)

    hours = st.number_input("⏱️ ชั่วโมง", min_value=0.0, step=0.1)

    st.markdown("### 🔧 แปลงถ่านส่วน LOWER")
    upper = []
    cols = st.columns(8)
    for i in range(32):
        col = cols[i % 8]
        with col:
            st.markdown(f"<div style='text-align: center;'>แปลงถ่านที่ {i+1}</div>", unsafe_allow_html=True)
            value = st.text_input(f"{i+1}", key=f"u{i}", label_visibility="collapsed", placeholder="0.00")
            try:
                upper.append(float(value))
            except:
                upper.append(0.0)

    st.markdown("### 🔧 แปลงถ่านส่วน UPPER")
    lower = []
    cols = st.columns(8)
    for i in range(32):
        col = cols[i % 8]
        with col:
            st.markdown(f"<div style='text-align: center;'>แปลงถ่านที่ {i+1}</div>", unsafe_allow_html=True)
            value = st.text_input(f"{i+1}", key=f"l{i}", label_visibility="collapsed", placeholder="0.00")
            try:
                lower.append(float(value))
            except:
                lower.append(0.0)

    if st.button("📤 บันทึก"):
        try:
            ws.update("H1", [[hours]])
            ws.update("C3:C34", [[v] for v in upper])
            ws.update("F3:F34", [[v] for v in lower])
            st.success(f"✅ บันทึกลง {selected_sheet} แล้วเรียบร้อย")
        except Exception as e:
            st.error(f"❌ {e}")

    # ------------------ แสดงตารางรวม ------------------
    st.subheader("📄 ตารางรวม Upper + Lower (Current / Previous)")
    xls = pd.ExcelFile("https://docs.google.com/spreadsheets/d/1SOkIH9jchaJi_0eck5UeyUR8sTn2arndQofmXv5pTdQ/export?format=xlsx")
   
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
        st.markdown(f"📆 วันที่ Previous: **`{date_prev}`** | วันที่ Current: **`{date_curr}`**")
        st.markdown(f"#### ⏱️ ชั่วโมงจาก `{selected_view_sheet}`: `{hour_val}` ชั่วโมง")

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
    sheet_id = "1SOkIH9jchaJi_0eck5UeyUR8sTn2arndQofmXv5pTdQ"
    sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    xls = pd.ExcelFile(sheet_url)
    
    

    sheet_count = st.number_input("📌 กรอกจำนวนชีตย้อนหลังที่ต้องใช้ (1-7)", min_value=1, max_value=7, value=6)
    sheet_names = [f"Sheet{i}" for i in range(1, sheet_count + 1)]
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

    avg_rate_upper = [avg_positive(upper_rates[n]) for n in brush_numbers]
    avg_rate_lower = [avg_positive(lower_rates[n]) for n in brush_numbers]

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
    fig_upper.add_annotation(x=5, y=35, text="⚠️ Threshold 35 mm", showarrow=False, font=dict(color="firebrick", size=12), bgcolor="white")

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
    fig_lower.add_annotation(x=5, y=35, text="⚠️ Threshold 35 mm", showarrow=False, font=dict(color="firebrick", size=12), bgcolor="white")

    fig_lower.update_layout(title="🔻 ความยาว Lower ตามเวลา", xaxis_title="ชั่วโมง", yaxis_title="mm",
                            xaxis=dict(dtick=10, range=[0, 200]), yaxis=dict(range=[30, 65]))
    st.plotly_chart(fig_lower, use_container_width=True)