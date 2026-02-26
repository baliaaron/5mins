import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from datetime import datetime

st.set_page_config(page_title="醫療帳務資料合併工具", layout="wide")
st.title("🏥 醫療帳務資料合併工具")
st.markdown("請將 Excel 檔案拖至下方框中，系統將自動核對並保留原始格式。")

# --- 檔案上傳區 ---
uploaded_files = st.file_uploader("請同時選擇或拖入「主模板」與「每日來源資料」兩個檔案", type=["xlsx", "xlsm"], accept_multiple_files=True)

template_file = None
day_file = None

if uploaded_files:
    for f in uploaded_files:
        try:
            xls = pd.ExcelFile(f)
            sheet_names = xls.sheet_names
            if "代號表" in sheet_names or "工作表1" in sheet_names:
                day_file = f
            elif any(s.startswith("115") for s in sheet_names):
                template_file = f
        except Exception:
            continue

if template_file and day_file:
    st.info(f"📁 已偵測到：\n- 主模板：{template_file.name}\n- 來源資料：{day_file.name}")
    if st.button("🚀 開始執行並產生報表", type="primary"):
        with st.spinner("正在為您處理資料，請稍候..."):
            try:
                # 1. 讀取代號表
                df_codes = pd.read_excel(day_file, sheet_name="代號表")
                code_dict = {}
                for _, row in df_codes.iterrows():
                    name = str(row['名字']).strip()
                    for col in ['代號1', '代號2', '代號3']:
                        if col in df_codes.columns and pd.notna(row[col]):
                            val = str(row[col]).split('.')[0]
                            c = val.zfill(2) if val.isdigit() and len(val) < 3 else val
                            code_dict[c] = name

                # 2. 載入模板保留格式
                template_file.seek(0)
                wb = load_workbook(template_file)
                
                # 異動紀錄清單
                summary_data = []

                def safe_val(v): return float(v) if pd.notna(v) else 0
                
                def add_to_cell(ws, row, col, val, reason, name, date_obj):
                    if val == 0: return
                    curr_val = ws.cell(row=row, column=col).value
                    old_val = float(curr_val) if curr_val is not None else 0
                    ws.cell(row=row, column=col).value = old_val + val
                    summary_data.append({
                        "日期": date_obj.strftime('%Y-%m-%d'),
                        "項目": reason,
                        "對象": name,
                        "金額": val
                    })

                # 欄位映射設定
                opd_stu = {'李':(40,41,42),'珩':(43,44,45),'芳':(46,47,48),'東':(49,50,51),'澍':(52,53,54),'張明揚':(55,56,57),'李建南':(58,59,60),'影像':(64,65,66)}
                opd_no_stu = {'鄭':61, '許越涵':62, '陳思宇':63}
                room_map = {'李':85,'珩':86,'芳':87,'東':88,'澍':89,'李建南':90,'張明揚':91,'鄭':92,'陳思宇':93,'林慧雯':94}
                nurs_map = {'李':115,'珩':116,'芳':117,'東':118,'澍':119,'李建南':120,'張明揚':121,'林慧雯':122}

                target_date_str = None

                # 3. 處理 工作表1 (OPD)
                if "工作表1" in pd.ExcelFile(day_file).sheet_names:
                    df1 = pd.read_excel(day_file, sheet_name="工作表1")
                    df1['看診日期'] = pd.to_datetime(df1['看診日期'], errors='coerce')
                    
                    # 以工作表1的最新日期作為報表顯示基準
                    if not df1['看診日期'].dropna().empty:
                        target_date_str = df1['看診日期'].dropna().max().strftime('%Y-%m-%d')
                    
                    for _, row in df1.iterrows():
                        dt = row['看診日期']
                        if pd.isna(dt): continue
                        m_str = f"115{dt.month:02d}"
                        if m_str not in wb.sheetnames: continue
                        ws, r_idx = wb[m_str], dt.day + 3
                        c = str(row['醫生代碼']).strip().split('.')[0].zfill(2)
                        name = code_dict.get(c)
                        val = safe_val(row['小計']) - safe_val(row['掛號']) - safe_val(row['部份負擔'])
                        
                        if name == '兒科': add_to_cell(ws, r_idx, 70, val, "門診", name, dt)
                        elif name in opd_no_stu: add_to_cell(ws, r_idx, opd_no_stu[name], val, "門診", name, dt)
                        elif name in opd_stu:
                            s = str(row['診次']).upper()
                            idx = 0 if s=='S' else (1 if s=='T' else 2)
                            add_to_cell(ws, r_idx, opd_stu[name][idx], val, f"門診({s})", name, dt)

                # 4. 處理 工作表2 (出院)
                if "工作表2" in pd.ExcelFile(day_file).sheet_names:
                    df2 = pd.read_excel(day_file, sheet_name="工作表2")
                    for _, row in df2.iterrows():
                        if pd.isna(row['住院日期']): continue
                        dt = pd.to_datetime(row['住院日期'])
                        m_str = f"115{dt.month:02d}"
                        if m_str not in wb.sheetnames: continue
                        ws, r_idx = wb[m_str], dt.day + 3
                        c = str(row['醫生代碼']).strip().split('.')[0].zfill(2)
                        name = code_dict.get(c)
                        
                        if name and name in room_map:
                            add_to_cell(ws, r_idx, room_map[name], safe_val(row['病房費']), "病房費", name, dt)
                            add_to_cell(ws, r_idx, room_map[name]+10, safe_val(row['材料費']), "材料費", name, dt)
                            add_to_cell(ws, r_idx, room_map[name]+20, safe_val(row['伙食費']), "伙食費", name, dt)
                        
                        pre = safe_val(row['預收款'])
                        if pre != 0:
                            reason = "生產(預收)" if pre > 0 else "出院結算"
                            val = pre if pre > 0 else abs(pre)-safe_val(row['麻醉費'])-safe_val(row['產費'])
                            col = 217 if pre > 0 else 224 # 假設 217是預收款欄, 224是HP欄
                            add_to_cell(ws, r_idx, col, val, reason, name if name else "未知", dt)

                # 5. 處理 工作表3 (嬰兒室)
                if "工作表3" in pd.ExcelFile(day_file).sheet_names:
                    df3 = pd.read_excel(day_file, sheet_name="工作表3")
                    for _, row in df3.iterrows():
                        if pd.isna(row['住院日期']): continue
                        dt = pd.to_datetime(row['住院日期'])
                        m_str = f"115{dt.month:02d}"
                        if m_str not in wb.sheetnames: continue
                        ws, r_idx = wb[m_str], dt.day + 3
                        c = str(row['醫生代碼']).strip().split('.')[0].zfill(2)
                        name = code_dict.get(c)
                        if name in nurs_map:
                            add_to_cell(ws, r_idx, nurs_map[name], safe_val(row['小計']), "嬰兒室", name, dt)

                # 6. 處理 工作表4 (欠款)
                if "工作表4" in pd.ExcelFile(day_file).sheet_names:
                    df4 = pd.read_excel(day_file, sheet_name="工作表4")
                    date_col = next((col for col in df4.columns if '日期' in str(col)), df4.columns[0])
                    for _, row in df4.iterrows():
                        if pd.isna(row[date_col]): continue
                        dt = pd.to_datetime(row[date_col])
                        m_str = f"115{dt.month:02d}"
                        if m_str not in wb.sheetnames: continue
                        ws, r_idx = wb[m_str], dt.day + 3
                        val = safe_val(row['未收額'])
                        add_to_cell(ws, r_idx, 135, val, "欠款(未收額)", "全科", dt)

                # 7. 處理 工作表5 (還款)
                if "工作表5" in pd.ExcelFile(day_file).sheet_names:
                    df5 = pd.read_excel(day_file, sheet_name="工作表5")
                    date_col = next((col for col in df5.columns if '日期' in str(col)), df5.columns[0])
                    for _, row in df5.iterrows():
                        if pd.isna(row[date_col]): continue
                        dt = pd.to_datetime(row[date_col])
                        m_str = f"115{dt.month:02d}"
                        if m_str not in wb.sheetnames: continue
                        ws, r_idx = wb[m_str], dt.day + 3
                        val = safe_val(row['還款金額'])
                        add_to_cell(ws, r_idx, 123, val, "還款", "全科", dt)

                # 8. 製作下載檔案
                out = io.BytesIO()
                wb.save(out)
                processed_output = out.getvalue()

                st.success("✅ 處理完成！所有資料均已成功寫入。")
                st.download_button(label="💾 下載結果檔案", data=processed_output, file_name=f"合併結果_{datetime.now().strftime('%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

                # --- 顯示最新日期摘要表格 ---
                st.divider()
                if target_date_str:
                    st.subheader(f"📊 今日更動摘要 ({target_date_str})")
                    report_df = pd.DataFrame(summary_data)
                    if not report_df.empty:
                        # 只篩選「目標日期」的資料顯示
                        latest_report = report_df[report_df['日期'] == target_date_str]
                        if not latest_report.empty:
                            final_table = latest_report.groupby(['項目', '對象'])['金額'].sum().reset_index()
                            final_table['金額'] = final_table['金額'].apply(lambda x: f"{x:,.0f}")
                            st.table(final_table)
                        else:
                            st.info(f"{target_date_str} 無更動數據。")
                    else:
                        st.warning("未偵測到任何異動資料。")
                else:
                    st.info("無法辨識有效日期，已完成檔案處理但無摘要可顯示。")

            except Exception as e:
                st.error(f"發生錯誤: {e}")