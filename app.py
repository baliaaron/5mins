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
        with st.spinner("正在依照規則處理資料，請稍候..."):
            try:
                # 1. 讀取代號表 (規則 1)
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
                
                summary_data = []
                target_date_str = None

                def safe_val(v): return float(v) if pd.notna(v) else 0
                
                def add_to_cell(ws, row, col, val, reason, name, date_obj, category):
                    if val == 0: return
                    curr_val = ws.cell(row=row, column=col).value
                    old_val = float(curr_val) if curr_val is not None else 0
                    ws.cell(row=row, column=col).value = old_val + val
                    summary_data.append({
                        "日期": date_obj.strftime('%Y-%m-%d'),
                        "分類": category,
                        "項目": reason,
                        "對象": name,
                        "金額": val
                    })

                # 欄位映射設定
                opd_stu = {'李':(40,41,42),'珩':(43,44,45),'芳':(46,47,48),'東':(49,50,51),'澍':(52,53,54),'張明揚':(55,56,57),'李建南':(58,59,60),'影像':(64,65,66)}
                opd_no_stu = {'鄭':61, '許越涵':62, '陳思宇':63}
                room_map = {'李':85,'珩':86,'芳':87,'東':88,'澍':89,'李建南':90,'張明揚':91,'鄭':92,'陳思宇':93,'林慧雯':94}
                nurs_map = {'李':115,'珩':116,'芳':117,'東':118,'澍':119,'李建南':120,'張明揚':121,'林慧雯':122}

                day_xls_info = pd.ExcelFile(day_file)
                all_day_sheets = day_xls_info.sheet_names

                # 3. 處理 工作表1 (門診 - 規則 2)
                if "工作表1" in all_day_sheets:
                    df1 = pd.read_excel(day_file, sheet_name="工作表1")
                    df1['看診日期'] = pd.to_datetime(df1['看診日期'], errors='coerce')
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
                        
                        if name == '兒科': add_to_cell(ws, r_idx, 70, val, "兒科", name, dt, "1. 門診收入")
                        elif name in opd_no_stu: add_to_cell(ws, r_idx, opd_no_stu[name], val, "不分診", name, dt, "1. 門診收入")
                        elif name in opd_stu:
                            s = str(row['診次']).upper()
                            # 映射 S->早, T->午, U->晚
                            s_map = {'S':'早', 'T':'午', 'U':'晚'}
                            ss = s_map.get(s, s)
                            idx = 0 if s=='S' else (1 if s=='T' else 2)
                            add_to_cell(ws, r_idx, opd_stu[name][idx], val, ss, name, dt, "1. 門診收入")

                # 4. 處理 工作表2 (出院 - 規則 3)
                if "工作表2" in all_day_sheets:
                    df2 = pd.read_excel(day_file, sheet_name="工作表2")
                    for _, row in df2.iterrows():
                        if pd.isna(row['住院日期']): continue
                        dt = pd.to_datetime(row['住院日期'])
                        m_str = f"115{dt.month:02d}"
                        if m_str not in wb.sheetnames: continue
                        ws, r_idx = wb[m_str], dt.day + 3
                        c = str(row['醫生代碼']).strip().split('.')[0].zfill(2)
                        name = code_dict.get(c, "其他")
                        
                        if name in room_map:
                            add_to_cell(ws, r_idx, room_map[name], safe_val(row['病房費']), "病房費", name, dt, "2. 住院明細")
                            add_to_cell(ws, r_idx, room_map[name]+10, safe_val(row['材料費']), "材料費", name, dt, "2. 住院明細")
                            add_to_cell(ws, r_idx, room_map[name]+20, safe_val(row['伙食費']), "伙食費", name, dt, "2. 住院明細")
                        
                        pre = safe_val(row['預收款'])
                        if pre != 0:
                            reason = "生產(預收)" if pre > 0 else "出院結算"
                            val = pre if pre > 0 else abs(pre)-safe_val(row['麻醉費'])-safe_val(row['產費'])
                            col = 217 if pre > 0 else 224
                            add_to_cell(ws, r_idx, col, val, reason, "總計", dt, "3. 財務結算")

                # 5. 處理 工作表3 (嬰兒室 - 規則 4)
                if "工作表3" in all_day_sheets:
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
                            add_to_cell(ws, r_idx, nurs_map[name], safe_val(row['小計']), "嬰兒室", name, dt, "2. 住院明細")

                # 6. 處理 工作表4 (欠款 - 規則 5)
                if "工作表4" in all_day_sheets:
                    df4 = pd.read_excel(day_file, sheet_name="工作表4")
                    date_col = next((col for col in df4.columns if '日期' in str(col)), df4.columns[0])
                    for _, row in df4.iterrows():
                        if pd.isna(row[date_col]): continue
                        dt = pd.to_datetime(row[date_col])
                        m_str = f"115{dt.month:02d}"
                        if m_str not in wb.sheetnames: continue
                        ws, r_idx = wb[m_str], dt.day + 3
                        val = safe_val(row['未收額'])
                        add_to_cell(ws, r_idx, 135, val, "今日欠款", "總計", dt, "3. 財務結算")

                # 7. 處理 工作表5 (還款 - 規則 6)
                if "工作表5" in all_day_sheets:
                    df5 = pd.read_excel(day_file, sheet_name="工作表5")
                    date_col = next((col for col in df5.columns if '日期' in str(col)), df5.columns[0])
                    for _, row in df5.iterrows():
                        if pd.isna(row[date_col]): continue
                        dt = pd.to_datetime(row[date_col])
                        m_str = f"115{dt.month:02d}"
                        if m_str not in wb.sheetnames: continue
                        ws, r_idx = wb[m_str], dt.day + 3
                        val = safe_val(row['還款金額'])
                        add_to_cell(ws, r_idx, 123, val, "今日還款", "總計", dt, "3. 財務結算")

                # 8. 製作下載檔案
                out = io.BytesIO()
                wb.save(out)
                processed_output = out.getvalue()

                st.success("✅ 處理完成！")
                st.download_button(label="💾 下載結果檔案", data=processed_output, file_name=f"對帳用_醫療帳務_{datetime.now().strftime('%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

                # --- 專為對帳設計的摘要表格 ---
                st.divider()
                if target_date_str:
                    st.header(f"📊 對帳摘要：{target_date_str}")
                    report_df = pd.DataFrame(summary_data)
                    day_report = report_df[report_df['日期'] == target_date_str]
                    
                    if not day_report.empty:
                        # --- 1. 門診對帳表 (橫向診次) ---
                        st.subheader("① 門診收入對帳 (OPD 早/午/晚)")
                        opd_df = day_report[day_report['分類'] == "1. 門診收入"]
                        if not opd_df.empty:
                            opd_pivot = opd_df.pivot_table(
                                index='對象', 
                                columns='項目', 
                                values='金額', 
                                aggfunc='sum', 
                                fill_value=0
                            )
                            # 確保順序 早->午->晚
                            cols = [c for c in ['早', '午', '晚', '兒科', '不分診'] if c in opd_pivot.columns]
                            opd_pivot = opd_pivot[cols]
                            opd_pivot['總計'] = opd_pivot.sum(axis=1)
                            st.table(opd_pivot.style.format("{:,.0f}"))
                        else:
                            st.info("今日無門診異動。")

                        # --- 2. 住院費用對帳表 ---
                        st.subheader("② 住院與嬰兒室明細")
                        ipd_df = day_report[day_report['分類'] == "2. 住院明細"]
                        if not ipd_df.empty:
                            ipd_pivot = ipd_df.pivot_table(
                                index='對象', 
                                columns='項目', 
                                values='金額', 
                                aggfunc='sum', 
                                fill_value=0
                            )
                            # 確保順序
                            cols = [c for c in ['病房費', '材料費', '伙食費', '嬰兒室'] if c in ipd_pivot.columns]
                            ipd_pivot = ipd_pivot[cols]
                            ipd_pivot['總計'] = ipd_pivot.sum(axis=1)
                            st.table(ipd_pivot.style.format("{:,.0f}"))
                        else:
                            st.info("今日無住院相關費用。")

                        # --- 3. 財務結算加總 (欠款、還款、預收) ---
                        st.subheader("③ 財務與結算總額")
                        fin_df = day_report[day_report['分類'] == "3. 財務結算"]
                        if not fin_df.empty:
                            fin_summary = fin_df.groupby('項目')['金額'].sum().reset_index()
                            fin_summary.columns = ['項目名稱', '當日總額']
                            st.table(fin_summary.set_index('項目名稱').style.format("{:,.0f}"))
                        else:
                            st.info("今日無財務結算異動。")
                    else:
                        st.warning("偵測日期範圍內無異動資料。")
                else:
                    st.info("未偵測到有效日期數據，請檢查 Excel 內容。")

            except Exception as e:
                st.error(f"發生錯誤: {e}")
                st.exception(e)