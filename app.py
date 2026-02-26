import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from datetime import datetime

st.set_page_config(page_title="醫療帳務資料合併工具", layout="wide")
st.title("🏥 醫療帳務資料合併工具")
st.markdown("請將 Excel 檔案拖至下方框中，系統將自動核對並保留原始格式。")

# --- 檔案上傳區 ---
col1, col2 = st.columns(2)
with col1:
    template_file = st.file_uploader("1. 拖入主模板 (115年度明細表新.xlsx)", type=["xlsx", "xlsm"])
with col2:
    day_file = st.file_uploader("2. 拖入每日來源資料 (day.xlsx)", type=["xlsx", "xlsm"])

if template_file and day_file:
    if st.button("🚀 開始執行並產生報表", type="primary"):
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

            # 3. 處理 工作表1 (OPD)
            df1 = pd.read_excel(day_file, sheet_name="工作表1")
            df1['看診日期'] = pd.to_datetime(df1['看診日期'])
            latest_date = df1['看診日期'].max()
            
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
            df2 = pd.read_excel(day_file, sheet_name="工作表2")
            for _, row in df2.iterrows():
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
            df3 = pd.read_excel(day_file, sheet_name="工作表3")
            for _, row in df3.iterrows():
                dt = pd.to_datetime(row['住院日期'])
                m_str = f"115{dt.month:02d}"
                if m_str not in wb.sheetnames: continue
                ws, r_idx = wb[m_str], dt.day + 3
                c = str(row['醫生代碼']).strip().split('.')[0].zfill(2)
                name = code_dict.get(c)
                if name in nurs_map:
                    add_to_cell(ws, r_idx, nurs_map[name], safe_val(row['小計']), "嬰兒室", name, dt)

            # 6. 製作下載檔案
            out = io.BytesIO()
            wb.save(out)
            processed_output = out.getvalue()

            st.success("✅ 處理完成！")
            st.download_button(label="💾 下載結果檔案", data=processed_output, file_name=f"合併結果_{datetime.now().strftime('%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

            # --- 顯示最新日期摘要表格 ---
            st.divider()
            st.subheader(f"📊 最新日期更動摘要 ({latest_date.strftime('%Y-%m-%d')})")
            
            report_df = pd.DataFrame(summary_data)
            if not report_df.empty:
                latest_report = report_df[report_df['日期'] == latest_date.strftime('%Y-%m-%d')]
                if not latest_report.empty:
                    final_table = latest_report.groupby(['項目', '對象'])['金額'].sum().reset_index()
                    # 格式化金額顯示
                    final_table['金額'] = final_table['金額'].apply(lambda x: f"{x:,.0f}")
                    st.table(final_table)
                else:
                    st.info("最新日期無更動數據。")
            else:
                st.warning("未偵測到任何異動資料，請檢查代號表與日期。")

        except Exception as e:
            st.error(f"發生錯誤: {e}")
