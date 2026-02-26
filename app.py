import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from datetime import datetime
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="醫療帳務資料合併工具 - 精準座標版", layout="wide")
st.title("🏥 醫療帳務資料合併工具")
st.markdown("本版本已更新：預收款 >= 0 時，自動拆分至醫師生產實收欄位 (BX+)。")

# --- 初始化 Session State ---
# 用於在點擊下載後保留表格顯示
if 'processed_output' not in st.session_state:
    st.session_state.processed_output = None
if 'detailed_records' not in st.session_state:
    st.session_state.detailed_records = []
if 'target_date_str' not in st.session_state:
    st.session_state.target_date_str = None
if 'data_pool' not in st.session_state:
    st.session_state.data_pool = {}

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
    
    # 點擊「開始」會執行運算並存入 session_state
    if st.button("🚀 開始精準合併並對帳", type="primary"):
        with st.spinner("正在執行對帳座標運算，包含生產實收拆分..."):
            try:
                # 1. 讀取代號表
                df_codes = pd.read_excel(day_file, sheet_name="代號表")
                code_dict = {}
                for _, row in df_codes.iterrows():
                    name = str(row.iloc[0]).strip()
                    for i in range(1, len(row)):
                        if pd.notna(row.iloc[i]):
                            val = str(row.iloc[i]).split('.')[0]
                            c = val.zfill(2) if val.isdigit() and len(val) < 3 else val
                            code_dict[c] = name

                # 2. 建立資料彙整池
                st.session_state.data_pool = {}
                st.session_state.detailed_records = []
                st.session_state.target_date_str = None
                
                def collect_data(date_obj, col, val, reason, name):
                    if val == 0: return
                    d_str = date_obj.strftime('%Y-%m-%d')
                    key = (d_str, col)
                    old_v, _, _ = st.session_state.data_pool.get(key, (0.0, "", ""))
                    st.session_state.data_pool[key] = (old_v + val, reason, name)
                    st.session_state.detailed_records.append({
                        "日期": d_str, "醫師/對象": name, "欄位編號": col, "項目內容": reason, "金額": val
                    })

                # 座標地圖
                opd_stu = {'李':(40,41,42),'珩':(43,44,45),'芳':(46,47,48),'東':(49,50,51),'澍':(52,53,54),'張明揚':(55,56,57),'李建南':(58,59,60),'影像':(64,65,66)}
                opd_no_stu = {'鄭':61, '許越涵':62, '陳思宇':63}
                birth_map = {'李':76,'珩':77,'芳':78,'東':79,'澍':80,'李建南':81,'張明揚':82,'鄭':83,'陳思宇':84}
                room_map = {'李':85,'珩':86,'芳':87,'東':88,'澍':89,'李建南':90,'張明揚':91,'鄭':92,'陳思宇':93,'林慧雯':94}
                nurs_map = {'李':115,'珩':116,'芳':117,'東':118,'澍':119,'李建南':120,'張明揚':121,'林慧雯':122}

                def safe_num(v):
                    try: return float(v) if pd.notna(v) else 0.0
                    except: return 0.0

                day_xls = pd.ExcelFile(day_file)
                all_sheets = day_xls.sheet_names

                # 3. 工作表1 (門診)
                if "工作表1" in all_sheets:
                    df1 = pd.read_excel(day_file, sheet_name="工作表1", header=None, skiprows=1)
                    for _, row in df1.iterrows():
                        dt = pd.to_datetime(row.iloc[0], errors='coerce')
                        if pd.isna(dt): continue
                        if st.session_state.target_date_str is None or dt.strftime('%Y-%m-%d') > st.session_state.target_date_str:
                            st.session_state.target_date_str = dt.strftime('%Y-%m-%d')
                        c = str(row.iloc[1]).strip().split('.')[0].zfill(2)
                        name = code_dict.get(c)
                        val = safe_num(row.iloc[16]) - safe_num(row.iloc[4]) - safe_num(row.iloc[5])
                        if name == '兒科': collect_data(dt, 70, val, "門診(兒科)", "兒科")
                        elif name in opd_no_stu: collect_data(dt, opd_no_stu[name], val, "門診", name)
                        elif name in opd_stu:
                            s = str(row.iloc[2]).strip().upper()
                            s_idx = 0 if s=='S' else (1 if s=='T' else 2)
                            label = {'S':'早', 'T':'午', 'U':'晚'}.get(s, s)
                            collect_data(dt, opd_stu[name][s_idx], val, f"門診({label})", name)

                # 4. 工作表2 (出院)
                if "工作表2" in all_sheets:
                    df2 = pd.read_excel(day_file, sheet_name="工作表2", header=None, skiprows=1)
                    hp_agg = {}
                    for _, row in df2.iterrows():
                        dt = pd.to_datetime(row.iloc[0], errors='coerce')
                        if pd.isna(dt): continue
                        c = str(row.iloc[2]).strip().split('.')[0].zfill(2)
                        name = code_dict.get(c, "其他")
                        iAnes, iRoom, iBirth, iMat, iPre, iFood = safe_num(row.iloc[7]), safe_num(row.iloc[8]), safe_num(row.iloc[9]), safe_num(row.iloc[10]), safe_num(row.iloc[11]), safe_num(row.iloc[12])
                        if name in room_map:
                            collect_data(dt, room_map[name], iRoom, "病房費", name)
                            collect_data(dt, room_map[name]+10, iMat, "材料費", name)
                            collect_data(dt, room_map[name]+20, iFood, "伙食費", name)
                        if iPre >= 0:
                            birth_total = iAnes + iBirth + iPre
                            if birth_total != 0 and name in birth_map:
                                collect_data(dt, birth_map[name], birth_total, "生產實收(麻+產+預)", name)
                        else:
                            hp_val = abs(iPre) - iAnes - iBirth
                            d_str = dt.strftime('%Y-%m-%d')
                            hp_agg[d_str] = hp_agg.get(d_str, 0.0) + hp_val
                    for d_str, total in hp_agg.items():
                        if total != 0: collect_data(datetime.strptime(d_str, '%Y-%m-%d'), 224, total, "出院結算(HP)", "總計")

                # 5. 工作表3 (嬰兒室)
                if "工作表3" in all_sheets:
                    df3 = pd.read_excel(day_file, sheet_name="工作表3", header=None, skiprows=1)
                    for _, row in df3.iterrows():
                        dt = pd.to_datetime(row.iloc[0], errors='coerce')
                        if pd.isna(dt): continue
                        c = str(row.iloc[2]).strip().split('.')[0].zfill(2)
                        val = safe_num(row.iloc[6])
                        name = code_dict.get(c)
                        if name in nurs_map: collect_data(dt, nurs_map[name], val, "嬰兒室費用", name)

                # 6 & 7. 欠款與還款
                for sheet, col_keyword, label, target_col in [("工作表4", "未收額", "今日欠款", 135), ("工作表5", "還款金額", "今日還款", 123)]:
                    if sheet in all_sheets:
                        tmp = pd.read_excel(day_file, sheet_name=sheet)
                        dt_col = next((c for c in tmp.columns if '日期' in str(c)), tmp.columns[0])
                        val_col = next((c for c in tmp.columns if col_keyword in str(c)), None)
                        if val_col:
                            for _, row in tmp.iterrows():
                                dt = pd.to_datetime(row[dt_col], errors='coerce')
                                if pd.isna(dt): continue
                                collect_data(dt, target_col, safe_num(row[val_col]), label, "總計")

                # --- 寫入 Excel ---
                template_file.seek(0)
                wb = load_workbook(template_file)
                for (d_str, col), (val, reason, name) in st.session_state.data_pool.items():
                    dt = datetime.strptime(d_str, '%Y-%m-%d')
                    m_key = f"115{dt.month:02d}"
                    if m_key in wb.sheetnames: wb[m_key].cell(row=dt.day + 3, column=col).value = val

                out = io.BytesIO()
                wb.save(out)
                st.session_state.processed_output = out.getvalue()
                st.success("✅ 運算完成！")

            except Exception as e:
                st.error(f"發生錯誤: {e}")
                st.exception(e)

# --- 顯示結果區域 (受 Session State 保護，下載後不消失) ---
if st.session_state.processed_output is not None:
    st.divider()
    st.download_button(
        label="💾 下載結果檔案", 
        data=st.session_state.processed_output, 
        file_name=f"{datetime.now().strftime('%Y%m%d')}_財務對帳版.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
        type="primary"
    )

    if st.session_state.target_date_str:
        st.header(f"📊 詳細對帳單 ({st.session_state.target_date_str})")
        # 直接從 session_state 的 pool 提取當日資料
        day_pool = {k: v for k, v in st.session_state.data_pool.items() if k[0] == st.session_state.target_date_str}
        if day_pool:
            final_list = []
            for (d, c), (v, r, n) in day_pool.items():
                final_list.append({"醫師/對象": n, "項目名稱": r, "Excel欄位": f"{get_column_letter(c)} ({c})", "金額": v, "編號": c})
            display_df = pd.DataFrame(final_list).sort_values(by=['醫師/對象', '編號'])
            display_df['金額'] = display_df['金額'].apply(lambda x: f"{x:,.0f}")
            st.dataframe(display_df[['醫師/對象', '項目名稱', 'Excel欄位', '金額']], use_container_width=True, hide_index=True)
            st.info("💡 提示：表格已鎖定，您可以放心地點擊上方按鈕下載檔案，表格不會消失。")
        else:
            st.warning("當日無異動。")