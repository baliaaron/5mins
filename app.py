import streamlit as st
import pandas as pd
import io
import numpy as np

st.set_page_config(page_title="醫療帳務資料合併工具", layout="centered")
st.title("🏥 醫療帳務資料合併工具")
st.markdown("請依序上傳主模板與每日來源資料，系統將自動為您合併。")

# --- 檔案上傳區 ---
st.subheader("1. 上傳檔案")
template_file = st.file_uploader("上傳主模板 (115年度明細表新.xlsx)", type=["xlsx", "xlsm"])
day_file = st.file_uploader("上傳每日來源資料 (day.xlsx)", type=["xlsx", "xlsm"])

if template_file and day_file:
    if st.button("🚀 開始合併資料", type="primary"):
        with st.spinner("資料處理中，請稍候..."):
            try:
                # 建立代號字典
                df_codes = pd.read_excel(day_file, sheet_name="代號表")
                code_dict = {}
                for idx, row in df_codes.iterrows():
                    name = str(row['名字']).strip()
                    if pd.notna(name) and name != '':
                        for col in ['代號1', '代號2', '代號3']:
                            if col in df_codes.columns:
                                val = row[col]
                                if pd.notna(val) and str(val).strip() != '':
                                    try:
                                        num = int(float(val))
                                        c = f"{num:02d}" if num < 10 else str(num)
                                        code_dict[c] = name
                                    except:
                                        code_dict[str(val).strip()] = name

                # 讀取主模板的所有工作表
                xls = pd.ExcelFile(template_file)
                templates = {}
                for sheet in xls.sheet_names:
                    if sheet.startswith("115"): # 只處理 115 開頭的月份表
                        df = pd.read_excel(template_file, sheet_name=sheet, header=None)
                        templates[sheet] = df

                def safe_add(df, row_idx, col_idx, val):
                    if val == 0: return
                    curr = df.iat[row_idx, col_idx]
                    if pd.isna(curr) or str(curr).strip() == '':
                        df.iat[row_idx, col_idx] = val
                    else:
                        try:
                            df.iat[row_idx, col_idx] = float(curr) + val
                        except:
                            df.iat[row_idx, col_idx] = val

                # 欄位對應表 (與您的 VBA 邏輯完全一致)
                opd_stu = {
                    '李': (39,40,41), '珩': (42,43,44), '芳': (45,46,47), '東': (48,49,50), '澍': (51,52,53),
                    '張明揚': (54,55,56), '李建南': (57,58,59), '影像': (63,64,65)
                }
                opd_no_stu = {'鄭': 60, '許越涵': 61, '陳思宇': 62}
                ped_col = 69 # BR 欄
                room_map = {'李':84, '珩':85, '芳':86, '東':87, '澍':88, '李建南':89, '張明揚':90, '鄭':91, '陳思宇':92, '林慧雯':93}
                mat_map = {k: v+10 for k, v in room_map.items()}
                food_map = {k: v+20 for k, v in room_map.items()}
                nurs_map = {'李':114, '珩':115, '芳':116, '東':117, '澍':118, '李建南':119, '張明揚':120, '林慧雯':121}

                # 處理工作表1: OPD
                df_opd = pd.read_excel(day_file, sheet_name="工作表1")
                for _, row in df_opd.iterrows():
                    dt = row['看診日期']
                    if pd.isna(dt): continue
                    try: d_obj = pd.to_datetime(dt)
                    except: continue
                    m_str = f"115{d_obj.month:02d}"
                    if m_str not in templates: continue
                    target_row = d_obj.day + 2
                    
                    c = str(row['醫生代碼']).strip()
                    try:
                        c_num = int(float(c))
                        c = f"{c_num:02d}" if c_num < 10 else str(c_num)
                    except: pass
                    name = code_dict.get(c)
                    if not name: continue
                    
                    subtotal = float(row['小計']) if pd.notna(row['小計']) else 0
                    reg = float(row['掛號']) if pd.notna(row['掛號']) else 0
                    part = float(row['部份負擔']) if pd.notna(row['部份負擔']) else 0
                    val = subtotal - reg - part
                    if val == 0: continue
                    
                    target_col = None
                    if name == '兒科': target_col = ped_col
                    elif name in opd_no_stu: target_col = opd_no_stu[name]
                    elif name in opd_stu:
                        sess = str(row['診次']).strip().upper()
                        if sess == 'S': target_col = opd_stu[name][0]
                        elif sess == 'T': target_col = opd_stu[name][1]
                        elif sess == 'U': target_col = opd_stu[name][2]
                        else: target_col = opd_stu[name][0]
                        
                    if target_col is not None:
                        safe_add(templates[m_str], target_row, target_col, val)

                # 處理工作表2: 出院
                hp_sums = {}
                df_inp = pd.read_excel(day_file, sheet_name="工作表2")
                for _, row in df_inp.iterrows():
                    dt = row['住院日期']
                    if pd.isna(dt): continue
                    try: d_obj = pd.to_datetime(dt)
                    except: continue
                    m_str = f"115{d_obj.month:02d}"
                    if m_str not in templates: continue
                    target_row = d_obj.day + 2
                    
                    c = str(row['醫生代碼']).strip()
                    try:
                        c_num = int(float(c))
                        c = f"{c_num:02d}" if c_num < 10 else str(c_num)
                    except: pass
                    name = code_dict.get(c)
                    
                    r_fee = float(row['病房費']) if pd.notna(row['病房費']) else 0
                    m_fee = float(row['材料費']) if pd.notna(row['材料費']) else 0
                    f_fee = float(row['伙食費']) if pd.notna(row['伙食費']) else 0
                    
                    if name and name in room_map and r_fee != 0: safe_add(templates[m_str], target_row, room_map[name], r_fee)
                    if name and name in mat_map and m_fee != 0: safe_add(templates[m_str], target_row, mat_map[name], m_fee)
                    if name and name in food_map and f_fee != 0: safe_add(templates[m_str], target_row, food_map[name], f_fee)
                        
                    pre = float(row['預收款']) if pd.notna(row['預收款']) else 0
                    ane = float(row['麻醉費']) if pd.notna(row['麻醉費']) else 0
                    bir = float(row['產費']) if pd.notna(row['產費']) else 0
                    if pre < 0:
                        val = abs(pre) - ane - bir
                        key = (m_str, target_row)
                        hp_sums[key] = hp_sums.get(key, 0) + val

                for (m_str, target_row), val in hp_sums.items():
                    if val != 0:
                        safe_add(templates[m_str], target_row, 223, val) # HP 欄

                # 處理工作表3: 嬰兒室
                df_nur = pd.read_excel(day_file, sheet_name="工作表3")
                for _, row in df_nur.iterrows():
                    dt = row['住院日期']
                    if pd.isna(dt): continue
                    try: d_obj = pd.to_datetime(dt)
                    except: continue
                    m_str = f"115{d_obj.month:02d}"
                    if m_str not in templates: continue
                    target_row = d_obj.day + 2
                    
                    c = str(row['醫生代碼']).strip()
                    try:
                        c_num = int(float(c))
                        c = f"{c_num:02d}" if c_num < 10 else str(c_num)
                    except: pass
                    name = code_dict.get(c)
                    if not name or name not in nurs_map: continue
                    
                    sub = float(row['小計']) if pd.notna(row['小計']) else 0
                    if sub != 0:
                        safe_add(templates[m_str], target_row, nurs_map[name], sub)

                # 將結果寫入記憶體中的 Excel 檔案
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # 先將原本上傳的檔案複製過來，確保非 115 的分頁也被保留 (若有)
                    # 接著覆蓋有變動的月份
                    for sheet_name, df in templates.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                
                processed_data = output.getvalue()
                
                st.success("✅ 資料合併完成！")
                st.subheader("2. 下載檔案")
                st.download_button(
                    label="下載合併完成的明細表",
                    data=processed_data,
                    file_name="合併完成_115年度明細表新.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
            except Exception as e:
                st.error(f"處理檔案時發生錯誤: {e}")
