import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys
import zipfile
import re
import unicodedata
import calendar
import datetime

def run_scheduling(input_path):
    """接收选中的Excel完整路径，执行全部排班逻辑"""
    work_dir = os.path.dirname(input_path)
    if work_dir:
        os.chdir(work_dir)
    # =====================下面全部是第二段内容=====================
    from ortools.sat.python import cp_model
    import pandas as pd

    def group_consecutive(days_list):
        if not days_list:
            return []
        days_list = sorted(days_list)
        blocks = []
        cur = [days_list[0]]
        for d in days_list[1:]:
            if d == cur[-1] + 1:
                cur.append(d)
            else:
                blocks.append(cur)
                cur = [d]
        blocks.append(cur)
        return blocks

    def resolve_sheet_index(file_path, keyword, default_index):
        """按工作表名字定位，找不到回退到默认序号"""
        try:
            with zipfile.ZipFile(file_path, 'r') as z:
                xml = z.read('xl/workbook.xml').decode('utf-8')
            names = re.findall(r'<sheet[^>]*?name="([^"]+)"', xml)
            for i, nm in enumerate(names):
                if nm.strip().replace(' ', '').lower().startswith(keyword.lower()):
                    return i
        except Exception:
            pass
        return default_index

    def safe_read_excel(file_path, sheet_index):
        try:
            return pd.read_excel(file_path, sheet_name=sheet_index)
        except TypeError as e:
            if "DataValidation" in str(e) or "id" in str(e):
                print(f"⚠️ 检测到 WPS 特殊格式，正在暴力清洗格式并读取工作表...")
                temp_file = file_path.replace(".xlsx", "_temp.xlsx")
                with zipfile.ZipFile(file_path, 'r') as zin:
                    with zipfile.ZipFile(temp_file, 'w') as zout:
                        for item in zin.infolist():
                            buffer = zin.read(item.filename)
                            if item.filename.startswith('xl/worksheets/') and item.filename.endswith('.xml'):
                                text = buffer.decode('utf-8')
                                text = re.sub(r'<dataValidations[^>]*>.*?</dataValidations>', '', text, flags=re.DOTALL)
                                text = re.sub(r'<dataValidation[^>]* />', '', text)
                                text = re.sub(r'<extLst>.*?</extLst>', '', text, flags=re.DOTALL)
                                buffer = text.encode('utf-8')
                            zout.writestr(item, buffer)
                df = pd.read_excel(temp_file, sheet_name=sheet_index)
                os.remove(temp_file)
                return df
            else:
                raise e

    model = cp_model.CpModel()
    print("正在读取选中的Excel文件 ...")
    try:
        df_config = safe_read_excel(input_path, sheet_index=resolve_sheet_index(input_path, "Sheet4", 3))
        config = {}
        for index, row in df_config.iterrows():
            key = str(row.iloc[0]).strip()
            if pd.isna(row.iloc[0]) or key == "": continue
            val_b = row.iloc[1] if len(row) > 1 else None
            val_c = row.iloc[2] if len(row) > 2 else None
            val_d = row.iloc[3] if len(row) > 3 else None
            config[key] = (val_b, val_c, val_d)
    except Exception as e:
        print(f"⚠️ 未找到 [Sheet4] 配置文件，将使用默认参数... ({e})")
        config = {}
        df_config = pd.DataFrame()

    def get_cfg_val(key, default, col=0):
        if key in config:
            val = config[key][col]
            if pd.notna(val):
                try: return float(val)
                except: return val
        return default

    def parse_list(val, default):
        if pd.isna(val) or str(val).strip() == "":
            return default
        try:
            res = [int(float(x.strip())) for x in str(val).replace('，', ',').split(',') if x.strip() != ""]
            return res if res else default
        except:
            return default

    if df_config.empty:
        print("❌ 错误：配置工作表为空！")
        return False
    YEAR = int(get_cfg_val("年份", 2026, 0))
    MONTH = int(get_cfg_val("月份", 11, 0))
    A_min = int(get_cfg_val("A班人数", 2, 0))
    A_max = int(get_cfg_val("A班人数", 2, 1))
    B_min = int(get_cfg_val("B班人数", 0, 0))
    B_max = int(get_cfg_val("B班人数", 14, 1))
    C_min = int(get_cfg_val("C班人数", 0, 0))
    C_max = int(get_cfg_val("C班人数", 14, 1))
    D_min = int(get_cfg_val("D班人数", 1, 0))
    D_max = int(get_cfg_val("D班人数", 1, 1))
    OFF_min = int(get_cfg_val("休息人数下限", 0, 0))
    OFF_max = int(get_cfg_val("休息人数上限", 6, 0))
    resignation_dict = {}
    for index, row in df_config.iterrows():
        key = str(row.iloc[0]).strip()
        if key in ["本月离职一", "本月离职二"]:
            if len(row) > 2 and pd.notna(row.iloc[1]) and pd.notna(row.iloc[2]):
                emp_name = str(row.iloc[1]).strip()
                try:
                    emp_day = int(float(row.iloc[2]))
                    resignation_dict[emp_name] = emp_day
                except: pass
    a_shift_name = str(get_cfg_val("A班名称", "A班", 0))
    b_shift_name = str(get_cfg_val("B班名称", "B班", 0))
    d_shift_name = str(get_cfg_val("D班名称", "D班", 0))
    A_gap_max = int(get_cfg_val("A班累计极差上限", 1, 0))
    D_gap_max = int(get_cfg_val("D班累计极差上限", 1, 0))
    AD_gap_max = int(get_cfg_val("AD班总和累计极差上限", 2, 0))
    B_gap_max = int(get_cfg_val("B班累计极差上限", 3, 0))
    score_B_gap = int(get_cfg_val("B班人数", 300, 2))
    score_C_gap = int(get_cfg_val("C班人数", 300, 2))
    max_emp_penalty = int(get_cfg_val("单员工扣分上限", 0, 0))
    force_female_less_D = str(get_cfg_val("是否强制女员工关门次数少于等于男员工", "否", 0)).strip() == "是"

    # =========新增：连续上班天数配置=========
    consec_work_min = int(get_cfg_val("连续上班天数", 4, 0))
    consec_work_max = int(get_cfg_val("连续上班天数", 8, 1))
    # 0=关闭约束
    if consec_work_min <= 0:
        consec_work_min = 0
    if consec_work_max <=0:
        consec_work_max = 0

    print(f"📋连续上班配置：下限={consec_work_min}天，上限={consec_work_max}天")
    if consec_work_min>0 and consec_work_max>0 and consec_work_min>consec_work_max:
        print("⚠️警告：连续上班下限大于上限！参数不合理！")

    num_days = calendar.monthrange(YEAR, MONTH)[1]
    start_date = datetime.date(YEAR, MONTH, 1)
    days = range(1, num_days + 1)
    shift_types = ["A", "B", "C", "D", "Z", "休"]
    holiday_dict = {}
    for index, row in df_config.iterrows():
        key = str(row.iloc[0]).strip()
        if key.startswith("法定假期"):
            try:
                date_str = str(row.iloc[1]).replace('，', ',').replace(' ', '')
                d_list = [int(d) for d in date_str.split(',') if d.isdigit()]
                h_A = int(row.iloc[2]) if pd.notna(row.iloc[2]) else 2
                h_Bmin = int(row.iloc[3]) if pd.notna(row.iloc[3]) else 0
                h_Bmax = int(row.iloc[4]) if pd.notna(row.iloc[4]) else 14
                h_Cmin = int(row.iloc[5]) if pd.notna(row.iloc[5]) else 0
                h_Cmax = int(row.iloc[6]) if pd.notna(row.iloc[6]) else 14
                h_D = int(row.iloc[7]) if pd.notna(row.iloc[7]) else 1
                h_Z = int(row.iloc[8]) if len(row) > 8 and pd.notna(row.iloc[8]) else 0
                for d in d_list:
                    if d in days:
                        holiday_dict[d] = {
                            "A": h_A, "B_min": h_Bmin, "B_max": h_Bmax,
                            "C_min": h_Cmin, "C_max": h_Cmax, "D": h_D, "Z": h_Z
                        }
            except: pass
    emps = []
    emp_gender = {}
    for i in range(1, 15):
        emp_data = config.get(f"员工{i}")
        if emp_data and pd.notna(emp_data[0]) and str(emp_data[0]).strip() != "":
            name = str(emp_data[0]).strip()
            emps.append(name)
            if len(emp_data) > 1 and pd.notna(emp_data[1]):
                emp_gender[name] = str(emp_data[1]).strip()
    FULL_EXEMPT_EMPS = ["新1", "新2"]
    week_names = ["一", "二", "三", "四", "五", "六", "日"]
    date_labels = []
    for i in range(num_days):
        current_date = start_date + datetime.timedelta(days=i)
        label = f"{current_date.month}月{current_date.day}日(周{week_names[current_date.weekday()]})"
        date_labels.append(label)
    try:
        df_emp = safe_read_excel(input_path, sheet_index=resolve_sheet_index(input_path, "Sheet1", 0))
        df_pref = safe_read_excel(input_path, sheet_index=resolve_sheet_index(input_path, "Sheet2", 1))
        has_days_off = "常规休假天数" in df_emp.columns
        has_exempt = "豁免连休" in df_emp.columns
        has_exempt_count = "豁免次数" in df_emp.columns
        has_special_rest = "特殊指定休息日" in df_emp.columns
        has_banned = "固定不排班次" in df_emp.columns
        has_exempt_gap = "是否豁免班数极差" in df_emp.columns
        has_cw_max = "豁免连续上班天数" in df_emp.columns
        has_c_pre_restrict = "A班前一天是否愿意上C班" in df_emp.columns

    except Exception as e:
        print(f"\n❌ 读取排班条件失败！【原因】{e}")
        return False
    try:
        df_special = safe_read_excel(input_path, sheet_index=resolve_sheet_index(input_path, "Sheet3", 2))
        has_special = True
        for col in ["姓名", "日期", "班次", "要求"]:
            if col not in df_special.columns: has_special = False
    except:
        has_special = False
    hist_A_dict, hist_B_dict, hist_D_dict = {}, {}, {}
    special_rest_emp_set = set()
    for index, row in df_emp.iterrows():
        name = str(row["姓名"]).strip()
        if name not in emps:
            continue
        total_hist_A, total_hist_B, total_hist_D = 0, 0, 0
        for col_name in df_emp.columns:
            col_str = str(col_name).strip().upper()
            if col_str.endswith("_A"):
                if pd.notna(row[col_name]): total_hist_A += int(row[col_name])
            if col_str.endswith("_D"):
                if pd.notna(row[col_name]): total_hist_D += int(row[col_name])
            if col_str.endswith("_B"):
                if pd.notna(row[col_name]): total_hist_B += int(row[col_name])
        if "A班基础极差" in df_emp.columns and pd.notna(row["A班基础极差"]):
            total_hist_A += int(row["A班基础极差"])
        if "D班基础极差" in df_emp.columns and pd.notna(row["D班基础极差"]):
            total_hist_D += int(row["D班基础极差"])
        hist_A_dict[name] = total_hist_A
        hist_D_dict[name] = total_hist_D
        hist_B_dict[name] = total_hist_B
    force_work_set = set()
    allowed_z_days = set()
    if has_special:
        for _, srow in df_special.iterrows():
            if str(srow["要求"]).strip() == "必上":
                s_name = str(srow["姓名"]).strip()
                s_date_str = str(srow["日期"]).replace('，', ',').replace(' ', '')
                s_dates = [int(x) for x in s_date_str.split(',') if x.isdigit()]
                s_shift_str = str(srow["班次"]).replace('，', ',').replace(' ', '')
                s_shifts = [x.strip() for x in s_shift_str.split(',') if x.strip() in shift_types]
                for sd in s_dates:
                    for ss in s_shifts:
                        force_work_set.add((s_name, sd, ss))
                        if ss == "Z": allowed_z_days.add((s_name, sd))
    x = {}
    for e in emps:
        for d in days:
            for s in shift_types:
                x[(e, d, s)] = model.NewBoolVar(f"{e}_{d}_{s}")
    shift_times = {}
    for index, row in df_config.iterrows():
        shift_name = str(row.iloc[0]).strip()
        if shift_name in ["A班", "B班", "C班", "D班", "Z班"]:
            shift_key = shift_name[0]
            try:
                sh = float(row.iloc[1]) if pd.notna(row.iloc[1]) else 0.0
                sm = float(row.iloc[2]) if pd.notna(row.iloc[2]) else 0.0
                eh = float(row.iloc[3]) if pd.notna(row.iloc[3]) else 0.0
                em = float(row.iloc[4]) if pd.notna(row.iloc[4]) else 0.0
                shift_times[shift_key] = (sh + sm/60.0, eh + em/60.0)
            except: pass
    default_times = {"A": (8.0, 16.0), "B": (12.0, 20.0), "C": (14.0, 22.0), "D": (15.5, 23.5), "Z": (10.0, 22.0)}
    for k, v in default_times.items():
        if k not in shift_times: shift_times[k] = v

    def get_allowed_shifts(time_str):
        if pd.isna(time_str): return []
        time_str = str(time_str).strip().replace('—', '-').replace('－', '-').replace('~', '-')
        if "-" not in time_str: return []
        parts = time_str.split("-")
        if len(parts) != 2: return []
        def time_to_float(t_str):
            t_str = t_str.strip()
            try:
                if ":" in t_str:
                    h, m = t_str.split(":")
                    return float(h) + float(m) / 60.0
                return float(t_str)
            except: return 0.0
        target_start, target_end = time_to_float(parts[0]), time_to_float(parts[1])
        allowed = []
        for shift, (s_start, s_end) in shift_times.items():
            if s_start <= target_start and s_end >= target_end:
                allowed.append(shift)
        return allowed
    for e in emps:
        for d in days:
            model.Add(sum(x[(e, d, s)] for s in shift_types) == 1)
    for d in days:
        a_cnt = sum(x[(e, d, "A")] for e in emps)
        b_cnt = sum(x[(e, d, "B")] for e in emps)
        c_cnt = sum(x[(e, d, "C")] for e in emps)
        d_cnt = sum(x[(e, d, "D")] for e in emps)
        z_cnt = sum(x[(e, d, "Z")] for e in emps)
        off_cnt = sum(x[(e, d, "休")] for e in emps)
        if d in holiday_dict:
            hd = holiday_dict[d]
            model.Add(a_cnt == hd["A"])
            model.Add(b_cnt >= hd["B_min"])
            model.Add(b_cnt <= hd["B_max"])
            model.Add(c_cnt >= hd["C_min"])
            model.Add(c_cnt <= hd["C_max"])
            model.Add(d_cnt == hd["D"])
            model.Add(z_cnt == hd["Z"])
        else:
            model.Add(a_cnt >= A_min)
            model.Add(a_cnt <= A_max)
            model.Add(b_cnt >= B_min)
            model.Add(b_cnt <= B_max)
            model.Add(c_cnt >= C_min)
            model.Add(c_cnt <= C_max)
            model.Add(d_cnt >= D_min)
            model.Add(d_cnt <= D_max)
            model.Add(off_cnt >= OFF_min)
            model.Add(off_cnt <= OFF_max)
    for e in emps:
        for d in range(1, num_days):
            model.Add(x[(e, d, "D")] + x[(e, d+1, "A")] <= 1)
    emp_days_off_dict = {}
    exempt_count_dict = {}
    banned_shifts_dict = {}
    exempt_gap_emps = set()
    emp_penalty = {e: [] for e in emps}
    penalty = []
    for name, w_day in resignation_dict.items():
        if name in emps:
            exempt_gap_emps.add(name)
            for d in range(w_day + 1, num_days + 1):
                model.Add(x[(name, d, "休")] == 1)

    for index, row in df_emp.iterrows():
        name = str(row["姓名"]).strip()
        if name not in emps: continue
        has_special_rest = pd.notna(row["特殊指定休息日"])
        fixed_days = []
        week_map = {"周一":0, "周二":1, "周三":2, "周四":3, "周五":4, "周六":5, "周日":6}
        if has_special_rest:
            special_rest_emp_set.add(name)
            raw_text = str(row["特殊指定休息日"])
            raw_text = ''.join(c for c in raw_text if not unicodedata.category(c).startswith('C'))
            special_vals = raw_text.replace('，', ',').split(",")
            for sv in special_vals:
                sv = sv.strip()
                if not sv:
                    continue
                if sv.isdigit():
                    day_no = int(sv)
                    if day_no in days:
                        fixed_days.append(day_no)
                        model.Add(x[(name, day_no, "休")] == 1)
                elif "周" in sv and sv in week_map:
                    target_wd = week_map[sv]
                    for d in days:
                        cur_date = start_date + datetime.timedelta(days=d-1)
                        if cur_date.weekday() == target_wd:
                            model.Add(x[(name, d, "休")] == 1)

        if name in FULL_EXEMPT_EMPS:
            for d in days:
                model.Add(x[(name, d, "休")] == 1)
            continue


        cw_max_emp = consec_work_max
        if has_cw_max and pd.notna(row["豁免连续上班天数"]):
            try:
                cw_max_emp = int(float(row["豁免连续上班天数"]))
            except:
                cw_max_emp = consec_work_max

        # ========== 【全新】连续上班片段约束（只看本人实际上班，法定值班算上班，休息即打断） ==========
        is_work = {}
        for d in days:
            is_work[d] = model.NewBoolVar(f"is_work_{name}_{d}")
            model.Add(is_work[d] == 1 - x[(name, d, "休")])

        # ---------------- 上限硬约束：禁止连续上班超过 consec_work_max ----------------
        if cw_max_emp > 0:
            for d_start in range(1, num_days - cw_max_emp + 1):
                window_vars = [is_work[d_start + offset] for offset in range(cw_max_emp + 1)]
                model.Add(sum(window_vars) <= cw_max_emp)


        # ---------------- 追踪上班片段，用于下限硬约束+碎片软惩罚 ----------------
        start_work = {}
        for d in days:
            start_work[d] = model.NewBoolVar(f"start_work_{name}_{d}")
        model.Add(start_work[1] == is_work[1])
        for d in range(2, num_days+1):
            model.Add(start_work[d] >= is_work[d] - is_work[d-1])
            model.Add(start_work[d] <= is_work[d])
            model.Add(start_work[d] <= 1 - is_work[d-1])

        seg_len = {}
        for d in days:
            seg_len[d] = model.NewIntVar(0, num_days, f"seg_len_{name}_{d}")
        for d in days:
            expr_parts = []
            for offset in range(0, num_days - d +1):
                dd = d + offset
                b_tmp = model.NewBoolVar(f"tmp_{name}_{d}_{dd}")
                model.Add(b_tmp <= is_work[dd])
                if dd+1 <= num_days:
                    model.Add(b_tmp <= 1 - is_work[dd+1])
                else:
                    model.Add(b_tmp == is_work[dd])
                expr_parts.append((offset+1) * b_tmp)
            model.Add(seg_len[d] == sum(expr_parts))

        seg_need_hard_min = {}
        seg_short_penalty = []
        resign_end = resignation_dict.get(name, num_days)
        for d in days:
            seg_need_hard_min[d] = model.NewBoolVar(f"seg_hard_{name}_{d}")
            if consec_work_min <=0:
                model.Add(seg_need_hard_min[d]==0)
                continue

            cond_is_start = start_work[d]
            cond_not_month_start = 1 if (d>1 and (d-1) not in holiday_dict) else 0

            end_var = model.NewIntVar(d, num_days, f"seg_end_{name}_{d}")
            for offset in range(0, num_days - d +1):
                dd = d + offset
                b_end = model.NewBoolVar(f"bend_{name}_{d}_{dd}")
                model.Add(b_end <= is_work[dd])
                if dd+1 <= num_days:
                    model.Add(b_end <= 1 - is_work[dd+1])
                else:
                    model.Add(b_end == is_work[dd])
                model.Add(end_var >= dd - 1000*(1-b_end))
                model.Add(end_var <= dd + 1000*(1-b_end))

            # not_month_end = True → end_var < num_days (片段不在月末截断)
            not_month_end = model.NewBoolVar(f"nmend_{name}_{d}")
            M = 1000
            # 如果 not_month_end=1，则 end_var ≤ num_days‑1
            model.Add(end_var <= (num_days -1) + M * (1 - not_month_end))
            # 如果 not_month_end=0，则 end_var ≥ num_days
            model.Add(end_var >= num_days - M * not_month_end)

            # before_not_resign = True → end_var < resign_end (片段没有碰到离职日期)
            before_not_resign = model.NewBoolVar(f"bfrsig_{name}_{d}")
            # before_not_resign=1 → end_var ≤ resign_end‑1
            model.Add(end_var <= (resign_end -1) + M * (1 - before_not_resign))
            # before_not_resign=0 → end_var ≥ resign_end
            model.Add(end_var >= resign_end - M * before_not_resign)


            model.Add(seg_need_hard_min[d] <= cond_is_start)
            if cond_not_month_start:
             # 新增：多个条件同时满足，seg_need_hard_min[d]才强制等于1
                model.Add(seg_need_hard_min[d] >= cond_is_start + not_month_end + before_not_resign - 2)

                model.Add(seg_need_hard_min[d] <= 1)
            else:
                model.Add(seg_need_hard_min[d] <=0)
            model.Add(seg_need_hard_min[d] <= not_month_end)
            model.Add(seg_need_hard_min[d] <= before_not_resign)

            model.Add(seg_len[d] >= consec_work_min * seg_need_hard_min[d])

            is_fragment = model.NewBoolVar(f"isfrag_{name}_{d}")
            model.Add(is_fragment >= start_work[d] - seg_need_hard_min[d])
            model.Add(is_fragment <= start_work[d])
            model.Add(is_fragment <= 1 - seg_need_hard_min[d])

            delta = model.NewIntVar(0, consec_work_min, f"delta_{name}_{d}")
            model.AddMaxEquality(delta, [0, consec_work_min - seg_len[d]])
            pen = model.NewIntVar(0, 5000, f"fragpen_{name}_{d}")
            model.Add(pen >= delta * 300 - 5000*(1-is_fragment))
            model.Add(pen <= delta * 300)
            seg_short_penalty.append(pen)
        emp_penalty[name].extend(seg_short_penalty)

        if has_exempt_gap and str(row["是否豁免班数极差"]).strip() == "是":
            exempt_gap_emps.add(name)
        c_pre_ok = False
        if has_c_pre_restrict and pd.notna(row["A班前一天是否愿意上C班"]):
            c_pre_ok = (str(row["A班前一天是否愿意上C班"]).strip() == "是")
        if not c_pre_ok:
            for d in range(2, num_days + 1):
                v_pre_c = model.NewBoolVar(f"v_pre_c_{name}_{d}")
                model.Add(v_pre_c >= x[(name, d, "A")] + x[(name, d-1, "C")] - 1)
                penalty.append(v_pre_c * 500)
        days_off = 0
        if has_days_off and pd.notna(row["常规休假天数"]):
            try: days_off += int(row["常规休假天数"])
            except: pass
        has_holiday_inc = "法定假期增加天数" in df_emp.columns
        if has_holiday_inc and pd.notna(row["法定假期增加天数"]):
            try: days_off += int(row["法定假期增加天数"])
            except: pass
        if name in resignation_dict:
            w_day = resignation_dict[name]
            days_off += (num_days - w_day)
        exempt_cnt = 0
        if has_exempt and str(row["豁免连休"]).strip() == "是":
            exempt_cnt = 1
            if has_exempt_count and pd.notna(row["豁免次数"]):
                try:
                    exempt_cnt = int(row["豁免次数"])
                except:
                    pass
        exempt_count_dict[name] = exempt_cnt
        if days_off > 0:
            emp_days_off_dict[name] = days_off
            model.Add(sum(x[(name, d, "休")] for d in days) == days_off)
            exempt_consec = (str(row["豁免连休"]).strip() == "是")
            allow_consec_count = exempt_count_dict[name]
            if not exempt_consec:
                for d in range(1, num_days):
                    d1 = d
                    d2 = d + 1
                    if (d1 in fixed_days) and (d2 in fixed_days):
                        continue
                    if (d1 in holiday_dict) or (d2 in holiday_dict):
                        continue
                    model.Add(x[(name, d1, "休")] + x[(name, d2, "休")] <= 1)
            else:
                for d in range(1, num_days -1):
                    d1=d
                    d2=d+1
                    d3=d+2
                    if (d1 in fixed_days) and (d2 in fixed_days) and (d3 in fixed_days):
                        continue
                    if (d1 in holiday_dict) or (d2 in holiday_dict) or (d3 in holiday_dict):
                        continue
                    model.Add(x[(name,d,"休")] + x[(name,d+1,"休")] + x[(name,d+2,"休")] <=2)
                consec_pair_vars = []
                for d in range(1, num_days):
                    if (d in holiday_dict) or (d+1 in holiday_dict):
                        continue
                    v_pair = model.NewBoolVar(f"v_consec_pair_{name}_{d}")
                    model.Add(v_pair >= x[(name,d,"休")] + x[(name,d+1,"休")] -1)
                    model.Add(v_pair <= x[(name,d,"休")])
                    model.Add(v_pair <= x[(name,d+1,"休")])
                    consec_pair_vars.append(v_pair)
                total_consec_pairs = model.NewIntVar(0, num_days, f"total_consec_pairs_{name}")
                model.Add(total_consec_pairs == sum(consec_pair_vars))
                excess_consec = model.NewIntVar(0, num_days, f"excess_consec_{name}")
                model.AddMaxEquality(excess_consec, [0, total_consec_pairs - allow_consec_count])
                emp_penalty[name].append(excess_consec * 800)
        banned_shift = ""
        if has_banned and pd.notna(row["固定不排班次"]):
            banned_shift = str(row["固定不排班次"]).strip()
            if banned_shift in shift_types:
                for d in days:
                    if (name, d, banned_shift) not in force_work_set:
                        model.Add(x[(name, d, banned_shift)] == 0)
        banned_shifts_dict[name] = banned_shift

    for name, w_day in resignation_dict.items():
        if name not in emps or name in FULL_EXEMPT_EMPS:
            continue
        total_staff = len([e for e in emps if e not in FULL_EXEMPT_EMPS])
        if total_staff <= 0 or w_day < 2:
            continue
        target_A = round(A_min * w_day / total_staff)
        target_D = round(D_min * w_day / total_staff)
        banned = banned_shifts_dict.get(name, "")
        if "A" not in banned:
            sum_A = sum(x[(name, d, "A")] for d in range(1, w_day + 1))
            lo_A = max(1, target_A - 1) if target_A >= 1 else 0
            hi_A = max(1, target_A + 1)
            model.Add(sum_A >= lo_A)
            model.Add(sum_A <= hi_A)
            dev_A = model.NewIntVar(0, w_day, f"dev_A_resign_{name}")
            model.AddAbsEquality(dev_A, sum_A - target_A)
            emp_penalty[name].append(dev_A * 500)
        if "D" not in banned:
            sum_D = sum(x[(name, d, "D")] for d in range(1, w_day + 1))
            lo_D = max(1, target_D - 1) if (target_D >= 1 and w_day >= 3) else 0
            hi_D = max(1, target_D + 1)
            model.Add(sum_D >= lo_D)
            model.Add(sum_D <= hi_D)
            dev_D = model.NewIntVar(0, w_day, f"dev_D_resign_{name}")
            model.AddAbsEquality(dev_D, sum_D - target_D)
            emp_penalty[name].append(dev_D * 500)
        if "A" not in banned and "D" not in banned and target_A + target_D >= 1:
            sum_AD = sum(x[(name, d, "A")] for d in range(1, w_day + 1)) + sum(x[(name, d, "D")] for d in range(1, w_day + 1))
            target_AD = target_A + target_D
            model.Add(sum_AD >= max(1, target_AD - 1))
            model.Add(sum_AD <= target_AD + 1)
    if has_special:
        for index, row in df_special.iterrows():
            name = str(row["姓名"]).strip()
            if name not in emps or name in FULL_EXEMPT_EMPS: continue
            date_val = str(row["日期"]).strip()
            shift_str = str(row["班次"]).replace('，', ',').replace(' ', '')
            shift_list = [s for s in shift_str.split(',') if s in shift_types]
            req_type = str(row["要求"]).strip()
            if not shift_list: continue
            if date_val == "休前":
                for d in range(1, num_days):
                    for target_s in shift_list:
                        if req_type == "必上":
                            model.Add(sum(x[(name, d, ts)] for ts in shift_list)
                                      >= 1 - 1000 * (1 - x[(name, d+1, "休")] + x[(name, d, "休")]))
                        elif req_type == "不上":
                            model.Add(x[(name, d+1, "休")] + x[(name, d, target_s)] <= 1)
                        elif req_type == "尽量不上":

                            v = model.NewBoolVar(f"v_br_ns_{name}_{d}_{target_s}")
                            model.Add(v <= x[(name, d+1, "休")])
                            model.Add(v <= x[(name, d, target_s)])
                            model.Add(v >= x[(name, d+1, "休")] + x[(name, d, target_s)] - 1)
                            emp_penalty[name].append(v * 150)
                        elif req_type == "尽量上":
                            v = model.NewBoolVar(f"v_br_ys_{name}_{d}_{target_s}")
                            model.Add(v <= x[(name, d+1, "休")])
                            model.Add(v <= 1 - x[(name, d, target_s)])
                            model.Add(v >= x[(name, d+1, "休")] - x[(name, d, target_s)])
                            emp_penalty[name].append(v * 150)
            elif date_val == "休后":
                for d in range(2, num_days + 1):
                    for target_s in shift_list:
                        if req_type == "必上":
                            model.Add(sum(x[(name, d, ts)] for ts in shift_list)
                                      >= 1 - 1000 * (1 - x[(name, d-1, "休")] + x[(name, d, "休")]))
                        elif req_type == "不上":
                            model.Add(x[(name, d-1, "休")] + x[(name, d, target_s)] <= 1)
                        elif req_type == "尽量不上":

                            v = model.NewBoolVar(f"v_ar_ns_{name}_{d}_{target_s}")
                            model.Add(v <= x[(name, d-1, "休")])
                            model.Add(v <= x[(name, d, target_s)])
                            model.Add(v >= x[(name, d-1, "休")] + x[(name, d, target_s)] - 1)
                            emp_penalty[name].append(v * 150)
                        elif req_type == "尽量上":
                            v = model.NewBoolVar(f"v_ar_ys_{name}_{d}_{target_s}")
                            model.Add(v <= x[(name, d-1, "休")])
                            model.Add(v <= 1 - x[(name, d, target_s)])
                            model.Add(v >= x[(name, d-1, "休")] - x[(name, d, target_s)])
                            emp_penalty[name].append(v * 150)
            else:
                try:
                    date_str = date_val.replace('，', ',').replace(' ', '')
                    date_list = [int(d) for d in date_str.split(',') if d.isdigit()]
                    for target_d in date_list:
                        if target_d not in days:
                            continue
                        if req_type == "必上":
                            model.Add(sum(x[(name, target_d, target_s)] for target_s in shift_list) >= 1)
                            if "Z" in shift_list:
                                allowed_z_days.add((name, target_d))
                        elif req_type == "不上":
                            for target_s in shift_list:
                                model.Add(x[(name, target_d, target_s)] == 0)
                        elif req_type == "尽量上":
                            for target_s in shift_list:
                                emp_penalty[name].append((1 - x[(name, target_d, target_s)]) * 150)
                        elif req_type == "尽量不上":
                            for target_s in shift_list:
                                emp_penalty[name].append(x[(name, target_d, target_s)] * 150)
                except: pass
    has_date_col = "日期" in df_pref.columns
    week_map = {"周一":0, "周二":1, "周三":2, "周四":3, "周五":4, "周六":5, "周日":6}
    v_rest_mid = {e: [] for e in emps if e not in FULL_EXEMPT_EMPS}
    v_work_mid = {e: [] for e in emps if e not in FULL_EXEMPT_EMPS}
    v_rest_high = {e: [] for e in emps if e not in FULL_EXEMPT_EMPS}
    v_work_high = {e: [] for e in emps if e not in FULL_EXEMPT_EMPS}
    for index, row in df_pref.iterrows():
        name = str(row["姓名"]).strip()
        if name not in emps or name in FULL_EXEMPT_EMPS: continue
        if pd.isna(row["时间段"]): continue
        wd_str = str(row["星期几"]) if "星期几" in row and pd.notna(row["星期几"]) else ""
        wd_list = [wd.strip() for wd in wd_str.replace('，', ',').split(',') if wd.strip()]
        target_wds = [week_map[wd] for wd in wd_list if wd in week_map]
        date_val = str(row["日期"]) if has_date_col and pd.notna(row["日期"]) else ""
        date_list = [int(d) for d in date_val.replace('，', ',').replace(' ', '').split(',') if d.isdigit()]
        if not target_wds and not date_list: continue
        time_desc = str(row["时间段"])
        time_segments = [t.strip() for t in time_desc.replace('，', ',').split(',') if t.strip()]
        level = str(row.get("重要程度", "一般可协调")).strip()
        if "必" in level: level_type = "必上"
        elif "尽" in level: level_type = "尽量安排"
        else: level_type = "一般可协调"
        for d in days:
            if d in holiday_dict: continue
            current_date = start_date + datetime.timedelta(days=d-1)
            if d in date_list or current_date.weekday() in target_wds:
                for seg in time_segments:
                    allowed_shifts = get_allowed_shifts(seg)
                    if not allowed_shifts: continue
                    if level_type == "必上":
                        emp_penalty[name].append((1 - sum(x[(name, d, s)] for s in allowed_shifts)) * 400)
                        continue
                    v_expr = model.NewBoolVar(f"v_{name}_{d}_{seg}")
                    model.Add(v_expr == 1 - sum(x[(name, d, s)] for s in allowed_shifts))
                    v_r = model.NewBoolVar(f"vr_{name}_{d}_{seg}")
                    model.Add(v_r <= v_expr)
                    model.Add(v_r <= x[(name, d, "休")])
                    model.Add(v_r >= v_expr + x[(name, d, "休")] - 1)
                    v_w = model.NewBoolVar(f"vw_{name}_{d}_{seg}")
                    model.Add(v_w <= v_expr)
                    model.Add(v_w <= 1 - x[(name, d, "休")])
                    model.Add(v_w >= v_expr - x[(name, d, "休")])
                    if level_type == "尽量安排":
                        v_rest_high[name].append(v_r)
                        v_work_high[name].append(v_w)
                    else:
                        v_rest_mid[name].append(v_r)
                        v_work_mid[name].append(v_w)
    for e in emps:
        if e in FULL_EXEMPT_EMPS: continue
        for d in days:
            if (e, d) not in allowed_z_days:
                model.Add(x[(e, d, "Z")] == 0)
    score_Z = int(get_cfg_val("Z班惩罚", 200, 0))
    score_continuous_rest = int(get_cfg_val("连休惩罚", 100, 0))
    score_evening_short = int(get_cfg_val("晚班缺人惩罚", 1, 0))
    score_short_gap = int(get_cfg_val("休息过短惩罚", 220, 0))
    raw_rest_mid = parse_list(config.get("一般排休分")[0], [0, 0, 2, 3, 5, 7, 10, 13, 16, 19, 22, 25, 28, 31, 34, 37, 40, 43, 46, 49, 52, 55])
    raw_work_mid = parse_list(config.get("一般排班分")[0], [0, 2, 8, 13, 18, 23, 28, 33, 38, 43, 48, 53, 58, 63, 68, 73, 78, 83, 88, 93, 98, 103])
    raw_rest_high = parse_list(config.get("尽量排休分")[0], [i * 10 for i in range(30)])
    raw_work_high = parse_list(config.get("尽量排班分")[0], [0, 5, 20, 50, 100, 175, 280, 420, 600, 825, 1100, 1425, 1800, 2225, 2700, 3225, 3800, 4425, 5100, 5825, 6600, 7425])
    for e in emps:
        if e in FULL_EXEMPT_EMPS: continue
        for d in days:
            if d in holiday_dict: continue
            emp_penalty[e].append(x[(e, d, "Z")] * score_Z)
    for e in emps:
        if e in FULL_EXEMPT_EMPS: continue
        if e in special_rest_emp_set: continue
        emp_end_day = resignation_dict.get(e, num_days)
        for d in range(num_days - 4):
            window_days = list(range(d+1, d+6))
            if any(day > emp_end_day for day in window_days):
                continue
            rest_sum = model.NewIntVar(0, 5, f"rest_sum_{e}_{d}")
            model.Add(rest_sum == sum(x[(e, day, "休")] for day in window_days))
            pen_var = model.NewIntVar(0, 5000, f"pen_gap_{e}_{d}")
            model.AddElement(rest_sum, [0, 0, score_short_gap, score_short_gap*2, score_short_gap*3, score_short_gap*4], pen_var)
            emp_penalty[e].append(pen_var)
    for d in days:
        if d in holiday_dict: continue
        b_count = model.NewIntVar(0, len(emps), f"b_count_{d}")
        model.Add(b_count == sum(x[(e, d, "B")] for e in emps))
        c_count = model.NewIntVar(0, len(emps), f"c_count_{d}")
        model.Add(c_count == sum(x[(e, d, "C")] for e in emps))
        c18 = model.NewIntVar(0, len(emps), f"c18_{d}")
        model.Add(c18 == sum(x[(e, d, "B")] + x[(e, d, "C")] + x[(e, d, "D")] + x[(e, d, "Z")] for e in emps))
        s18 = model.NewIntVar(0, len(emps), f"s18_{d}")
        model.AddMaxEquality(s18, [0, 6 - c18])
        penalty.append(s18 * score_evening_short)
        penalty.append(s18 * score_evening_short)
        c20 = model.NewIntVar(0, len(emps), f"c20_{d}")
        model.Add(c20 == sum(x[(e, d, "C")] + x[(e, d, "D")] + x[(e, d, "Z")] for e in emps))
        s20 = model.NewIntVar(0, len(emps), f"s20_{d}")
        model.AddMaxEquality(s20, [0, 5 - c20])
        penalty.append(s20 * score_evening_short)
        penalty.append(s20 * score_evening_short)
    def add_penalty_pool(pool_dict, pen_list, var_prefix):
        for e in emps:
            if e in FULL_EXEMPT_EMPS: continue
            if pool_dict[e]:
                V = model.NewIntVar(0, len(pool_dict[e]), f"V_{var_prefix}_{e}")
                model.Add(V == sum(pool_dict[e]))
                p = model.NewIntVar(0, 500000, f"p_{var_prefix}_{e}")
                safe_len = len(pen_list)
                while len(pen_list) <= len(pool_dict[e]):
                    pen_list.append(pen_list[-1] if safe_len > 0 else 0)
                prefix_pen_list = []
                curr = 0
                for x in pen_list:
                    curr += x
                    prefix_pen_list.append(curr)
                model.AddElement(V, prefix_pen_list, p)
                emp_penalty[e].append(p)
    add_penalty_pool(v_rest_mid, raw_rest_mid, "rm")
    add_penalty_pool(v_work_mid, raw_work_mid, "wm")
    add_penalty_pool(v_rest_high, raw_rest_high, "rh")
    add_penalty_pool(v_work_high, raw_work_high, "wh")
    pool_A = [e for e in emps if e not in FULL_EXEMPT_EMPS and e not in exempt_gap_emps and "A" not in banned_shifts_dict.get(e, "")]
    pool_D = [e for e in emps if e not in FULL_EXEMPT_EMPS and e not in exempt_gap_emps and "D" not in banned_shifts_dict.get(e, "")]
    pool_AD = [e for e in emps if e not in FULL_EXEMPT_EMPS and e not in exempt_gap_emps and "A" not in banned_shifts_dict.get(e, "") and "D" not in banned_shifts_dict.get(e, "")]
    pool_B = [e for e in emps if e not in FULL_EXEMPT_EMPS and "B" not in banned_shifts_dict.get(e, "")]
    if len(pool_A) > 1:
        total_A = {e: model.NewIntVar(0, num_days * 36, f"totalA_{e}") for e in pool_A}
        for e in pool_A: model.Add(total_A[e] == sum(x[(e, d, "A")] for d in days) + hist_A_dict.get(e, 0))
        max_A, min_A = model.NewIntVar(0, num_days * 36, "max_A"), model.NewIntVar(0, num_days * 36, "min_A")
        model.AddMaxEquality(max_A, [total_A[e] for e in pool_A])
        model.AddMinEquality(min_A, [total_A[e] for e in pool_A])
        model.Add(max_A - min_A <= A_gap_max)
    if len(pool_D) > 1:
        total_D = {e: model.NewIntVar(0, num_days * 36, f"totalD_{e}") for e in pool_D}
        for e in pool_D: model.Add(total_D[e] == sum(x[(e, d, "D")] for d in days) + hist_D_dict.get(e, 0))
        max_D, min_D = model.NewIntVar(0, num_days * 36, "max_D"), model.NewIntVar(0, num_days * 36, "min_D")
        model.AddMaxEquality(max_D, [total_D[e] for e in pool_D])
        model.AddMinEquality(min_D, [total_D[e] for e in pool_D])
        model.Add(max_D - min_D <= D_gap_max)
    if len(pool_AD) > 1:
        total_AD = {e: model.NewIntVar(0, num_days * 72, f"totalAD_{e}") for e in pool_AD}
        for e in pool_AD: model.Add(total_AD[e] == sum(x[(e, d, "A")] for d in days) + sum(x[(e, d, "D")] for d in days) + hist_A_dict.get(e, 0) + hist_D_dict.get(e, 0))
        max_AD, min_AD = model.NewIntVar(0, num_days * 72, "max_AD"), model.NewIntVar(0, num_days * 72, "min_AD")
        model.AddMaxEquality(max_AD, [total_AD[e] for e in pool_AD])
        model.AddMinEquality(min_AD, [total_AD[e] for e in pool_AD])
        model.Add(max_AD - min_AD <= AD_gap_max)
    if len(pool_B) > 1:
        total_B = {e: model.NewIntVar(0, num_days * 36, f"totalB_{e}") for e in pool_B}
        for e in pool_B: model.Add(total_B[e] == sum(x[(e, d, "B")] for d in days) + hist_B_dict.get(e, 0))
        max_B, min_B = model.NewIntVar(0, num_days * 36, "max_B"), model.NewIntVar(0, num_days * 36, "min_B")
        model.AddMaxEquality(max_B, [total_B[e] for e in pool_B])
        model.AddMinEquality(min_B, [total_B[e] for e in pool_B])
        model.Add(max_B - min_B <= B_gap_max)
    females = [e for e in emps if emp_gender.get(e) == "女" and e not in FULL_EXEMPT_EMPS]
    males = [e for e in emps if emp_gender.get(e) == "男" and e not in FULL_EXEMPT_EMPS]
    if force_female_less_D and len(females) > 0 and len(males) > 0:
        count_D_female = {e: model.NewIntVar(0, num_days, f"cntD_f_{e}") for e in females}
        for e in females: model.Add(count_D_female[e] == sum(x[(e, d, "D")] for d in days) + hist_D_dict.get(e, 0))
        max_female_D = model.NewIntVar(0, num_days * 36, "max_female_D")
        model.AddMaxEquality(max_female_D, [count_D_female[e] for e in females])
        count_D_male = {e: model.NewIntVar(0, num_days, f"cntD_m_{e}") for e in males}
        for e in males: model.Add(count_D_male[e] == sum(x[(e, d, "D")] for d in days) + hist_D_dict.get(e, 0))
        min_male_D = model.NewIntVar(0, num_days * 36, "min_male_D")
        model.AddMinEquality(min_male_D, [count_D_male[e] for e in males])
        model.Add(max_female_D <= min_male_D)
    for e in emps:
        if emp_penalty[e]:
            total_p = sum(emp_penalty[e])
            if max_emp_penalty > 0:
                model.Add(total_p <= max_emp_penalty)
            penalty.append(total_p)
    model.Minimize(sum(penalty))
    print(f"正在计算 {YEAR}年{MONTH}月 排班，请稍候...")
    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 60.0
    status = solver.Solve(model)


    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        print(f"✅ 排班成功！正在分析惩罚原因并打包 Excel...")
        schedule = {e: {} for e in emps}
        for e in emps:
            for d in days:
                for s in shift_types:
                    if solver.Value(x[(e, d, s)]) == 1:
                        schedule[e][d] = s
        df_emp_out = pd.DataFrame(index=emps, columns=days)
        for e in emps:
            for d in days: df_emp_out.at[e, d] = schedule[e][d]
        df_emp_out.columns = date_labels
        stats_list = []
        for shift, name in [("A", "A班合计"), ("B", "B班合计"), ("C", "C班合计"), ("D", "D班合计"), ("Z", "Z班合计"), ("休", "休息合计")]:
            row = {col: sum(1 for e in emps if schedule[e][d] == shift) for d, col in zip(days, date_labels)}
            row["员工"] = name
            stats_list.append(row)
        df_stats = pd.DataFrame(stats_list).set_index("员工")
        blank_row = pd.DataFrame([{col: "" for col in date_labels}], index=[" "])
        stats_header = {"员工": "天数统计"}
        headers_list = ["A班天数", "B班天数", "C班天数", "D班天数", "Z班天数", "休息天数"]
        for i, col_name in enumerate(date_labels):
            stats_header[col_name] = headers_list[i] if i < 6 else ""
        if len(date_labels) > 9:
            stats_header[date_labels[7]] = "累计A班总数"
            stats_header[date_labels[8]] = "累计B班总数"
            stats_header[date_labels[9]] = "累计D班总数"
        if len(date_labels) > 11:
            stats_header[date_labels[11]] = "休息日期"
        header_row = pd.DataFrame([stats_header], index=["说明"])
        emp_totals = []
        summary_list = []
        for e in emps:
            a_cnt = sum(1 for d in days if schedule[e][d] == "A")
            b_cnt = sum(1 for d in days if schedule[e][d] == "B")
            c_cnt = sum(1 for d in days if schedule[e][d] == "C")
            d_cnt = sum(1 for d in days if schedule[e][d] == "D")
            z_cnt = sum(1 for d in days if schedule[e][d] == "Z")
            off_cnt = sum(1 for d in days if schedule[e][d] == "休")
            off_day_list = [str(d) for d in days if schedule[e][d] == "休"]
            off_day_str = ",".join(off_day_list)
            row_vals = [a_cnt, b_cnt, c_cnt, d_cnt, z_cnt, off_cnt] + [""] * (num_days - 6)
            row = {"员工": e}
            for i, col_name in enumerate(date_labels): row[col_name] = row_vals[i]
            if len(date_labels) > 9:
                row[date_labels[7]] = a_cnt + hist_A_dict.get(e, 0)
                row[date_labels[8]] = b_cnt + hist_B_dict.get(e, 0)
                row[date_labels[9]] = d_cnt + hist_D_dict.get(e, 0)
            if len(date_labels) > 11:
                row[date_labels[11]] = off_day_str
            emp_totals.append(row)
            summary_list.append({"姓名": e, "A班次数": a_cnt, "B班次数": b_cnt, "D班次数": d_cnt})
        df_totals = pd.DataFrame(emp_totals).set_index("员工")
        df_summary = pd.DataFrame(summary_list).set_index("姓名")
        df_final_emp = pd.concat([df_emp_out, df_stats, blank_row, header_row, df_totals])
        df_final_emp.index.name = "员工"
        df_day = pd.DataFrame(index=days, columns=["A班", "B班", "C班", "D班", "Z班", "休息"])
        for d in days:
            for e in emps:
                s = schedule[e][d]
                col = "休息" if s == "休" else f"{s}班"
                if pd.isna(df_day.at[d, col]): df_day.at[d, col] = e
                else: df_day.at[d, col] = f"{df_day.at[d, col]}, {e}"
        df_day.index = date_labels
        cal_data = [["" for _ in range(8)] for _ in range(36)]
        cal_data[0] = ["班次", "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"]
        start_wd = (start_date.weekday() + 1) % 7
        for d in days:
            current_date = start_date + datetime.timedelta(days=d-1)
            col_idx = (current_date.weekday() + 2) % 7
            if col_idx == 0:
                col_idx = 7
            week_idx = (d + start_wd - 1) // 7
            base_row = 1 + week_idx * 6
            a_list = [e for e in emps if schedule[e][d] == "A"]
            b_list = [e for e in emps if schedule[e][d] == "B"]
            d_list = [e for e in emps if schedule[e][d] == "D"]
            off_list = [e for e in emps if schedule[e][d] == "休"]
            cal_data[base_row][col_idx] = f"{MONTH}月{d}日"
            cal_data[base_row + 1][col_idx] = ", ".join(a_list)
            cal_data[base_row + 2][col_idx] = ", ".join(b_list)
            cal_data[base_row + 3][col_idx] = ", ".join(d_list)
            cal_data[base_row + 4][col_idx] = ", ".join(off_list)
        for w in range(6):
            base_row = 1 + w * 6
            cal_data[base_row][0] = "日期"
            cal_data[base_row + 1][0] = a_shift_name
            cal_data[base_row + 2][0] = b_shift_name
            cal_data[base_row + 3][0] = d_shift_name
            cal_data[base_row + 4][0] = "休息"
        df_calendar = pd.DataFrame(cal_data)
        penalty_records = []
        total_penalty = 0
        cnt_rm = {e: 0 for e in emps}
        cnt_wm = {e: 0 for e in emps}
        cnt_rh = {e: 0 for e in emps}
        cnt_wh = {e: 0 for e in emps}
        if has_special:
            for index, row in df_special.iterrows():
                name = str(row["姓名"]).strip()
                if name not in emps or name in FULL_EXEMPT_EMPS: continue
                date_val = str(row["日期"]).strip()
                shift_str = str(row["班次"]).replace('，', ',').replace(' ', '')
                shift_list = [s for s in shift_str.split(',') if s in shift_types]
                req_type = str(row["要求"]).strip()
                if not shift_list: continue
                if date_val == "休前":
                    for d in range(1, num_days):
                        if d in holiday_dict or d+1 in holiday_dict: continue
                        if schedule[name][d+1] == "休":
                            for target_s in shift_list:
                                if req_type == "尽量不上" and schedule[name][d] == target_s:
                                    total_penalty += 150
                                    penalty_records.append({"员工": name, "日期": date_labels[d-1], "原因": f"休前尽量不上{target_s}班", "扣分": 150, "说明": "软约束违规"})
                                elif req_type == "尽量上" and schedule[name][d] != target_s:
                                    total_penalty += 150
                                    penalty_records.append({"员工": name, "日期": date_labels[d-1], "原因": f"休前尽量上{target_s}班", "扣分": 150, "说明": "软约束违规"})
                elif date_val == "休后":
                    for d in range(2, num_days + 1):
                        if d in holiday_dict or d-1 in holiday_dict: continue
                        if schedule[name][d-1] == "休":
                            for target_s in shift_list:
                                if req_type == "尽量不上" and schedule[name][d] == target_s:
                                    total_penalty += 150
                                    penalty_records.append({"员工": name, "日期": date_labels[d-1], "原因": f"休后尽量不上{target_s}班", "扣分": 150, "说明": "软约束违规"})
                                elif req_type == "尽量上" and schedule[name][d] != target_s:
                                    total_penalty += 150
                                    penalty_records.append({"员工": name, "日期": date_labels[d-1], "原因": f"休后尽量上{target_s}班", "扣分": 150, "说明": "软约束违规"})
                else:
                    try:
                        date_str = date_val.replace('，', ',').replace(' ', '')
                        date_list = [int(d) for d in date_str.split(',') if d.isdigit()]
                        for target_d in date_list:
                            if target_d in days and target_d not in holiday_dict:
                                for target_s in shift_list:
                                    if req_type == "尽量上" and schedule[name][target_d] != target_s:
                                        total_penalty += 150
                                        penalty_records.append({"员工": name, "日期": date_labels[target_d-1], "原因": f"尽量上{target_s}班", "扣分": 150, "说明": "软约束违规"})
                                    elif req_type == "尽量不上" and schedule[name][target_d] == target_s:
                                        total_penalty += 150
                                        penalty_records.append({"员工": name, "日期": date_labels[target_d-1], "原因": f"尽量不上{target_s}班", "扣分": 150, "说明": "软约束违规"})
                    except: pass
        for e in emps:
            if e in FULL_EXEMPT_EMPS: continue
            emp_end_day = resignation_dict.get(e, num_days)
            row_e = df_emp[df_emp["姓名"] == e]
            c_pre_ok = False
            if not row_e.empty and has_c_pre_restrict and pd.notna(row_e.iloc[0]["A班前一天是否愿意上C班"]):
                c_pre_ok = (str(row_e.iloc[0]["A班前一天是否愿意上C班"]).strip() == "是")
            if not c_pre_ok:
                for d in range(2, num_days + 1):
                    if d in holiday_dict or (d-1) in holiday_dict: continue
                    if schedule[e][d] == "A" and schedule[e][d-1] == "C":
                        total_penalty += 500
                        penalty_records.append({"员工": e, "日期": f"{date_labels[d-2]} 至 {date_labels[d-1]}", "原因": "A班前一天排了C班", "扣分": 500, "说明": "软约束违规"})
            for d in days:
                if d in holiday_dict: continue
                current_date = start_date + datetime.timedelta(days=d-1)
                wd = current_date.weekday()
                s = schedule[e][d]
                if s == "Z":
                    total_penalty += score_Z
                    z_reason = "为补足晚班覆盖人数"
                    emp_prefs = df_pref[df_pref["姓名"] == e]
                    for _, pref_row in emp_prefs.iterrows():
                        wd_str = str(pref_row["星期几"]) if "星期几" in pref_row and pd.notna(pref_row["星期几"]) else ""
                        wd_list = [wd_t.strip() for wd_t in wd_str.replace('，', ',').split(',') if wd_t.strip()]
                        target_wds = [week_map[wd_t] for wd_t in wd_list if wd_t in week_map]
                        date_val = str(pref_row["日期"]) if has_date_col and pd.notna(pref_row["日期"]) else ""
                        date_list = [int(d_t) for d_t in date_val.replace('，', ',').replace(' ', '').split(',') if d_t.isdigit()]
                        if d in date_list or wd in target_wds:
                            z_reason = f"为满足个人在岗要求"
                            break
                    penalty_records.append({"员工": e, "日期": date_labels[d-1], "原因": "排了特殊班(Z)", "扣分": score_Z, "说明": z_reason})
                emp_prefs = df_pref[df_pref["姓名"] == e]
                for _, pref_row in emp_prefs.iterrows():
                    wd_str = str(pref_row["星期几"]) if "星期几" in pref_row and pd.notna(pref_row["星期几"]) else ""
                    wd_list = [wd_t.strip() for wd_t in wd_str.replace('，', ',').split(',') if wd_t.strip()]
                    target_wds = [week_map[wd_t] for wd_t in wd_list if wd_t in week_map]
                    date_val = str(pref_row["日期"]) if has_date_col and pd.notna(pref_row["日期"]) else ""
                    date_list = [int(d_t) for d_t in date_val.replace('，', ',').replace(' ', '').split(',') if d_t.isdigit()]
                    if d in date_list or wd in target_wds:
                        time_segments = [t.strip() for t in str(pref_row["时间段"]).replace('，', ',').split(',') if t.strip()]
                        level = str(pref_row.get("重要程度", "一般可协调")).strip()
                        if "必" in level: level_type = "必上"
                        elif "尽" in level: level_type = "尽量安排"
                        else: level_type = "一般可协调"
                        for seg in time_segments:
                            allowed = get_allowed_shifts(seg)
                            if not allowed: continue
                            if s not in allowed:
                                if level_type == "必上":
                                    total_penalty += 400
                                    penalty_records.append({"员工": e, "日期": date_labels[d-1], "原因": f"未满足必上要求({seg})", "扣分": 400})

                                    total_penalty += 400
                                    penalty_records.append({"员工": e, "日期": date_labels[d-1], "原因": f"未满足必上要求({seg})", "扣分": 400, "说明": "极高惩罚分"})
                                    continue
                                if s == "休":
                                    if level_type == "尽量安排":
                                        cnt_rh[e] +=1
                                        idx = min(cnt_rh[e], len(raw_rest_high)-1)
                                        score = raw_rest_high[idx]
                                        total_penalty += score
                                        penalty_records.append({"员工": e, "日期": date_labels[d-1], "原因": f"未满足在岗时段({seg})", "扣分": score, "说明": f"尽量安排排班为休(第{cnt_rh[e]}次)"})
                                    else:
                                        cnt_rm[e] +=1
                                        idx = min(cnt_rm[e], len(raw_rest_mid)-1)
                                        score = raw_rest_mid[idx]
                                        total_penalty += score
                                        penalty_records.append({"员工": e, "日期": date_labels[d-1], "原因": f"未满足在岗时段({seg})", "扣分": score, "说明": f"一般可协调排班为休(第{cnt_rm[e]}次)"})
                                else:
                                    if level_type == "尽量安排":
                                        cnt_wh[e] +=1
                                        idx = min(cnt_wh[e], len(raw_work_high)-1)
                                        score = raw_work_high[idx]
                                        total_penalty += score
                                        penalty_records.append({"员工": e, "日期": date_labels[d-1], "原因": f"未满足在岗时段({seg})", "扣分": score, "说明": f"尽量安排排班为{s}(第{cnt_wh[e]}次)"})
                                    else:
                                        cnt_wm[e] +=1
                                        idx = min(cnt_wm[e], len(raw_work_mid)-1)
                                        score = raw_work_mid[idx]
                                        total_penalty += score
                                        penalty_records.append({"员工": e, "日期": date_labels[d-1], "原因": f"未满足在岗时段({seg})", "扣分": score, "说明": f"一般可协调排班为{s}(第{cnt_wm[e]}次)"})
        for e in emps:
            if e in FULL_EXEMPT_EMPS: continue
            if e in special_rest_emp_set: continue
            emp_end_day = resignation_dict.get(e, num_days)
            for d in range(num_days - 4):
                window_days = list(range(d+1, d+6))
                if any(day > emp_end_day for day in window_days):
                    continue
                if any(day in holiday_dict for day in window_days): continue
                r_cnt = sum(1 for day in window_days if schedule[e][day] == "休")
                if r_cnt >= 2:
                    p_score = score_short_gap * (r_cnt - 1)
                    total_penalty += p_score
                    penalty_records.append({"员工": e, "日期": f"{date_labels[d]} 至 {date_labels[d+4]}", "原因": "休息间隔过短(上不足4天班)", "扣分": p_score, "说明": f"5天内休了{r_cnt}天"})
        for d in days:
            if d in holiday_dict: continue
            b_cnt = sum(1 for e in emps if schedule[e][d] == "B")
            if b_cnt < B_min:
                total_penalty += (B_min - b_cnt) * score_B_gap
                penalty_records.append({"员工": "全员", "日期": date_labels[d-1], "原因": f"B班人数不足(仅{b_cnt}人)", "扣分": (B_min - b_cnt)*score_B_gap, "说明": f"B班少1人扣{score_B_gap}分"})
            c_cnt = sum(1 for e in emps if schedule[e][d] == "C")
            if c_cnt < C_min:
                total_penalty += (C_min - c_cnt) * score_C_gap
                penalty_records.append({"员工": "全员", "日期": date_labels[d-1], "原因": f"C班人数不足(仅{c_cnt}人)", "扣分": (C_min - c_cnt)*score_C_gap, "说明": f"C班少1人扣{score_C_gap}分"})
            c_18 = sum(1 for e in emps if schedule[e][d] in ["B", "C", "D", "Z"])
            if c_18 < 6:
                total_penalty += (6 - c_18) * score_evening_short
                penalty_records.append({"员工": "全员", "日期": date_labels[d-1], "原因": f"18-20点覆盖少{6-c_18}人", "扣分": (6 - c_18)*score_evening_short, "说明": "晚班覆盖不足"})
            c_20 = sum(1 for e in emps if schedule[e][d] in ["C", "D", "Z"])
            if c_20 < 5:
                total_penalty += (5 - c_20) * score_evening_short
                penalty_records.append({"员工": "全员", "日期": date_labels[d-1], "原因": f"20-22点覆盖少{5-c_20}人", "扣分": (5 - c_20)*score_evening_short, "说明": "晚班覆盖不足"})
        if len(pool_A) > 1:
            a_counts = {e: sum(1 for d in days if schedule[e][d] == "A") + hist_A_dict.get(e,0) for e in pool_A}
            diff_A = max(a_counts.values()) - min(a_counts.values())
            penalty_records.append({"员工":"全员","日期":"全月","原因":"A班极差情况","扣分":0,"说明":f"累计A班极差{diff_A}次(上限{A_gap_max})"})
        if len(pool_D) >1:
            d_counts = {e: sum(1 for d in days if schedule[e][d] == "D") + hist_D_dict.get(e,0) for e in pool_D}
            diff_D = max(d_counts.values()) - min(d_counts.values())
            penalty_records.append({"员工":"全员","日期":"全月","原因":"D班极差情况","扣分":0,"说明":f"累计D班极差{diff_D}次(上限{D_gap_max})"})
        if len(pool_AD) >1:
            ad_counts = {e: sum(1 for d in days if schedule[e][d] in ["A","D"]) + hist_A_dict.get(e,0)+hist_D_dict.get(e,0) for e in pool_AD}
            diff_AD = max(ad_counts.values()) - min(ad_counts.values())
            penalty_records.append({"员工":"全员","日期":"全月","原因":"A+D极差情况","扣分":0,"说明":f"累计A+D极差{diff_AD}次(上限{AD_gap_max})"})
        if len(pool_B) >1:
            b_counts = {e: sum(1 for d in days if schedule[e][d] == "B") + hist_B_dict.get(e,0) for e in pool_B}
            diff_B = max(b_counts.values()) - min(b_counts.values())
            penalty_records.append({"员工":"全员","日期":"全月","原因":"B班极差情况","扣分":0,"说明":f"累计B班极差{diff_B}次(上限{B_gap_max})"})
        if not penalty_records:
            penalty_records.append({"员工":"全员","日期":"全月","原因":"无","扣分":0,"说明":"完美满足所有软约束！"})
        df_penalty = pd.DataFrame(penalty_records)
        output_filename = f"{YEAR}年{MONTH}月排班结果.xlsx"
        try:
            with pd.ExcelWriter(output_filename) as writer:
                df_final_emp.to_excel(writer, sheet_name="表1_员工视角")
                df_day.to_excel(writer, sheet_name="表2_日期视角")
                df_penalty.to_excel(writer, sheet_name="表3_惩罚明细", index=False)
                df_summary.to_excel(writer, sheet_name="表4_本月汇总")
                df_calendar.to_excel(writer, sheet_name="表5_月历视图", index=False, header=False)
            print(f"🎉 文件已输出：{output_filename}")
        except PermissionError:
            import time
            alt_name = f"{YEAR}年{MONTH}月排班结果_防覆盖_{int(time.time())}.xlsx"
            with pd.ExcelWriter(alt_name) as writer:
                df_final_emp.to_excel(writer, sheet_name="表1_员工视角")
                df_day.to_excel(writer, sheet_name="表2_日期视角")
                df_penalty.to_excel(writer, sheet_name="表3_惩罚明细", index=False)
                df_summary.to_excel(writer, sheet_name="表4_本月汇总")
                df_calendar.to_excel(writer, sheet_name="表5_月历视图", index=False, header=False)
            print(f"⚠️原文件被占用，已另存为：{alt_name}")
        print(f"总惩罚分：{total_penalty}")
        return True
    else:
        print("❌ 无解或超时！请检查：极差上限、必上条件、假期人数配置。")
        return False

def main():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="请选择排班输入Excel文件",
        filetypes=[("Excel文件","*.xlsx"),("全部文件","*.*")]
    )
    if not file_path:
        messagebox.showinfo("提示","未选择文件，程序退出。")
        return
    messagebox.showinfo("提示",f"已选择：\n{os.path.basename(file_path)}\n点击确定开始排班，请等待……")
    try:
        ok = run_scheduling(file_path)
    except Exception as err:
        import traceback
        err_text = traceback.format_exc()
        print("\n==========完整错误堆栈==========")
        print(err_text)
        print("================================\n")
        messagebox.showerror("程序异常",f"运行出错，详情请看控制台黑窗口！\n（黑窗口里面有完整报错信息）")
        return
    if ok:
        messagebox.showinfo("完成","✅排班计算完成！结果保存在该excel同目录下。")
    else:
        messagebox.showerror("失败","❌排班无解/计算超时，请核对Excel配置！")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        print("\n=====程序崩溃=====")
        traceback.print_exc()
        print("\n按回车键关闭窗口……")
        input()
