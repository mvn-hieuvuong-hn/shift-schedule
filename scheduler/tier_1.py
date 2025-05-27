from ortools.sat.python import cp_model
import pandas as pd
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import re
import calendar
from datetime import date
import os
import random

from members.tier_1_members import members
from weight.tier_1_weight_map import weight_map

# Define scaled_weight_map for integer arithmetic (multiply by 10)
scaled_weight_map = {
    'weekday': {k: int(v * 10) for k, v in weight_map['weekday'].items()},
    'weekend': {k: int(v * 10) for k, v in weight_map['weekend'].items()},
    'holiday': {k: int(v * 10) for k, v in weight_map['holiday'].items()}
}

# Hàm classify_day
def classify_day(d_obj, holiday_list_param):
    if d_obj in holiday_list_param:
        return 'holiday'
    elif d_obj.weekday() >= 5:
        return 'weekend'
    return 'weekday'

def generate_tier1_schedule_file(desired_month, desired_year, holiday_input_str, output_dir):
    
    try:
        # Tính toán các biến phụ thuộc vào `members` và `weight_map`
        num_members = len(members)
        sa_members_local = [m for m in members if "(SA)" in m]
        non_sa_members_local = [m for m in members if m not in sa_members_local]
        num_sa = len(sa_members_local)
        member_to_id = {m: i for i, m in enumerate(members)}
        sa_members_indices = [member_to_id[m] for m in sa_members_local]
        non_sa_members_indices = [member_to_id[m] for m in non_sa_members_local]
        random_start_offset = random.randint(0, num_members - 1) if num_members > 0 else 0

        # --- Xử lý ngày lễ ---
        holiday_list = []
        if holiday_input_str:
            day_strings = re.split(r'[, ]+', holiday_input_str)
            _, max_day_in_month = calendar.monthrange(desired_year, desired_month)
            for day_str in day_strings:
                try:
                    day_num = int(day_str)
                    if 1 <= day_num <= max_day_in_month:
                        holiday_list.append(date(desired_year, desired_month, day_num))
                    else:
                        print(f"Cảnh báo: Ngày '{day_str}' không hợp lệ cho tháng {desired_month}/{desired_year} và sẽ bị bỏ qua.")
                except ValueError:
                    print(f"Cảnh báo: '{day_str}' không phải là một số hợp lệ và sẽ bị bỏ qua.")
            holiday_list = sorted(list(set(holiday_list)))

        # --- Chuẩn bị dữ liệu lịch ---
        _, num_days = calendar.monthrange(desired_year, desired_month)
        date_list = [date(desired_year, desired_month, d) for d in range(1, num_days + 1)]

        weekdays_idx = [i for i, d in enumerate(date_list) if d.weekday() < 5 and d not in holiday_list]
        weekends_idx = [i for i, d in enumerate(date_list) if d.weekday() >= 5 or d in holiday_list]

        monthly_total_weighted_hours_float = 0
        for day_num in range(1, num_days + 1):
            current_date = date(desired_year, desired_month, day_num)
            day_type = classify_day(current_date, holiday_list)
            monthly_total_weighted_hours_float += sum(weight_map.get(day_type, {}).values())

        total_shifts_per_type = {}
        for day_idx, day_obj in enumerate(date_list):
            day_type = classify_day(day_obj, holiday_list)
            current_day_shifts_info = weight_map.get(day_type, {})
            for shift in current_day_shifts_info.keys():
                total_shifts_per_type[(day_type, shift)] = total_shifts_per_type.get((day_type, shift), 0) + 1

        # --- Xây dựng mô hình CP-SAT ---
        model = cp_model.CpModel()
        shift_vars = {}
        all_shifts_by_day = {}

        for day_idx, day_obj in enumerate(date_list):
            day_type = classify_day(day_obj, holiday_list)
            shifts = list(weight_map.get(day_type, {}).keys())
            all_shifts_by_day[day_idx] = shifts
            for shift in shifts:
                for m_idx in range(num_members):
                    shift_vars[(day_idx, shift, m_idx)] = model.NewBoolVar(f"shift_d{day_idx}_s{shift}_m{m_idx}")

        # --- Ràng buộc mỗi ca phải có 1 người ---
        for day_idx in range(len(date_list)):
            shifts = all_shifts_by_day[day_idx]
            for shift in shifts:
                model.AddExactlyOne(shift_vars[(day_idx, shift, m_idx)] for m_idx in range(num_members))

        # --- Ràng buộc mỗi người chỉ 1 ca/ngày ---
        for day_idx in range(len(date_list)):
            shifts = all_shifts_by_day[day_idx]
            for m_idx in range(num_members):
                model.AddAtMostOne(shift_vars[(day_idx, shift, m_idx)] for shift in shifts)

        # --- Ràng buộc Ca 4 trong tuần chỉ SA ---
        for day_idx, day_obj in enumerate(date_list):
            day_type = classify_day(day_obj, holiday_list)
            if day_type == 'weekday':
                if "Ca 4" in all_shifts_by_day[day_idx]:
                    for m_idx in range(num_members):
                        if members[m_idx] not in sa_members_local:
                            model.Add(shift_vars[(day_idx, "Ca 4", m_idx)] == 0)

        # --- Ràng buộc không cho người cùng lúc 3 ca 1,2,3 liền nhau ---
        ca123 = ["Ca 1", "Ca 2", "Ca 3"]
        for m_idx in range(num_members):
            for day_idx in range(len(date_list) - 1):
                sum_shifts_d = sum(shift_vars[(day_idx, s, m_idx)] for s in ca123 if s in all_shifts_by_day[day_idx])
                sum_shifts_d_plus_1 = sum(shift_vars[(day_idx + 1, s, m_idx)] for s in ca123 if s in all_shifts_by_day[day_idx+1])
                model.Add(sum_shifts_d + sum_shifts_d_plus_1 <= 1)

        # --- Ràng buộc nghỉ ngơi sau ca muộn/đêm (Ca 5, Ca 6 trong tuần; Ca 7, Ca 8 cuối tuần/lễ) 
        for m_idx in range(num_members):
            for day_idx in range(len(date_list) - 1):
                current_day = date_list[day_idx]
                day_type = classify_day(current_day, holiday_list)

                shifts_today_requiring_rest = []
                if day_type == 'weekday':
                    if "Ca 5" in all_shifts_by_day[day_idx]: shifts_today_requiring_rest.append("Ca 5")
                    if "Ca 6" in all_shifts_by_day[day_idx]: shifts_today_requiring_rest.append("Ca 6")
                elif day_type in ['weekend', 'holiday']:
                    if "Ca 7" in all_shifts_by_day[day_idx]: shifts_today_requiring_rest.append("Ca 7")
                    if "Ca 8" in all_shifts_by_day[day_idx]: shifts_today_requiring_rest.append("Ca 8")

                for shift_today in shifts_today_requiring_rest:
                    shift_today_var = shift_vars[(day_idx, shift_today, m_idx)]
                    for next_day_shift in ca123:
                        if next_day_shift in all_shifts_by_day[day_idx + 1]:
                            shift_tomorrow_ca_var = shift_vars[(day_idx + 1, next_day_shift, m_idx)]
                            model.Add(shift_today_var + shift_tomorrow_ca_var <= 1)

        # --- Ràng buộc: Chia đều Ca 4 trong tuần cho thành viên SA ---
        total_ca4_weekday_shifts_count = 0
        for day_idx in weekdays_idx:
            if "Ca 4" in all_shifts_by_day[day_idx]: total_ca4_weekday_shifts_count += 1
        if total_ca4_weekday_shifts_count > 0 and num_sa > 0:
            base_ca4_per_sa = total_ca4_weekday_shifts_count // num_sa
            remainder_ca4 = total_ca4_weekday_shifts_count % num_sa
            ca4_sa_tolerance = 1
            for i, sa_idx in enumerate(sa_members_indices):
                sum_ca4_per_sa_var = model.NewIntVar(0, num_days, f"sa_{sa_idx}_ca4_weekday_sum")
                ca4_shifts_for_this_sa = []
                for day_idx_in_weekday_idx in weekdays_idx:
                    if "Ca 4" in all_shifts_by_day[day_idx_in_weekday_idx]: ca4_shifts_for_this_sa.append(shift_vars[(day_idx_in_weekday_idx, "Ca 4", sa_idx)])
                if ca4_shifts_for_this_sa: model.Add(sum_ca4_per_sa_var == sum(ca4_shifts_for_this_sa))
                else: model.Add(sum_ca4_per_sa_var == 0)
                min_target = base_ca4_per_sa
                max_target = base_ca4_per_sa + (1 if (m_idx + random_start_offset) % num_members < remainder_ca4 else 0)
                model.Add(sum_ca4_per_sa_var >= max(0, min_target - ca4_sa_tolerance))
                model.Add(sum_ca4_per_sa_var <= max_target + ca4_sa_tolerance)

        # --- Ràng buộc cân bằng ca trong tuần cho mỗi thành viên ---
        shift_weekday_to_balance = ["Ca 1", "Ca 2", "Ca 3", "Ca 5", "Ca 6"]
        shift_weekend_holiday_to_balance = ["Ca 1", "Ca 2", "Ca 3", "Ca 4", "Ca 5", "Ca 6", "Ca 7", "Ca 8"]
        individual_shift_tolerance = 1
        # --- Ràng buộc cân bằng từng ca trong tuần cho mỗi thành viên (Ca 1,2,3,5,6) ---
        print(f"\n⚡️ Ràng buộc cân bằng từng ca trong tuần ({', '.join(shift_weekday_to_balance)}) cho mỗi thành viên.")

        for m_idx in range(num_members):
            for shift_type in shift_weekday_to_balance:
                total_individual_shift_weekday_m = model.NewIntVar(0, num_days, f"total_{shift_type}_weekday_m{m_idx}")
                individual_shift_expr = []
                for day_idx in weekdays_idx:
                    if shift_type in all_shifts_by_day[day_idx]:
                        individual_shift_expr.append(shift_vars[(day_idx, shift_type, m_idx)])
                
                # Link the sum variable to the actual shift variables
                if individual_shift_expr:
                    model.Add(total_individual_shift_weekday_m == sum(individual_shift_expr))
                else:
                    model.Add(total_individual_shift_weekday_m == 0) # No shifts possible, so sum is 0

                # Tính tổng số lần 'shift_type' này xuất hiện trong tất cả các ngày trong tuần
                total_occurrences_of_this_shift_weekday = 0
                for day_idx in weekdays_idx:
                    if shift_type in all_shifts_by_day[day_idx]:
                        total_occurrences_of_this_shift_weekday += 1
                
                if total_occurrences_of_this_shift_weekday > 0 and num_members > 0:
                    avg_individual_shift_per_member = total_occurrences_of_this_shift_weekday // num_members
                    remainder_individual_shift = total_occurrences_of_this_shift_weekday % num_members
                    min_target_individual_shift = avg_individual_shift_per_member
                    max_target_individual_shift = avg_individual_shift_per_member + (1 if (m_idx + random_start_offset) % num_members < remainder_individual_shift else 0)
                    
                    model.Add(total_individual_shift_weekday_m >= max(0, min_target_individual_shift - individual_shift_tolerance))
                    model.Add(total_individual_shift_weekday_m <= max_target_individual_shift + individual_shift_tolerance)
                    
                    print(f"    - Thành viên {members[m_idx]}: Ca {shift_type} trong tuần [{max(0, min_target_individual_shift - individual_shift_tolerance)}-{max_target_individual_shift + individual_shift_tolerance}]")
                else:
                    model.Add(total_individual_shift_weekday_m == 0)
                    print(f"    - Thành viên {members[m_idx]}: Ca {shift_type} trong tuần [0-0] (Không có ca hoặc không có thành viên để phân bổ)")


        # --- Ràng buộc cân bằng từng ca cuối tuần/lễ cho mỗi thành viên (Ca 1,2,3,4,5,6,7,8) ---
        print(f"\n⚡️ Ràng buộc cân bằng từng ca cuối tuần/lễ ({', '.join(shift_weekend_holiday_to_balance)}) cho mỗi thành viên.")

        for m_idx in range(num_members):
            for shift_type in shift_weekend_holiday_to_balance:
                total_individual_shift_weekend_m = model.NewIntVar(0, num_days, f"total_{shift_type}_weekend_m{m_idx}")
                individual_shift_expr = []
                for day_idx in weekends_idx: # Chỉ xét các ngày cuối tuần và ngày lễ
                    if shift_type in all_shifts_by_day[day_idx]:
                        individual_shift_expr.append(shift_vars[(day_idx, shift_type, m_idx)])
                
                # Link the sum variable to the actual shift variables
                if individual_shift_expr:
                    model.Add(total_individual_shift_weekend_m == sum(individual_shift_expr))
                else:
                    model.Add(total_individual_shift_weekend_m == 0)

                # Tính tổng số lần 'shift_type' này xuất hiện trong tất cả các ngày cuối tuần/lễ
                total_occurrences_of_this_shift_weekend = 0
                for day_idx in weekends_idx:
                    if shift_type in all_shifts_by_day[day_idx]:
                        total_occurrences_of_this_shift_weekend += 1
                
                if total_occurrences_of_this_shift_weekend > 0 and num_members > 0:
                    avg_individual_shift_per_member = total_occurrences_of_this_shift_weekend // num_members
                    remainder_individual_shift = total_occurrences_of_this_shift_weekend % num_members
                    min_target_individual_shift = avg_individual_shift_per_member
                    max_target_individual_shift = avg_individual_shift_per_member + (1 if (m_idx + random_start_offset) % num_members < remainder_individual_shift else 0)
                    model.Add(total_individual_shift_weekend_m >= max(0, min_target_individual_shift - individual_shift_tolerance))
                    model.Add(total_individual_shift_weekend_m <= max_target_individual_shift + individual_shift_tolerance)
                    
                    print(f"    - Thành viên {members[m_idx]}: Ca {shift_type} cuối tuần/lễ [{max(0, min_target_individual_shift - individual_shift_tolerance)}-{max_target_individual_shift + individual_shift_tolerance}]")
                else:
                    model.Add(total_individual_shift_weekend_m == 0)
                    print(f"    - Thành viên {members[m_idx]}: Ca {shift_type} cuối tuần/lễ [0-0] (Không có ca hoặc không có thành viên để phân bổ)")

        # --- Ràng buộc và mục tiêu tổng giờ. SA hơn non-SA khoảng 20 giờ---
        total_hours = []
        for m_idx in range(num_members):
            hours_expr_components = []
            for day_idx, day_obj_in_loop in enumerate(date_list):
                day_type = classify_day(day_obj_in_loop, holiday_list)
                current_day_scaled_shifts = scaled_weight_map.get(day_type, {})
                for shift in all_shifts_by_day[day_idx]:
                    if shift in current_day_scaled_shifts:
                        weight = current_day_scaled_shifts[shift]
                        hours_expr_components.append(shift_vars[(day_idx, shift, m_idx)] * weight)
            total_hours.append(model.NewIntVar(0, 100000, f"total_hours_m{m_idx}"))
            model.Add(total_hours[m_idx] == sum(hours_expr_components))

        avg_sa = model.NewIntVar(0, 100000, "avg_sa")
        avg_non_sa = model.NewIntVar(0, 100000, "avg_non_sa")

        if num_sa > 0: model.Add(avg_sa * num_sa == sum(total_hours[m_idx] for m_idx in sa_members_indices))
        else: model.Add(avg_sa == 0)
        if len(non_sa_members_indices) > 0: model.Add(avg_non_sa * len(non_sa_members_indices) == sum(total_hours[m_idx] for m_idx in non_sa_members_indices))
        else: model.Add(avg_non_sa == 0)
        if num_sa > 0 and len(non_sa_members_indices) > 0: model.Add(avg_sa >= avg_non_sa + 200)

        total_members_count = num_sa + len(non_sa_members_indices)
        if total_members_count > 0:
            x_float = (monthly_total_weighted_hours_float - num_sa * 20.0) / total_members_count
        else: x_float = 0.0
        target_non_sa_float = x_float
        target_sa_float = x_float + 20.0
        target_non_sa_scaled = int(target_non_sa_float * 10)
        target_sa_scaled = int(target_sa_float * 10)
        tolerance_scaled = 10

        for m_idx in sa_members_indices: model.Add(total_hours[m_idx] >= target_sa_scaled - tolerance_scaled); model.Add(total_hours[m_idx] <= target_sa_scaled + tolerance_scaled)
        for m_idx in non_sa_members_indices: model.Add(total_hours[m_idx] >= target_non_sa_scaled - tolerance_scaled); model.Add(total_hours[m_idx] <= target_non_sa_scaled + tolerance_scaled)

        # --- Solver ---
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 300
        status = solver.Solve(model)

        # --- Kết quả ---
        if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
            data = []
            holiday_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
            weekend_fill = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")

            for m_idx in range(num_members):
                row_dict = {}; row_dict["Member"] = members[m_idx]; row_dict["Total Hours"] = solver.Value(total_hours[m_idx]) / 10
                for day_idx, day_obj_in_loop in enumerate(date_list):
                    assigned_shifts = [s for s in all_shifts_by_day[day_idx] if solver.Value(shift_vars[(day_idx, s, m_idx)]) == 1]
                    month_abbreviation = calendar.month_abbr[day_obj_in_loop.month]
                    display_value = assigned_shifts[0] if assigned_shifts else ""
                    row_dict[f"{day_obj_in_loop.day}-{month_abbreviation}"] = display_value
                data.append(row_dict)
            df = pd.DataFrame(data)

            total_sum_hours = df["Total Hours"].sum()
            total_row_dict = {"Member": "Tổng cộng", "Total Hours": total_sum_hours}
            for col in df.columns:
                if col not in ["Member", "Total Hours"]:
                    total_row_dict[col] = "" # Các cột ngày không có giá trị tổng
            df = pd.concat([df, pd.DataFrame([total_row_dict])], ignore_index=True)
            

            # Thay đổi logic lưu file để trả về cho Flask
            output_filename_base = f"Lich_truc_{desired_month}_{desired_year}.xlsx"
            output_filepath = os.path.join(output_dir, output_filename_base)

            workbook = Workbook()
            if 'Sheet' in workbook.sheetnames: workbook.remove(workbook['Sheet'])
            sheet = workbook.create_sheet('LichTruc')
            for r_idx, r in enumerate(dataframe_to_rows(df, index=False, header=True)): sheet.append(r)

            day_column_start_excel_idx = 3
            sheet.column_dimensions[get_column_letter(1)].width = 20
            sheet.column_dimensions[get_column_letter(2)].width = 10

            for col_offset, day_obj_in_loop in enumerate(date_list):
                current_col_excel_idx = day_column_start_excel_idx + col_offset
                col_letter = get_column_letter(current_col_excel_idx)
                sheet.column_dimensions[col_letter].width = 8
                fill_color = None
                if day_obj_in_loop in holiday_list: fill_color = holiday_fill
                elif day_obj_in_loop.weekday() >= 5: fill_color = weekend_fill
                if fill_color:
                    for row_num in range(1, sheet.max_row + 1): sheet[f'{col_letter}{row_num}'].fill = fill_color

            workbook.save(output_filepath)

            return True, "Lịch trực đã được tạo thành công!", output_filepath
        else:
            return False, f"Không tìm được lịch phù hợp. Trạng thái: {solver.StatusName(status)}", None
    except Exception as e:
        return False, f"Đã xảy ra lỗi trong quá trình tạo lịch: {e}", None

