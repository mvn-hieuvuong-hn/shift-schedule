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

        day_types = [classify_day(d, holiday_list) for d in date_list]
        monthly_total_weighted_hours_float = sum(
            sum(weight_map.get(day_types[i], {}).values()) for i in range(num_days)
        )

        total_shifts_per_type = {}
        for i, t in enumerate(day_types):
            for shift in weight_map.get(t, {}):
                total_shifts_per_type[(t, shift)] = total_shifts_per_type.get((t, shift), 0) + 1

        # --- Xây dựng mô hình CP-SAT ---
        model = cp_model.CpModel()
        shift_vars = {}
        all_shifts_by_day = {}

        # --- Khởi tạo ca ---
        ca123 = ["Ca 1", "Ca 2", "Ca 3"]
        shift_weekday_to_balance = ["Ca 1", "Ca 2", "Ca 3", "Ca 5", "Ca 6"]
        shift_weekend_holiday_to_balance = ["Ca 1", "Ca 2", "Ca 3", "Ca 4", "Ca 5", "Ca 6", "Ca 7", "Ca 8"]

        for day_idx, day_type in enumerate(day_types):
            shifts = list(weight_map.get(day_type, {}).keys())
            all_shifts_by_day[day_idx] = shifts
            for shift in shifts:
                for m_idx in range(num_members):
                    shift_vars[(day_idx, shift, m_idx)] = model.NewBoolVar(f"shift_d{day_idx}_s{shift}_m{m_idx}")

        for day_idx, shifts in all_shifts_by_day.items():
            # --- Ràng buộc mỗi ca phải có 1 người ---
            for shift in shifts:
                model.AddExactlyOne(shift_vars[(day_idx, shift, m_idx)] for m_idx in range(num_members))
            # --- Ràng buộc mỗi người chỉ 1 ca/ngày ---
            for m_idx in range(num_members):
                model.AddAtMostOne(shift_vars[(day_idx, shift, m_idx)] for shift in shifts)

        # --- Ràng buộc không cho một người trực các ca trong tuần liền nhau ---
        for m_idx in range(num_members):
            # Chỉ xét các cặp ngày thường liên tiếp để giảm số ràng buộc
            for i in range(len(weekdays_idx) - 1):
                day_idx = weekdays_idx[i]
                next_day_idx = weekdays_idx[i + 1]
                # Đảm bảo hai ngày này liên tiếp nhau trong lịch
                if next_day_idx == day_idx + 1:
                    shifts_today = [s for s in shift_weekday_to_balance if s in all_shifts_by_day[day_idx]]
                    shifts_tomorrow = [s for s in shift_weekday_to_balance if s in all_shifts_by_day[next_day_idx]]
                    if shifts_today and shifts_tomorrow:
                        sum_shifts_d = sum(shift_vars[(day_idx, s, m_idx)] for s in shifts_today)
                        sum_shifts_d_plus_1 = sum(shift_vars[(next_day_idx, s, m_idx)] for s in shifts_tomorrow)
                        model.Add(sum_shifts_d + sum_shifts_d_plus_1 <= 1)

        # --- Ràng buộc nghỉ ngơi sau ca muộn/đêm (Ca 5, Ca 6 trong tuần; Ca 7, Ca 8 cuối tuần/lễ) 
        for day_idx in range(len(date_list) - 1):
            day_type = day_types[day_idx]
            if day_type == 'weekday':
                shifts_today_requiring_rest = [s for s in ("Ca 5", "Ca 6") if s in all_shifts_by_day[day_idx]]
            elif day_type in ('weekend', 'holiday'):
                shifts_today_requiring_rest = [s for s in ("Ca 7", "Ca 8") if s in all_shifts_by_day[day_idx]]
            else:
                shifts_today_requiring_rest = []
            if not shifts_today_requiring_rest:
                continue
            shifts_tomorrow = [s for s in ca123 if s in all_shifts_by_day[day_idx + 1]]
            if not shifts_tomorrow:
                continue
            for m_idx in range(num_members):
                for shift_today in shifts_today_requiring_rest:
                    shift_today_var = shift_vars[(day_idx, shift_today, m_idx)]
                    for next_day_shift in shifts_tomorrow:
                        shift_tomorrow_ca_var = shift_vars[(day_idx + 1, next_day_shift, m_idx)]
                        model.Add(shift_today_var + shift_tomorrow_ca_var <= 1)

        # --- Ràng buộc Ca 4 trong tuần chỉ SA và chia đều Ca 4 cho SA ---
        ca4_weekday_day_indices = [day_idx for day_idx in weekdays_idx if "Ca 4" in all_shifts_by_day[day_idx]]

        # Cấm non-SA trực Ca 4
        for day_idx in ca4_weekday_day_indices:
            for m_idx in non_sa_members_indices:
                model.Add(shift_vars[(day_idx, "Ca 4", m_idx)] == 0)

        # Chia đều Ca 4 cho SA
        total_ca4_weekday_shifts_count = len(ca4_weekday_day_indices)
        if total_ca4_weekday_shifts_count > 0 and num_sa > 0:
            base_ca4_per_sa = total_ca4_weekday_shifts_count // num_sa
            remainder_ca4 = total_ca4_weekday_shifts_count % num_sa
            ca4_sa_tolerance = 1
            for i, sa_idx in enumerate(sa_members_indices):
                ca4_shifts_for_this_sa = [
                    shift_vars[(day_idx, "Ca 4", sa_idx)]
                    for day_idx in ca4_weekday_day_indices
                ]
                sum_ca4_per_sa_var = model.NewIntVar(0, num_days, f"sa_{sa_idx}_ca4_weekday_sum")
                model.Add(sum_ca4_per_sa_var == sum(ca4_shifts_for_this_sa)) if ca4_shifts_for_this_sa else model.Add(sum_ca4_per_sa_var == 0)
                min_target = base_ca4_per_sa
                max_target = base_ca4_per_sa + (1 if (i + random_start_offset) % num_sa < remainder_ca4 else 0)
                model.Add(sum_ca4_per_sa_var >= max(0, min_target - ca4_sa_tolerance))
                model.Add(sum_ca4_per_sa_var <= max_target + ca4_sa_tolerance)

        # --- Ràng buộc cân bằng từng loại ca trong tuần cho mỗi thành viên (Ca 1,2,3,5,6) ---
        individual_shift_tolerance = 1
        total_occurrences_per_shift_weekday = {
            shift_type: sum(1 for day_idx in weekdays_idx if shift_type in all_shifts_by_day[day_idx])
            for shift_type in shift_weekday_to_balance
        }
        for m_idx in range(num_members):
            for shift_type in shift_weekday_to_balance:
                total_individual_shift_weekday_m = model.NewIntVar(0, num_days, f"total_{shift_type}_weekday_m{m_idx}")
                individual_shift_expr = [
                    shift_vars[(day_idx, shift_type, m_idx)]
                    for day_idx in weekdays_idx if shift_type in all_shifts_by_day[day_idx]
                ]
                model.Add(total_individual_shift_weekday_m == sum(individual_shift_expr)) if individual_shift_expr else model.Add(total_individual_shift_weekday_m == 0)
                total_occurrences = total_occurrences_per_shift_weekday[shift_type]
                if total_occurrences > 0 and num_members > 0:
                    avg_individual_shift_per_member = total_occurrences // num_members
                    remainder_individual_shift = total_occurrences % num_members
                    min_target = avg_individual_shift_per_member
                    max_target = avg_individual_shift_per_member + (1 if (m_idx + random_start_offset) % num_members < remainder_individual_shift else 0)
                    model.Add(total_individual_shift_weekday_m >= max(0, min_target - individual_shift_tolerance))
                    model.Add(total_individual_shift_weekday_m <= max_target + individual_shift_tolerance)
                else:
                    model.Add(total_individual_shift_weekday_m == 0)

        # --- Ràng buộc cân bằng từng loại ca cuối tuần/lễ cho mỗi thành viên (Ca 1,2,3,4,5,6,7,8) ---
        total_occurrences_per_shift_weekend = {
            shift_type: sum(1 for day_idx in weekends_idx if shift_type in all_shifts_by_day[day_idx])
            for shift_type in shift_weekend_holiday_to_balance
        }
        for m_idx in range(num_members):
            for shift_type in shift_weekend_holiday_to_balance:
                total_individual_shift_weekend_m = model.NewIntVar(0, num_days, f"total_{shift_type}_weekend_m{m_idx}")
                individual_shift_expr = [
                    shift_vars[(day_idx, shift_type, m_idx)]
                    for day_idx in weekends_idx if shift_type in all_shifts_by_day[day_idx]
                ]
                model.Add(total_individual_shift_weekend_m == sum(individual_shift_expr)) if individual_shift_expr else model.Add(total_individual_shift_weekend_m == 0)
                total_occurrences = total_occurrences_per_shift_weekend[shift_type]
                if total_occurrences > 0 and num_members > 0:
                    avg_individual_shift_per_member = total_occurrences // num_members
                    remainder_individual_shift = total_occurrences % num_members
                    min_target = avg_individual_shift_per_member
                    max_target = avg_individual_shift_per_member + (1 if (m_idx + random_start_offset) % num_members < remainder_individual_shift else 0)
                    model.Add(total_individual_shift_weekend_m >= max(0, min_target - individual_shift_tolerance))
                    model.Add(total_individual_shift_weekend_m <= max_target + individual_shift_tolerance)
                else:
                    model.Add(total_individual_shift_weekend_m == 0)

        # --- Bổ sung: Ràng buộc TỔNG số ca 1,2,3 cho mỗi thành viên ---
        ca123_day_shift_pairs = [
            (day_idx, shift)
            for day_idx in range(len(date_list))
            for shift in ca123
            if shift in all_shifts_by_day[day_idx]
        ]
        total_possible_overall_ca123_sum = len(ca123_day_shift_pairs)
        avg_overall_ca123_per_member_sum = total_possible_overall_ca123_sum // num_members
        remainder_overall_ca123_sum = total_possible_overall_ca123_sum % num_members
        for m_idx in range(num_members):
            total_night_shifts_overall_m = model.NewIntVar(0, num_days, f"total_night_overall_m{m_idx}")
            night_shifts_overall_expr = [
                shift_vars[(day_idx, shift, m_idx)]
                for (day_idx, shift) in ca123_day_shift_pairs
            ]
            model.Add(total_night_shifts_overall_m == sum(night_shifts_overall_expr))
            rotated_m_idx_for_remainder = (m_idx + random_start_offset) % num_members
            min_target_night_overall_sum = avg_overall_ca123_per_member_sum
            max_target_night_overall_sum = avg_overall_ca123_per_member_sum + (1 if rotated_m_idx_for_remainder < remainder_overall_ca123_sum else 0)
            model.Add(total_night_shifts_overall_m >= max(0, min_target_night_overall_sum))
            model.Add(total_night_shifts_overall_m <= max_target_night_overall_sum)

        # --- Tính tổng giờ---
        total_hours = []
        for m_idx in range(num_members):
            hours_expr_components = []
            for day_idx, day_type in enumerate(day_types):
                current_day_scaled_shifts = scaled_weight_map.get(day_type, {})
                for shift in all_shifts_by_day[day_idx]:
                    if shift in current_day_scaled_shifts:
                        weight = current_day_scaled_shifts[shift]
                        hours_expr_components.append(shift_vars[(day_idx, shift, m_idx)] * weight)
            total_hours_var = model.NewIntVar(0, 10000, f"total_hours_m{m_idx}")
            model.Add(total_hours_var == sum(hours_expr_components))
            total_hours.append(total_hours_var)

        # --- Ràng buộc và mục tiêu tổng giờ. SA hơn non-SA khoảng 20 giờ ---
        total_members_count = num_sa + len(non_sa_members_indices)
        x_float = (monthly_total_weighted_hours_float - num_sa * 20.0) / total_members_count
        target_non_sa_float = x_float
        target_sa_float = x_float + 20.0
        target_non_sa_scaled = int(target_non_sa_float * 10)
        target_sa_scaled = int(target_sa_float * 10)
        tolerance_scaled = 10 # Cho phép sai số 1 giờ

        for m_idx in sa_members_indices:
            model.Add(total_hours[m_idx] >= target_sa_scaled - tolerance_scaled)
            model.Add(total_hours[m_idx] <= target_sa_scaled + tolerance_scaled)
        for m_idx in non_sa_members_indices:
            model.Add(total_hours[m_idx] >= target_non_sa_scaled - tolerance_scaled)
            model.Add(total_hours[m_idx] <= target_non_sa_scaled + tolerance_scaled)

        # --- Solver ---
        solver = cp_model.CpSolver()
        solver.parameters.max_time_in_seconds = 300
        status = solver.Solve(model)

        # --- Kết quả ---
        if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
            data = []
            holiday_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
            weekend_fill = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")

            month_abbr = calendar.month_abbr 
            day_columns = [f"{d.day}-{month_abbr[d.month]}" for d in date_list]

            for m_idx in range(num_members):
                row_dict = {
                    "Member": members[m_idx],
                    "Total Hours": solver.Value(total_hours[m_idx]) / 10
                }
                assigned_shifts = [
                    next((s for s in all_shifts_by_day[day_idx] if solver.Value(shift_vars[(day_idx, s, m_idx)]) == 1), "")
                    for day_idx in range(len(date_list))
                ]
                for col, shift in zip(day_columns, assigned_shifts):
                    row_dict[col] = shift
                data.append(row_dict)

            df = pd.DataFrame(data)
            total_sum_hours = df["Total Hours"].sum()
            total_row_dict = {"Member": "Tổng cộng", "Total Hours": total_sum_hours}
            for col in df.columns:
                if col not in ["Member", "Total Hours"]:
                    total_row_dict[col] = ""
            df = pd.concat([df, pd.DataFrame([total_row_dict])], ignore_index=True)
            
            # Thay đổi logic lưu file để trả về cho Flask
            output_filename_base = f"Lich_truc_{desired_month}_{desired_year}.xlsx"
            output_filepath = os.path.join(output_dir, output_filename_base)

            workbook = Workbook()
            if 'Sheet' in workbook.sheetnames:
                workbook.remove(workbook['Sheet'])
            sheet = workbook.create_sheet('LichTruc')

            for r in dataframe_to_rows(df, index=False, header=True):
                sheet.append(r)

            # Đặt độ rộng cột cho cột Member và Total Hours
            sheet.column_dimensions[get_column_letter(1)].width = 20
            sheet.column_dimensions[get_column_letter(2)].width = 10

            # Tạo sẵn danh sách màu cho từng ngày để tránh lặp lại phép kiểm tra
            day_fills = []
            for day_obj in date_list:
                if day_obj in holiday_list:
                    day_fills.append(holiday_fill)
                elif day_obj.weekday() >= 5:
                    day_fills.append(weekend_fill)
                else:
                    day_fills.append(None)

            # Đặt độ rộng và màu cho các cột ngày
            day_column_start_excel_idx = 3
            for col_offset, fill_color in enumerate(day_fills):
                col_letter = get_column_letter(day_column_start_excel_idx + col_offset)
                sheet.column_dimensions[col_letter].width = 8
                if fill_color:
                    for row_num in range(1, sheet.max_row + 1):
                        sheet[f'{col_letter}{row_num}'].fill = fill_color

            workbook.save(output_filepath)

            return True, "Lịch trực đã được tạo thành công!", output_filepath
        else:
            return False, f"Không tìm được lịch phù hợp. Trạng thái: {solver.StatusName(status)}", None
    except Exception as e:
        return False, f"Đã xảy ra lỗi trong quá trình tạo lịch: {e}", None

