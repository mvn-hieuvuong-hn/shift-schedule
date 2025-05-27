## Cài đặt
Tạo môi trường ảo (Virtual Environment)
```bash
python3 -m venv venv
```
Kích hoạt môi trường ảo
```bash
source venv/bin/activate
```
Cài đặt ortools, pandas, và openpyxl

```bash
pip install ortools pandas openpyxl
```

## Hướng dẫn dùng
1. Chạy file tier_1.py
```bash
(venv) hieu.vuong@MBA010060 shift-schedule % python3 tier_1.py
```
2. Nhập tháng, năm, ngày lễ
```bash
Nhập tháng (1-12): 8
Nhập năm (ví dụ: 2025): 2025

Nhập các ngày lễ trong tháng 8/2025 (ví dụ: 5, 10, 25), để trống nếu không có:
Các ngày lễ: 
```
3. Kết ra console, ví dụ
```bash
🎯 Tổng giờ trực trong tháng: 890.4 giờ
🎯 Giờ mục tiêu cho mỗi SA: 83.9 giờ
🎯 Giờ mục tiêu cho mỗi non-SA: 63.9 giờ

✅ Lịch trực đã được tạo:
                  Member  Total Hours 1-Aug 2-Aug 3-Aug 4-Aug 5-Aug 6-Aug 7-Aug 8-Aug  ... 22-Aug 23-Aug 24-Aug 25-Aug 26-Aug 27-Aug 28-Aug 29-Aug 30-Aug 31-Aug
0   Nguyễn Văn Tùng (SA)         84.8  Ca 1              Ca 3  Ca 6  Ca 4        Ca 4  ...          Ca 2   Ca 6   Ca 3                 Ca 5          Ca 3   Ca 7
1   Nguyễn Văn Tuấn (SA)         84.8  Ca 4        Ca 1  Ca 5  Ca 4        Ca 1        ...   Ca 5   Ca 8   Ca 5          Ca 4   Ca 4   Ca 2          Ca 1   Ca 6
2       Đỗ Tiến Đại (SA)         84.8  Ca 3              Ca 4  Ca 3        Ca 4  Ca 3  ...   Ca 4   Ca 6   Ca 7   Ca 4   Ca 5          Ca 4   Ca 4   Ca 8       
3      Nguyễn Ngọc Khánh         62.8        Ca 3  Ca 6                    Ca 5        ...          Ca 3          Ca 1          Ca 1          Ca 3          Ca 5
4          Phạm Đức Long         63.2        Ca 8  Ca 8        Ca 1        Ca 6        ...   Ca 6          Ca 8   Ca 5                 Ca 3          Ca 7       
5         Nguyễn Viết Tú         63.2        Ca 4  Ca 2  Ca 6        Ca 2        Ca 2  ...          Ca 7          Ca 6                 Ca 1   Ca 6          Ca 3
6   Nguyễn Bá Tuấn Nghĩa         64.8  Ca 5  Ca 7  Ca 7              Ca 5              ...   Ca 2                        Ca 2          Ca 6                 Ca 2
7      Trần Thị Ngọc Ánh         64.6        Ca 2              Ca 2        Ca 3  Ca 5  ...          Ca 1                 Ca 1   Ca 5          Ca 2          Ca 1
8        Nguyễn Đăng Quý         63.4              Ca 3              Ca 1        Ca 1  ...                 Ca 4                 Ca 2          Ca 5   Ca 6   Ca 4
9         Đặng Xuân Dũng         63.4  Ca 2  Ca 5        Ca 1                          ...   Ca 3   Ca 4   Ca 1                               Ca 1   Ca 5   Ca 8
10       Vương Đình Hiếu         63.2  Ca 6        Ca 4        Ca 5        Ca 2        ...                        Ca 2   Ca 6   Ca 6                 Ca 2       
11         Lê Triệu Sáng         64.6        Ca 1                    Ca 6        Ca 6  ...   Ca 1          Ca 2                 Ca 3                 Ca 4       
12      Nguyễn Xuân Khoa         62.8        Ca 6  Ca 5  Ca 2        Ca 3              ...          Ca 5   Ca 3          Ca 3                                   

[13 rows x 33 columns]
📁 Đã lưu lịch trực vào file: ShiftSchedule_Export/Lich_truc_8_2025.xlsx
```
4. Kết quả đầu ra được lưu trong folder ShiftSchedule_Export

## Ràng buộc

1. Mỗi người chỉ 1 ca/ngày
2. Ca 4 trong tuần chỉ SA, chia đều ca 4 cho các thành viên SA
3. Không cho cùng lúc 3 ca 1,2,3 liền nhau
4. Không có ca 1,2,3 sau ca muộn (Ca 5, Ca 6 trong tuần; Ca 7, Ca 8 cuối tuần/lễ)
5. Tổng số ca 1,2,3 trong tuần, cuối tuần cho mỗi thành viên gần bằng nhau (sai lệch +/- 1 ca)
6. Phân bổ đều ca 1,2,3 cho các thành viên (sai lệch +/- 1 ca)
7. SA hơn non-SA khoảng 20 giờ
