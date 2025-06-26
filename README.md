## Cài đặt
Tạo môi trường ảo (Virtual Environment)
```bash
python3 -m venv venv
```
Kích hoạt môi trường ảo
```bash
source venv/bin/activate
```
Cài đặt thư viện

```bash
pip install -r requirements.txt
```

## Hướng dẫn dùng
1. Chạy file app.py
```bash
(venv) hieu.vuong@MBA010060 shift-schedule % python3 app.py
```
2. Truy cập URL: http://127.0.0.1:5000/
3. Nhập ngày tháng
![Home](docs/home.png)
4. Tải file về 
![Result](docs/result.png)
5. File excel
![Result_home](docs/result_excel.png)

## Ràng buộc

1. Mỗi ca phải có 1 người, mỗi người chỉ 1 ca/ngày
2. Không cho một người trực các ca trong tuần liền nhau, cách nhau 1 ngày
3. Trong 5 ngày liên tục giới hạn 2 ca đêm
4. Không có ca 1,2,3 sau ca muộn (Ca 5, Ca 6 trong tuần; Ca 7, Ca 8 cuối tuần/lễ)
5. Ca 4 trong tuần chỉ SA, chia đều ca 4 cho các thành viên SA (sai lệch 1 ca)
6. Cân bằng từng loại ca trong tuần, cuối tuần và lễ cho mỗi thành viên (sai lệch +/- 1 ca)
7. Cân bằng tổng số ca 1,2,3 trong tuần, cuối tuần (sai lệch 1 ca)
8. SA hơn non-SA khoảng 20 giờ (sai lệch +/- 1.2 giờ)
9. Cân bằng số giờ mỗi thành viên (sai lệch +/- 1.2 giờ)
10. Không cho trực ca 1,2,3 liền nhau trong tất cả các ngày trong tháng
