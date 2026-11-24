# PRJ-35 · Nhóm 26 — Hướng dẫn tái lập

Báo cáo PDF ở cạnh `Workspace_PRJ35`. PDF nền được xuất từ Word, có mục lục và danh mục bảng/hình. Bản ngày 22/11/2026 đã đính chính trực tiếp vùng văn bản ở trang 19, 64, 67 và 68 bằng PyMuPDF, giữ nguyên 68 trang và số trang mục lục; không có bản Word đồng bộ trong gói (1). Mã nguồn đầy đủ nộp riêng theo hướng dẫn giảng viên đã được sinh viên xác nhận.

## 1. Môi trường đã kiểm tra

Windows, Python 3.12.10. Các phiên bản chính: NumPy 2.5.2, pandas 3.0.5, PyArrow 26.0.1, SciPy 1.18.1, scikit-learn 1.9.1. `requirements.txt` ghim thư viện chính; `requirements-lock.txt` ghim cả phụ thuộc của môi trường kiểm tra.

**Đặt môi trường ảo ở đường dẫn không dấu.** Lần thử đặt môi trường trong thư mục có dấu đã gặp lỗi `threadpoolctl` khi nạp DLL để đo hiệu năng. Dữ liệu và mã nguồn vẫn có thể ở thư mục tiếng Việt. Hướng dẫn dưới đây dùng LocalAppData; nếu tên tài khoản cũng có dấu, chọn một đường dẫn không dấu khác cho môi trường.

Mở PowerShell tại `02_MÃ_NGUỒN_CODE`, chạy:

```powershell
$prj35Env = Join-Path $env:LOCALAPPDATA 'PRJ35_venv'
py -3.12 -m venv $prj35Env
$prj35Python = Join-Path $prj35Env 'Scripts/python.exe'
& $prj35Python -m pip install -r requirements-lock.txt
& $prj35Python -m pip check
& $prj35Python chay_toan_bo.py
```

Không cần kích hoạt PowerShell script. `chay_toan_bo.py` tự xác định thư mục mã nguồn, không chứa đường dẫn tài khoản người dùng. Lần chạy có thể mất vài phút và ghi lại các tệp kết quả; nên chạy trên bản sao khi muốn giữ nguyên số đo hiệu năng của bản nộp.

## 2. Thứ tự chạy và đầu vào

1. `TOAN_BO_MA_NGUON_PRJ35.py`: phân tích chính và đối chiếu B1. Đầu vào gồm hai CSV thô trong `data/nguon_goc/` và `data/power_clean_csv1.csv` (bảng sạch bàn giao). B1 tái lập gần khớp bảy cột, không xác nhận nguồn đo thực.
2. `bo_sung_thuc_nghiem.py`: đọc Parquet do bước 1 tạo; chạy mô phỏng fallback, độ nhạy train/validation và ví dụ làm sạch. Quy tắc được lưu trong `quy_tac_bo_sung.json`.
3. `cap_nhat_dashboard.py`: nhúng lại các CSV vào HTML để mở ngoại tuyến; cập nhật cả bảng bổ sung.

Notebook `tieu_luan_prj35.ipynb` gồm 74 ô code của phần chính. Để chạy notebook với đúng môi trường, chọn kernel Python ở môi trường vừa cài, đặt thư mục làm việc tại thư mục mã nguồn rồi Restart Kernel and Run All. Phần bổ sung nằm ở tệp Python riêng, không được hiểu là đã có trong 74 ô cũ. Notebook B1 trong `ho_so_du_lieu/` chạy với thư mục làm việc tại đó.

## 3. Đầu ra và số đối chiếu

- Bảng ghép `data/merged_minute_float32_B3.parquet`: 2.075.259 dòng × 18 cột.
- Bộ chính: train/validation/test = 331/68/66 ngày.
- RF MAE validation ≈ 0,1837 kW; OLS và RF MAE test ≈ 0,1550 và 0,1698 kW.
- Fallback hồi cứu trên test: RF đủ đầu vào 83/289 ngày; hệ thống kết hợp đủ đầu vào 243/289 ngày; chấm được 205 ngày. Các sai số xem trong `fallback_tong_hop.csv` và đối chiếu cùng ngày trong `fallback_cung_ngay.csv`.
- Độ nhạy ngưỡng 15%/20%/25% chỉ dùng phần phát triển. Không chọn lại ngưỡng theo test.
- `dashboard_csv/` chứa 25 CSV, gồm 19 bảng gốc và 6 bảng bổ sung.
- RAM và thời gian có thể khác theo máy. Số đo là của tiến trình benchmark, không phải RAM đỉnh toàn bộ pipeline đọc hai triệu dòng.

## 4. Dashboard

Mở `03_DASHBOARD_TRỰC_QUAN/dashboard_prj35.html` bằng Edge hoặc Chrome. Có 9 mục chính và phần mở rộng “Thực nghiệm bổ sung”. Nút **Nạp CSV mới** nhận các tệp đúng tên do pipeline xuất; cập nhật cả biểu đồ và bảng bổ sung. Không cần máy chủ hay kết nối mạng để xem dữ liệu đã nhúng.

## 5. Giới hạn và minh chứng kiểm tra

Đây là thực nghiệm hồi cứu; test đã được xem trước. Thời tiết sẵn có đúng lúc là giả định. Tăng số ngày dự báo không tự chứng minh giảm sai số; phải đối chiếu trên cùng ngày có nhãn. Kết quả không xác nhận tiết kiệm điện hoặc tiền, không phải dịch vụ vận hành trực tuyến.

Thư mục `ket_qua_kiem_tra/` lưu log chạy môi trường sạch, kết quả kiểm tra phụ thuộc, bảng đối chiếu số liệu và kiểm tra dashboard. Nhật ký ở gốc workspace phân biệt thao tác chạy thực nghiệm với biên tập báo cáo. Điều kiện về đường dẫn môi trường được nêu tại mục 1.

Các output lưu sẵn trong notebook chính được giữ từ hồ sơ nguồn; số đo thời gian có thể thuộc lần chạy trước. Các CSV, dashboard và bảng benchmark trong PDF của bản nâng cấp dùng lần chạy môi trường sạch mới, có log tại `ket_qua_kiem_tra/log_pipeline.txt`. Chạy lại notebook sẽ tạo số đo thời gian mới.

## 6. Thẩm định và sửa lỗi ngày 22/11/2026

Đã sửa lỗi UnicodeEncodeError khi chuyển hướng log Windows: trình chạy thiết lập PYTHONIOENCODING=utf-8 cho các tiến trình con. Kiểm tra lại toàn pipeline trên bản sao đạt exit 0; 24 bảng kết quả khớp trong dung sai, bảng hiệu năng giữ số đo lần nộp vì thời gian/RAM thay đổi theo lần chạy. Dashboard đạt 9 mục, nạp/khôi phục CSV và không lỗi JavaScript. Minh chứng mới: `ket_qua_kiem_tra/kiem_tra_20261122/`. Mã chính không ghi đè bảng điện sạch bàn giao. PDF đã sửa mô tả công cụ, URL EIA lặp, số đo cũ ở phụ lục và bảng ánh xạ tiêu chí.

### Cải thiện cách đọc dashboard — 22/11/2026

Màn hình Validation/Test hiển thị số ngày đủ điều kiện chấm trên tổng ngày lịch, các chỉ số độ bao phủ RF và hệ thống dự phòng, kết luận theo MAE và giới hạn test hồi cứu. Các biểu đồ có chú thích về cách đọc và ứng dụng. Nút mở bảng độ bao phủ dẫn đến phần bổ sung. Tóm tắt số liệu và thứ hạng tự cập nhật khi nạp CSV; quy trình chọn RF trong tiểu luận vẫn được ghi rõ. Đã kiểm tra 9 mục, thay đổi/khôi phục CSV, màn hình rộng 390 px và khả năng giữ phần bổ sung sau khi chạy cap_nhat_dashboard.py. Không đổi dữ liệu hoặc kết quả mô hình.

## Dashboard nghiên cứu trước–sau bằng Python — hoàn tất 23/11/2026

Thư mục `04_DASHBOARD_NGHIEN_CUU/` chứa ứng dụng Streamlit mới gồm 5 phần: tổng quan, trước xử lý, quy trình, so sánh trước–sau và đánh giá. Xem `HUONG_DAN.txt` trong thư mục này để cài môi trường riêng và chạy `python -m streamlit run app.py`. Minh chứng được tính trên toàn CSV1/CSV2 thật; nhánh EDA, nhánh dự báo và thời tiết tách riêng. Có tải lên CSV/Excel, xử lý theo lựa chọn, bảo vệ train/test, xuất dữ liệu và nhật ký. Bảng trước–sau là minh chứng bổ sung tính lại từ nguồn, không sửa kết quả mô hình hoặc PDF. Dashboard HTML cũ tiếp tục phục vụ các biểu đồ mô hình Chương 4. Kiểm thử lưu tại `04_DASHBOARD_NGHIEN_CUU/kiem_tra/`.
