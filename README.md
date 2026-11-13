# BÀI TẬP LỚN XỬ LÝ ẢNH - NHÓM 1
## Đề tài: Phân đoạn đối tượng bằng hình thái học toán học và Watershed
**Học viện Công nghệ Bưu chính Viễn thông (PTIT) - Khoa Công nghệ Thông tin 1**

---

### Thành viên Nhóm 1
1. **Đinh Viết Huy (Trưởng nhóm)** - MSV: `B23DCCN381`
2. **Đỗ Trọng Ninh** - MSV: `B23DCCN633`
3. **Nguyễn Tuấn Tài** - MSV: `B23DCCN731`
4. **Cao Xuân Vũ** - MSV: `B23DCCN940`

---

## 1. Tổng quan đề tài & Mục tiêu
Bài tập lớn tập trung giải quyết bài toán cốt lõi trong phân đoạn ảnh y sinh và công nghiệp: **Phân tách các đối tượng tiếp xúc hoặc dính liền nhau** (overlapping/touching objects). Các phương pháp phân ngưỡng đơn thuần hoặc nhãn liên thông (Connected Components) sẽ bị dính chùm, trong khi thuật toán Watershed kinh điển chạy trên gradient lại gặp thảm họa quá phân đoạn (Over-segmentation).

Dự án triển khai hoàn chỉnh giải pháp **Marker-Controlled Morphological Watershed**:
- **Toán tử hình thái học toán học (Mathematical Morphology):** Co (Erosion), Giãn (Dilation), Mở (Opening), Đóng (Closing) với phần tử cấu trúc đĩa tròn (Disk SE) và Tái tạo hình thái (Geodesic Reconstruction by Dilation).
- **Biến đổi khoảng cách Euclid (Distance Transform):** Phát hiện các đỉnh cực đại địa phương (Local Maxima) định vị chính xác tâm đối tượng.
- **Watershed có điều khiển (Marker-Controlled Watershed):** Bơm nước từ các marker được lọc ngưỡng, dựng đập ngăn lũ (dams) để triệt tiêu 100% hiện tượng quá phân đoạn.
- **Đối chuẩn hệ thống chuẩn mực:** So sánh chéo với plugin **MorphoLibJ** trên nền tảng **Fiji / ImageJ**.
- **Thực nghiệm toàn diện:** Kiểm thử trên bộ dữ liệu tổng hợp chuẩn (5 seeds) và các bộ ảnh kinh điển trích xuất từ giáo trình *Digital Image Processing* (Gonzalez & Woods).

---

## 2. Cấu trúc thư mục dự án

```text
BTL/
├── src/watershed_btl/         # Mã nguồn Python chuẩn Clean Architecture
│   ├── config.py              # Cấu hình hạt giống tất định (381, 633, 731, 940, 101)
│   ├── morphology.py          # Triển khai Erosion, Dilation, Opening, Closing, Reconstruction
│   ├── pipeline.py            # Hàm segment() trọn gói và xuất ảnh các giai đoạn
│   ├── metrics.py             # Tính toán IoU, Dice, Precision, Recall, F1, Count Error
│   └── experiments.py         # Bộ kiểm định Ablation: Full vs No-Morphology vs CC
├── scripts/
│   ├── run_experiments.py     # Chạy thực nghiệm benchmark tổng hợp 5 seed
│   ├── run_real_image.py      # Chạy phân đoạn trên ảnh thực tế skimage.data.coins()
│   └── run_course_images.py   # Chạy phân đoạn trên ảnh giáo trình Gonzalez & Woods
├── fiji/
│   ├── morphological_watershed.ijm  # Macro ImageJ tự động hóa MorphoLibJ
│   └── README.md              # Hướng dẫn chạy đối chuẩn trên Fiji
├── slides/
│   ├── generate_slides.py     # Script tự động tạo 10 slide PowerPoint 16:9
│   └── BTL_Watershed_Canva.pptx # Slide thuyết trình tương thích 100% Canva
├── report/
│   ├── generate_report.py     # Script sinh báo cáo Word tự động chuẩn học thuật
│   └── Bao_cao_BTL_Watershed.docx # Báo cáo Word hoàn chỉnh (50 headings, 8 tables, 15 figures)
├── artifacts/                 # Dữ liệu ảnh các giai đoạn, metrics CSV, benchmark JSON
│   ├── run_001/               # Kết quả benchmark 5 seed tổng hợp
│   ├── real/                  # Kết quả phân đoạn ảnh coins()
│   └── course_images/         # Kết quả phân đoạn ảnh giáo trình Gonzalez & Woods
├── tests/                     # 18 unit tests tự động đảm bảo chất lượng phần mềm
├── requirements.txt           # Thư viện phụ thuộc
├── yeucau.txt                 # Đặc tả yêu cầu đề tài Bài tập lớn
└── README.md                  # Hướng dẫn chi tiết dự án
```

---

## 3. Cài đặt môi trường

Sử dụng Python 3.11+ trên Windows PowerShell:

```powershell
# Cài đặt các thư viện cần thiết
pip install -r requirements.txt
pip install python-pptx pytest
```

---

## 4. Hướng dẫn thực thi & Tái lập kết quả

### 4.1. Chạy thực nghiệm định lượng trên tập dữ liệu tổng hợp (5 Seeds)
```powershell
python scripts/run_experiments.py --output artifacts/run_001
```
*Đầu ra:* `artifacts/run_001/metrics.csv`, `benchmark.csv`, `summary.png` và các ảnh giai đoạn trong `stages/`.

### 4.2. Chạy trên tập ảnh giáo trình Gonzalez & Woods
```powershell
python scripts/run_course_images.py --output artifacts/course_images
```
*Đầu ra:* Phân đoạn các ảnh `FigP1036(blobs).tif` (89 đối tượng), `Fig1039(a)(polymersomes).tif` (13 túi tế bào), `Fig0941(a)(wood_dowels).tif` (48 tiết diện gỗ).

### 4.3. Chạy trên ảnh thực tế đồng xu (Coins)
```powershell
python scripts/run_real_image.py --output artifacts/real
```

### 4.4. Sinh báo cáo Word (.docx)
```powershell
python report/generate_report.py --artifacts artifacts/run_001 --output report/Bao_cao_BTL_Watershed.docx
```
*Quy chuẩn:* Font Times New Roman, dãn dòng 1.3, lề chuẩn PTIT (trên 3cm, dưới 3cm, trái 3cm, phải 2cm), có trang bìa, mục lục tự động, công thức toán OMML, 50 tiêu đề, 8 bảng biểu, 15 hình minh họa.

### 4.5. Sinh slide thuyết trình (.pptx)
```powershell
python slides/generate_slides.py
```
*Đặc điểm:* 10 slide thiết kế hiện đại tỷ lệ 16:9, bố cục thẻ card, bảng số liệu trực quan, chèn ảnh kết quả thực nghiệm, tương thích hoàn toàn khi import vào Canva để thuyết trình.

### 4.6. Chạy bộ kiểm thử tự động
```powershell
pytest -v
```
Toàn bộ 18 unit tests đều vượt qua (`18 passed`).

---

## 5. Kết quả thực nghiệm nổi bật

| Phương pháp | IoU (Jaccard) | Dice (F1) | Precision | Recall | Sai số đếm | Thời gian |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: |
| **Full Morphological Watershed** | **0.8812** | **0.9324** | **0.9571** | **0.9163** | **0.00** | **22.3 ms** |
| Không dùng Morphology (Watershed thuần) | 0.8415 | 0.9082 | 0.9120 | 0.9115 | +2.20 | 19.8 ms |
| Connected Components (Baseline) | 0.6842 | 0.8035 | 0.7814 | 0.8420 | -3.40 | 1.4 ms |

### Hiệu quả trên ảnh giáo trình Gonzalez & Woods:
- **FigP1036 (blobs - 89 hình tròn dính):** Connected Components chỉ nhận 46 cụm $\rightarrow$ Watershed tách trọn vẹn **89 đối tượng**.
- **Fig1039(a) (polymersomes - túi màng tế bào):** Connected Components nhận 6 cụm $\rightarrow$ Watershed nhận diện **13 túi riêng biệt**.
- **Fig0941(a) (wood dowels - bó gỗ tròn):** Connected Components nhận 30 cụm $\rightarrow$ Watershed phân định **48 tiết diện**.

---

## 6. Sản phẩm bàn giao (Deliverables)

1. **Báo cáo Word:** `report/Bao_cao_BTL_Watershed.docx` (đã copy ra `C:\Users\xytiv\Desktop\Bao_cao_BTL_Watershed.docx`)
2. **Slide thuyết trình:** `slides/BTL_Watershed_Canva.pptx` (đã copy ra `C:\Users\xytiv\Desktop\BTL_Watershed_Canva.pptx`)
3. **Mã nguồn và thực nghiệm:** Thư mục `src/`, `scripts/`, `fiji/`, `tests/`, `artifacts/`.
