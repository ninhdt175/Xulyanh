# Thiết kế bài tập lớn: Phân đoạn đối tượng bằng hình thái học toán học và Watershed

## Mục tiêu

Xây dựng một gói bài tập lớn có thể tái lập cho Nhóm 1, chứng minh quy trình phân đoạn đối tượng dựa trên hình thái học toán học và marker-controlled watershed. Gói phải dùng Fiji/ImageJ với MorphoLibJ làm nền tảng tham chiếu, có cài đặt Python độc lập, và đánh giá bằng chỉ số định lượng thay vì chỉ quan sát.

## Phạm vi và đầu ra

Thư mục dự án sẽ có bốn đầu ra chính:

1. Chương trình Python chạy từ dòng lệnh, tạo dữ liệu tổng hợp có nhãn chuẩn, phân đoạn ảnh tổng hợp và ảnh thực tế, xuất nhãn, ảnh từng bước, metric và biểu đồ.
2. ImageJ Macro cùng hướng dẫn workflow MorphoLibJ để chạy quy trình tương đương trong Fiji và xuất nhãn để đối chiếu.
3. Báo cáo Word khoảng 30 trang, được sinh từ dữ liệu thực nghiệm của mã.
4. Bản trình bày khoảng 10 slide dưới dạng PPTX, có thể nhập vào Canva để tiếp tục chỉnh sửa.

Không bao gồm việc điền logo trường, khoa, tên học phần, giảng viên hoặc thời hạn trên bìa; các vị trí này được để trống theo yêu cầu.

## Dữ liệu và tính tái lập

Quy trình dùng hai loại dữ liệu:

- **Tổng hợp:** các đĩa/ellipse chồng lấn, có nhiễu và độ sáng nền thay đổi. Mask instance ground truth được tạo cùng ảnh, nên IoU, Dice, precision, recall và số lượng đối tượng có ý nghĩa rõ ràng.
- **Thực tế:** một ảnh coin/cell kích thước nhỏ, công khai và được đóng gói/tải theo script. Ảnh này dùng để đánh giá trực quan và kiểm tra khả năng tổng quát; metric chính trên ảnh thực tế là mức đồng thuận nhãn Python–Fiji khi sử dụng cùng ngưỡng và marker.

Mọi tham số, seed sinh dữ liệu, đường dẫn đầu vào/đầu ra và phiên bản thư viện được ghi trong cấu hình và README. Không đưa kết quả tự nhận là đo từ Fiji nếu Fiji/MorphoLibJ không được cài và chạy trên máy; báo cáo sẽ ghi rõ cách xuất số liệu Fiji và khu vực bảng kết quả tương ứng.

## Thuật toán Python

Pipeline xử lý ảnh xám gồm: chuẩn hoá và khử nhiễu nhẹ; threshold tạo foreground; opening/closing loại nhiễu và lấp lỗ; distance transform; ngưỡng distance để tạo sure foreground; connected-components tạo marker; marker-controlled watershed trên ảnh gradient/distance âm; hậu xử lý và gán nhãn instance.

Các phép erosion, dilation, opening, closing, morphological reconstruction, distance transform, connected components và watershed được hiện thực/đóng gói thành các hàm nhỏ có đầu vào/đầu ra xác định. Thư viện dùng NumPy, SciPy, scikit-image, OpenCV (nếu cần), pandas và matplotlib. Các bước có thể chọn kernel, số iteration, threshold và khoảng cách marker bằng CLI/cấu hình.

Chương trình xuất: ảnh đầu vào, mask nhị phân, distance map, marker, nhãn watershed, overlay; CSV metric; CSV benchmark thời gian; và PNG biểu đồ so sánh.

## Đối chiếu Fiji/MorphoLibJ

Tài liệu sẽ mô tả chuỗi thao tác Fiji tương ứng: mở ảnh, threshold, Morphological Filters, Distance Map, Connected Components Labeling, Marker-controlled Watershed/Watershed và đo kết quả. Macro chạy các thao tác có thể tự động hoá, lưu ảnh nhãn và thông số vào thư mục kết quả. Hướng dẫn chỉ rõ menu/plugin cần có, các tham số tương ứng Python, cũng như cách đo diện tích/số lượng đối tượng trong Fiji.

## Đánh giá

Trên dữ liệu tổng hợp, chương trình so sánh mask và instance prediction với ground truth qua IoU, Dice, precision, recall, F1, sai số số lượng đối tượng và thời gian chạy. Có ablation tối thiểu so sánh watershed có/không có bước morphology hoặc marker để chỉ ra vai trò của từng bước.

Trên ảnh thực tế, báo cáo so sánh trực quan và đo agreement IoU/Dice giữa output nhị phân Python với output Fiji sau khi căn chỉnh kích thước và quy ước foreground. Bảng Fiji có chỗ nhập kết quả khi người dùng chạy macro; bảng Python được tạo trực tiếp từ mã. Kết luận phân biệt rõ kết quả đo tự động và nhận xét định tính.

## Báo cáo Word

Tệp DOCX dùng Times New Roman 14 pt. Bìa không mang số trang và chừa vùng thông tin trường/khoa/học phần/giảng viên. Từ trang mục lục trở đi, số trang bắt đầu tại 1, căn giữa chân trang. Phần nội dung dùng lề trên 3 cm, dưới 3 cm, trái 3 cm, phải 2 cm.

Bố cục: bìa; mục lục; danh mục hình/bảng; mở đầu; cơ sở lý thuyết; Fiji và MorphoLibJ; thiết kế/cài đặt Python; dữ liệu và phương pháp thí nghiệm; kết quả/đánh giá; kết luận/hướng phát triển; tài liệu tham khảo; phụ lục hướng dẫn chạy và macro. Hình, bảng và kết quả phải được lấy từ artifact của pipeline hoặc được chú thích rõ là sơ đồ minh hoạ.

## Slide Canva-compatible

PPTX gồm khoảng 10 slide: vấn đề/mục tiêu, trực giác morphology, watershed, workflow Fiji, kiến trúc Python, dataset/thiết lập, kết quả định lượng, kết quả trực quan, so sánh/thảo luận, kết luận. Slide dùng ít chữ, có sơ đồ pipeline, hình output và biểu đồ từ Python; font/đối tượng ở dạng có thể sửa sau khi import Canva.

## Kiểm thử và nghiệm thu

Kiểm thử đơn vị cho morphology/metric trên mảng nhỏ có kết quả biết trước, cùng smoke test chạy toàn pipeline với seed cố định. Kiểm tra output tồn tại và metric nằm trong giới hạn hợp lệ. DOCX được render kiểm tra trực quan: bìa không số trang, mục lục mang số 1, font/lề đúng trên nội dung. PPTX được render kiểm tra tràn chữ, thiếu ảnh và tính nhất quán.

## Rủi ro và xử lý lỗi

Thiếu Fiji hoặc MorphoLibJ không làm hỏng pipeline Python; README nêu cách cài và phần kết quả Fiji để tái tạo. Khác biệt pixel giữa hai nền tảng được kiểm soát bằng cùng ảnh đầu vào, quy ước foreground và marker, sau đó ghi nhận là khác biệt cài đặt nếu còn. Các download dữ liệu thất bại có đường dẫn fallback là ảnh tổng hợp đóng gói sẵn.
