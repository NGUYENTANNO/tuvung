# Từ vựng

Web app (PWA) học từ vựng tiếng Anh trên điện thoại, nhập trực tiếp từ file Word (.docx) có bảng từ vựng.

**Mở app:** https://nguyentanno.github.io/tuvung/

## Tính năng
- Nhập file `.docx` có bảng các cột **Từ / Loại / Nghĩa / Ngữ cảnh** (tự nhận dòng tiêu đề, bỏ qua cột STT).
- Flashcard có lặp lại ngắt quãng (hộp Leitner 0–5: ôn lại sau 0, 1, 2, 4, 8, 16 ngày).
- Kiểm tra trắc nghiệm 4 lựa chọn: Từ → Nghĩa hoặc Nghĩa → Từ.
- Danh sách có tìm kiếm, phát âm tiếng Anh (🔊).
- Chạy offline, cài lên màn hình chính Android (Chrome → ⋮ → *Thêm vào màn hình chính*).
- Dữ liệu lưu trên máy (localStorage); có xuất/khôi phục file sao lưu.

Không dùng thư viện ngoài: file .docx được giải nén bằng `DecompressionStream` có sẵn trong trình duyệt.
