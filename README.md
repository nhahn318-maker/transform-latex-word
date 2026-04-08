# ChatGPT Formula to Word

Web app nhỏ để dán nội dung từ ChatGPT, parse công thức LaTeX thành MathML và copy sang Word với độ ổn định cao hơn.

## Chạy nhanh

1. Mở trực tiếp `index.html` bằng trình duyệt hiện đại (Edge/Chrome).
2. Hoặc chạy local server:

```powershell
cd d:\TransformLatex
python -m http.server 5500
```

Sau đó mở `http://localhost:5500`.

## Cách dùng

1. Dán nội dung vào khung bên trái.
2. Bấm `Parse + Preview`.
3. Kiểm tra/chỉnh nội dung ở khung bên phải.
4. Cách ổn định nhất: bấm `Xuất DOCX chuẩn`, mở file `.docx` trong Word.
5. Nếu cần dán nhanh: bấm `Copy sang Word`, rồi dán vào Word bằng `Ctrl + V`.
6. Nếu trình duyệt chặn clipboard HTML, dùng `Tải file mở bằng Word`.

## Hỗ trợ công thức

- Inline: `$...$`, `\(...\)`
- Block: `$$...$$`, `\[...\]`

## Lưu ý

- Công thức lỗi parse sẽ hiển thị bằng nền đỏ để bạn sửa nhanh.
- Để Word nhận công thức tốt nhất, nên dùng bản Word mới (Microsoft 365 / Word 2019+).
- Sau khi cập nhật code, bấm `Ctrl + F5` để tải lại trang, tránh dùng bản JS cũ trong cache.
- Ưu tiên dùng `Xuất DOCX chuẩn` để giữ công thức ổn định nhất trên Word.
- Lần copy/xuất đầu tiên cần mạng để tải bộ chuyển `MathML -> OMML`.
- Khi Word chỉ nhận text thuần, app sẽ tự dùng dạng tuyến tính (ví dụ `1/3`) để tránh sai thành `13`.
- Nếu Word báo lỗi mở `.docx`, hãy xóa file cũ, `Ctrl + F5` rồi xuất lại từ bản web mới nhất.
