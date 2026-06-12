# 📝 TransPal — Trợ lý dịch thuật

> Công cụ dịch thuật và bản địa hóa cá nhân được hỗ trợ bởi AI và ngữ cảnh tùy chỉnh.

---

## ⚙️ Cài đặt (chỉ thực hiện một lần)

### Cài đặt thư viện

```bash
pip install keyboard pyperclip pyautogui google-genai deep-translator python-docx pypdf
```

### Thiết lập API

```bash
setx GEMINI_API_KEY "YOUR_API_KEY"
```

> Sau khi thiết lập API, hãy đóng và mở lại ứng dụng.

### Khởi chạy

Mở trực tiếp:

```text
transpal_gui.pyw
```

Không cần mở Terminal hoặc CMD.

---

## 📖 Hướng dẫn

### Dịch văn bản

1. Mở bất kỳ ứng dụng nào (trình duyệt, Word, phần mềm trò chuyện,...)
2. Bôi đen đoạn văn bản cần dịch.
3. Nhấn `Alt + Shift + T`.
4. Cửa sổ dịch sẽ xuất hiện với 3 lựa chọn:

- **Thay thế** (hoặc Enter): thay thế nội dung đã chọn bằng bản dịch.
- **Chỉ sao chép**: lưu bản dịch vào Clipboard.
- **Đóng** (hoặc Esc): hủy thao tác.

---

## ⌨️ Các phím tắt

| Phím tắt | Chức năng |
|-----------|-----------|
| Alt + Shift + T | Dịch văn bản đã chọn |
| Ctrl + Alt + M | Chuyển đổi AI / Google Translate |
| Ctrl + Alt + X | Thoát chương trình |

---

## 🧠 Về tệp ngữ cảnh (context.md)

Đây là thành phần quan trọng nhất để AI hiểu đúng nội dung cần dịch.

Khi chạy lần đầu, chương trình sẽ tự động tạo:

```text
context.md
```

trong cùng thư mục với TransPal.

Bạn có thể chỉnh sửa bằng:

- Notepad
- VS Code
- Notepad++
- Hoặc bất kỳ trình soạn thảo văn bản nào

Ví dụ:

```md
# Hướng dẫn dịch

- Dịch tự nhiên, văn phong Việt Nam.
- Xưng hô là "bạn".
- Đây là tài liệu game MMORPG.
- Guild → Bang hội
- Quest → Nhiệm vụ
- Giữ nguyên tên nhân vật.
```
*Ngoài ra có thể import trực tiếp trong giao diện sử dụng, ai hỗ trợ tách, chuyển đổi, merged với quy tắc context có sẵn*
### Một số điểm cần biết

- Không cần khởi động lại ứng dụng sau khi sửa `context.md`.
- Chỉ cần lưu tệp (`Ctrl + S`).
- Lần dịch tiếp theo sẽ tự động sử dụng nội dung mới.

Bạn có thể thêm:

- Quy tắc dịch
- Glossary
- Translation Memory
- Mô tả dự án
- Hướng dẫn văn phong
- Thuật ngữ chuyên ngành

> Lưu ý: Chế độ Google Translate không sử dụng nội dung từ `context.md`.
> Chỉ chế độ AI mới áp dụng ngữ cảnh và quy tắc dịch.

---

## 💡 Ghi chú

Đối với các mô hình AI, có thể tùy chỉnh mã nguồn để cân bằng giữa tốc độ, độ chính xác, giới hạn RPM/RPD và chi phí sử dụng.

Phiên bản hiện tại ưu tiên các mô hình miễn phí có tốc độ phản hồi cao và hạn mức sử dụng lớn.
