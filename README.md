### Công cụ dịch thuật cá nhân
Một công cụ được tạo ra cho công việc dịch thuật và bản địa hóa cá nhân của tôi. Công cụ này hoạt động dựa trên ngữ cảnh và được hỗ trợ bởi trí tuệ nhân tạo.
"""
## TransPal — Trợ lý dịch thuật
========================================================================================
### Cài đặt (chỉ 1 lần, mở CMD): 
pip cài đặt bàn phím pyperclip pyautogui google-genai deep-translator python-docx pypdf
Đặt khóa API (chỉ 1 lần đầu): 
setx GEMINI_API_KEY "key_cua_ban" 
(đóng mở lại ứng dụng sau khi setx)
Chạy: click đúp vào file transpal_gui.pyw là xong, không cần terminal.
"""
### ===== Hướng dẫn =========================================================================
##Dịch văn bản:

Mở bất kỳ ứng dụng nào (trình duyệt, Word, phần mềm trò chuyện...), bôi đen đoạn cần dịch
Bấm Alt+Shift+T rồi xả tay ra
Lúc này có 3 lựa chọn ở cửa sổ bật lên số:
Chỉnh sửa bản dịch trực tiếp trong ô nếu chưa tương ứng, rồi nhấn
"Thay thế" (hoặc Enter) → đoạn bôi đen sẽ được dán bằng bản dịch
"Chỉ sao chép" → bản dịch nằm trong clipboard, tự dán
"Đóng" (hoặc Esc) → bỏ qua, không làm gì

##Các phím tắt khác:

Ctrl+Alt+M: thay đổi nhanh giữa AI chế độ (chậm hơn nhưng hiểu ngữ cảnh) và dịch cơ bản Google (nhanh, không cần thiết chìa khóa)
Ctrl+Alt+X: tắt tập lệnh

Về tập tin cảnh — đây là phần quan trọng nhất để AI dịch đúng ý
Khi chạy lần đầu, tập lệnh đã tự động tạo tệp context.md nằm cùng thư mục với transpal.py. Mở nó bằng Notepad (hoặc VS Code) được chỉnh sửa, ví dụ bạn có thể viết:
markdown# Thuật toán hướng dẫn
- Dịch tự nhiên, văn phong Việt Nam, Khoa hô "bạn".
- Đây là tài liệu về game MMORPG, chứa đựng các từ: buff, debuff, raid, dungeon.
- "Guild" dịch là "bang hội", "quest" dịch là "nhiệm vụ".
- Tên nhân vật được đặt riêng, không có phiên bản âm thanh.
- Giọng văn trẻ trung, không quá trang trọng.
Một số điểm cần biết:

Edit context.md không cần khởi động lại tập lệnh — tệp có thể được đọc lại mỗi lần bạn nhấn phím dịch, nên chỉnh sửa xong việc lưu lại (Ctrl+S) là lần dịch tiếp theo được sử dụng ngay.
Viết những gì cũng được: quy tắc dịch, bảng thuật ngữ, mô tả bối cảnh tài liệu ("đây là hợp đồng pháp", "đây là truyện tiên hiệp")... AI đọc và làm theo.
Nếu có hướng dẫn tệp sẵn ở dạng .txt thì sao chép nội dung phong cách vào context.md hoặc nhập xong. Còn nếu là PDF/Word thì hiện tại cần sao chép văn bản thủ công
Lưu ý chế độ AI chỉ áp dụng ngữ cảnh; cơ bản dịch chế độ (Google) thì không đọc được tệp này vì nó chỉ là máy dịch tĩnh.
*Đối với các mô hình AI, có thể điều chỉnh mã để lấy lại tốc độ và độ chính xác tối đa, công cụ này đang sử dụng các mô hình miễn phí, tốc độ ưu tiên và RPM,RPD cao*
