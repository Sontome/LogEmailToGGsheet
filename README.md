# LogEmailToGGsheet
appscript lấy dữ liệu email qua api gmail đẩy lên ggsheet, tự lọc dữ liệu, triển khai api webapp,  gửi mail
- lọc email nội dung thay đổi chuyến bay > đẩy vào sheet Log > báo push sms tele
- lọc email nội dung mặt vé xuất > đẩy vào sheet Log mặt vé
- hàm check hãng bay và số mặt vé > đẩy dữ liệu vào dòng PNR - Email cần gửi tương ứng ( trigger )
- hàm check dòng PNR đủ thông tin , sẵn sàng gửi > chạy hàm xử lý đọc idmail gốc mặt vé > gọi api tương ứng hãng > tổng hợp file pdf mặt vé đã sửa > gửi email đến email đích với nội dung cài đặt trong hàm gửi mail
- sau khi gửi log ra sheet log đã gửi 
SheetName Cần thiết :
LogemailVJ
LogemailVNA
Logemailxuấtvéthànhcông
Email-SDT
Danh sách PNR cần gửi mail
LogGuiEmailPNR-KhachHang
Danh Sách BL gửi email chăm sóc
Gửi mail sau 12h xuất vé
Gửi mail trước 24h bay
Gửi mail sau 24h bay
