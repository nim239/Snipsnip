![License](https://img.shields.io/badge/license-BUSL--1.1-black)




# SnipSnip (v25.09.19) by NamNhọ@visualStation
25.09.12.1135 (Phần Mềm Chống Trĩ)

---

# !!! LƯU Ý QUAN TRỌNG !!!

**Để phần mềm hoạt động, bạn BẮT BUỘC phải đặt file `ffprobe.exe` vào cùng thư mục với file `.exe` của phần mềm.**

**Để phần mềm hoạt động, bạn BẮT BUỘC phải đặt file `ffprobe.exe` vào cùng thư mục với file `.exe` của phần mềm.**

**Để phần mềm hoạt động, bạn BẮT BUỘC phải đặt file `ffprobe.exe` vào cùng thư mục với file `.exe` của phần mềm.**

---

## Giới thiệu

Đây là công cụ giúp tự động tạo file XML (tương thích với Adobe Premiere Pro, Final Cut Pro) để cắt video hàng loạt dựa trên timecode được cung cấp trong file CSV. 
- **Xuất file XML:** Tạo ra file `.xml` có thể import trực tiếp vào các phần mềm dựng phim thông dụng.

## Hướng dẫn sử dụng

**Quan trọng:** Để ứng dụng hoạt động, file `ffprobe.exe` phải được đặt trong cùng thư mục với file `.exe` của ứng dụng.

1.  **Chạy ứng dụng:** Mở file `Sniipsnip.exe`.
2.  **Chọn file CSV:** Nhấn nút `Duyệt CSV` và chọn file CSV chứa danh sách video và timecode.
3.  **Chọn thư mục Video:** Nhấn nút `Duyệt Thư mục` và chọn thư mục chứa tất cả các file video cần cắt và sắp xếp.
4.  **Chọn nơi lưu XML:** Nhấn nút `Duyệt XML` và chọn vị trí cũng như đặt tên cho file XML đầu ra.
5.  **Chạy AutoCut:** Nhấn nút `🚀 Chạy AutoCut` để bắt đầu .
6.  **Xem kết quả:**
    -   **Bảng Preview bên phải:** Sẽ được cập nhật với trạng thái của tất cả các dòng trong file CSV - trạng thái file source video.
    -   **Console Log bên trái:** Hiển thị chi tiết quá trình xử lý.
7.  Sau khi hoàn tất, file XML sẽ được tạo ở vị trí đã chọn.

## Định dạng file CSV

Tool hỗ trợ 2 định dạng chính:

**1. CSV không có Header**

Dữ liệu có 2 cột bất kì theo định dạng:
- Cột 1: Tên file video (ví dụ: `my_video.mp4, video.avi, clip.mov`)
- Cột 2: Timecode (ví dụ: `00:10 - 00:25 hoặc 00:15:12:00-00:17:10:59`)

*Ví dụ:*
```csv
P1045702.MP4,01:22 - 01:29
DJI_0081.MP4,00:00 - 00:17
```

**2. CSV có Header**

File có thể có một dòng tiêu đề ở đầu. Tool sẽ tự động tìm các cột cần thiết dựa trên từ khóa trong tên cột.

*Ví dụ:*
```csv
Video File Name,Time In - Out,Notes
P1045702.MP4,01:22 - 01:29,Cảnh đẹp
DJI_0081.MP4,00:00 - 00:17,Cần chống rung
```

---

### Liên hệ & Feedback  `tele@nimvfx` hoặc `lotus@nam001`

Cảm ơn vì dám sử dụng.
Gia đình xin cảm ơn và hậu tạ <3!!!
