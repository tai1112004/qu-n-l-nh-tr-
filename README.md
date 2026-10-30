# HỆ THỐNG QUẢN LÝ VÀ CHO THUÊ NHÀ TRỌ (FULL-STACK DOCKERIZED)

Dự án website quản lý và cho thuê phòng trọ hoàn chỉnh bao gồm Frontend (ReactJS), Backend (Spring Boot 2.7.14) và Cơ sở dữ liệu (MySQL 8.0). Toàn bộ hệ thống đã được đóng gói container hóa (Docker & Docker Compose), người dùng chỉ cần clone repository về máy là có thể chạy ngay với một câu lệnh duy nhất.

---

## 🚀 HƯỚNG DẪN KHỞI CHẠY BẰNG DOCKER

### 1. Yêu cầu môi trường
- Máy tính đã cài đặt **[Docker Desktop](https://www.docker.com/products/docker-desktop/)** (Bao gồm Docker Engine và Docker Compose).

### 2. Khởi chạy hệ thống
Mở terminal (PowerShell, Command Prompt hoặc Git Bash) tại thư mục gốc của dự án và chạy câu lệnh:

```bash
docker compose up -d --build
```

Hệ thống sẽ tự động:
1. Kéo image MySQL 8.0 và khởi tạo cơ sở dữ liệu `rental_home`.
2. Build và đóng gói ứng dụng Backend Spring Boot (Java 17).
3. Build và đóng gói ứng dụng Frontend ReactJS được phục vụ qua Nginx.
4. Tự động kiểm tra kết nối và khởi tạo tài khoản Quản trị viên (Admin) mặc định.

---

## 🔑 THÔNG TIN TÀI KHOẢN VÀ CỔNG TRUY CẬP

### 🌐 Địa chỉ truy cập:
- **Giao diện người dùng (Frontend):** [http://localhost:3000](http://localhost:3000)
- **Đăng nhập Quản trị viên (Admin):** [http://localhost:3000/login-admin](http://localhost:3000/login-admin)
- **Đăng nhập Chủ trọ (Rentaler):** [http://localhost:3000/login-rentaler](http://localhost:3000/login-rentaler)
- **Backend REST API:** [http://localhost:8081](http://localhost:8081)
- **Cơ sở dữ liệu MySQL:** `localhost:3306` (Database: `rental_home`, User: `root`, Password: `root`)

### 👑 Tài khoản Quản trị viên (Admin) mặc định:
- **Email / Tên đăng nhập:** `admin@gmail.com` *(hoặc nhập `admin`)*
- **Mật khẩu:** `admin@123`
- **Quyền hạn:** `ROLE_ADMIN` (Kiểm duyệt tin đăng, khóa tài khoản, phân quyền, gửi email thông báo).

---

## 🛠️ CÁC LỆNH QUẢN LÝ DOCKER HỮU ÍCH

- **Xem log hoạt động thời gian thực của hệ thống:**
  ```bash
  docker compose logs -f
  ```
- **Xem riêng log của Backend:**
  ```bash
  docker compose logs -f backend
  ```
- **Dừng hệ thống:**
  ```bash
  docker compose down
  ```
- **Dừng và xóa toàn bộ dữ liệu database (Reset DB):**
  ```bash
  docker compose down -v
  ```
- **Khởi động lại các container:**
  ```bash
  docker compose restart
  ```
