
# Đơn thư/Công văn — Dashboard kết hợp (Flask 3.x) — Bản fix macOS

- Không có trang index riêng. Trang gốc `/` = **Dashboard**, hiển thị thống kê + tìm kiếm công văn.
- Đăng nhập nhanh → vào thẳng Dashboard.
- **Đã fix lỗi `hashlib.scrypt`** bằng cách dùng `pbkdf2:sha256` khi tạo mật khẩu (tương thích Python macOS).
- **Đổi port mặc định 5050** để tránh đụng AirPlay/AirTunes (403 khi dùng 5000).

## Chạy local
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .\.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt

export SECRET_KEY="something-strong"
export DATABASE_URL="sqlite:///data.db"

# chạy port 5050
python -m flask --app server run --debug --port 5050
# mở http://127.0.0.1:5050/login
```
Tài khoản mặc định: **admin/admin123** (ADMIN), **nhanvien/nhanvien123** (Nhân viên).

## Đưa lên GitHub
1) Tạo repo trống trên GitHub (ví dụ: `vanthu` dưới tài khoản của bạn).  
2) Trong thư mục dự án (nơi chứa `server/`, `requirements.txt`...):
```bash
git init
git add .
git commit -m "Initial commit: vanthu (Flask 3.x, fixed macOS scrypt & port)"
git branch -M main

# THAY <YOUR_USERNAME> và <YOUR_REPO> bằng giá trị thật, bỏ dấu < >
git remote add origin https://github.com/<YOUR_USERNAME>/<YOUR_REPO>.git
git push -u origin main
```
> Lưu ý: **Không gõ dấu `< >`** trong Terminal, đó chỉ là ký hiệu chỗ cần thay.  
> Nếu bật 2FA, khi `git push` GitHub sẽ yêu cầu mật khẩu — hãy dùng **Personal Access Token** thay vì mật khẩu.

## Deploy Render (tùy chọn)
- Build: `pip install -r requirements.txt`
- Start: `gunicorn server:app`
- Env: `SECRET_KEY`, `DATABASE_URL` (Postgres). Có thể thêm file `render.yaml` để auto tạo Postgres.
