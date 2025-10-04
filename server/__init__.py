# server/__init__.py
from __future__ import annotations
import os
import io
import logging
from datetime import datetime, date

from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, send_file
)
from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    LoginManager, UserMixin, login_user, login_required,
    current_user, logout_user
)
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook

# -----------------------------------------------------------------------------
# App & Config
# -----------------------------------------------------------------------------
APP_NAME = "Quản lý đơn thư đội 3"

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "dev-secret")

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///data.db")
app.config["SQLALCHEMY_DATABASE_URI"] = DATABASE_URL
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

login_mgr = LoginManager(app)
login_mgr.login_view = "login"

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

# -----------------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------------
CHUA_LAY = "NV chưa lấy đơn"
DANG_GQ = "Đang giải quyết"
HOAN_T = "Hoàn Thành"

STATUS_CHOICES = [CHUA_LAY, DANG_GQ, HOAN_T]
ROLES = ["admin", "staff"]

# -----------------------------------------------------------------------------
# Models
# -----------------------------------------------------------------------------
class User(db.Model, UserMixin):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), default="staff")

    def set_password(self, raw: str) -> None:
        self.password_hash = generate_password_hash(raw)

    def check_password(self, raw: str) -> bool:
        return check_password_hash(self.password_hash, raw)


class CongVan(db.Model):
    id = db.Column(db.Integer, primary_key=True)

    loai_don_thu = db.Column(db.String(120))  # Loại đơn thư
    thang = db.Column(db.Integer)             # Tháng (1..12)
    nam = db.Column(db.Integer)               # Năm (vd 2025)

    ma_kh = db.Column(db.String(50))          # Mã KH
    ten = db.Column(db.String(120))           # Tên KH
    dia_chi = db.Column(db.String(255))       # Địa chỉ
    nhan_vien = db.Column(db.String(80))      # Nhân viên phụ trách

    noi_dung = db.Column(db.Text)             # Nội dung đơn
    ngay_nhan = db.Column(db.Date)            # Ngày NV nhận đơn

    tinh_trang = db.Column(db.String(40))     # Trạng thái xử lý
    ghi_chu = db.Column(db.String(255))       # Ghi chú

# -----------------------------------------------------------------------------
# DB init & seed
# -----------------------------------------------------------------------------
def seed_users():
    """Tạo admin/staff mặc định nếu DB trống (an toàn, không ghi đè)."""
    if User.query.count() > 0:
        return
    admin_u = os.getenv("DEFAULT_ADMIN_USERNAME", "admin")
    admin_p = os.getenv("DEFAULT_ADMIN_PASSWORD", "admin123")
    staff_u = os.getenv("DEFAULT_STAFF_USERNAME", "nhanvien")
    staff_p = os.getenv("DEFAULT_STAFF_PASSWORD", "nhanvien123")

    admin = User(username=admin_u, role="admin")
    admin.set_password(admin_p)
    staff = User(username=staff_u, role="staff")
    staff.set_password(staff_p)
    db.session.add_all([admin, staff])
    db.session.commit()
    logger.info("Seeded default users: %s/%s, %s/%s", admin_u, "******", staff_u, "******")


with app.app_context():
    db.create_all()
    seed_users()

# -----------------------------------------------------------------------------
# Login
# -----------------------------------------------------------------------------
@login_mgr.user_loader
def load_user(user_id: str):
    return User.query.get(int(user_id))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password):
            login_user(user)
            return redirect(url_for("dashboard"))
        flash("Sai tài khoản hoặc mật khẩu", "error")
    return render_template("login.html", APP_NAME=APP_NAME)


@app.get("/quick")
def quick_login():
    """Đăng nhập nhanh với quyền staff (chỉ xem)."""
    u = User.query.filter_by(role="staff").first()
    if not u:
        flash("Chưa có tài khoản staff", "error")
        return redirect(url_for("login"))
    login_user(u)
    return redirect(url_for("dashboard"))


@app.get("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def _int(val, default=None):
    try:
        return int(val)
    except Exception:
        return default

def _date(value: str | None) -> date | None:
    if not value:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except Exception:
            pass
    return None

def apply_filters(q):
    """Áp dụng bộ lọc theo querystring."""
    if request.args.get("ma"):
        q = q.filter(CongVan.ma_kh.ilike(f"%{request.args.get('ma').strip()}%"))
    if request.args.get("ten"):
        q = q.filter(CongVan.ten.ilike(f"%{request.args.get('ten').strip()}%"))
    if request.args.get("dia_chi"):
        q = q.filter(CongVan.dia_chi.ilike(f"%{request.args.get('dia_chi').strip()}%"))
    if request.args.get("loai"):
        q = q.filter(CongVan.loai_don_thu == request.args.get("loai"))
    if request.args.get("q"):
        k = f"%{request.args.get('q').strip()}%"
        q = q.filter(
            db.or_(
                CongVan.noi_dung.ilike(k),
                CongVan.ghi_chu.ilike(k),
            )
        )
    thang = _int(request.args.get("thang"))
    if thang:
        q = q.filter(CongVan.thang == thang)
    nam = _int(request.args.get("nam"))
    if nam:
        q = q.filter(CongVan.nam == nam)
    if request.args.get("tinh_trang"):
        q = q.filter(CongVan.tinh_trang == request.args.get("tinh_trang"))
    return q

# -----------------------------------------------------------------------------
# Context processor – an toàn (chống 500 nếu DB “ẩm ương”)
# -----------------------------------------------------------------------------
@app.context_processor
def _ctx():
    loai_options = []
    try:
        rows = (
            db.session.query(CongVan.loai_don_thu)
            .filter(CongVan.loai_don_thu.isnot(None))
            .distinct()
            .all()
        )
        loai_options = [r[0] for r in rows if (r and r[0])]
    except Exception:
        loai_options = []

    def page_url(p):
        args = request.args.to_dict(flat=True)
        args["page"] = p
        return url_for("dashboard", **args)

    return dict(
        APP_NAME=APP_NAME,
        STATUS_CHOICES=STATUS_CHOICES,
        loai_options=loai_options,
        page_url=page_url,
    )

# -----------------------------------------------------------------------------
# Dashboard – chống 500 toàn diện
# -----------------------------------------------------------------------------
@app.get("/")
@login_required
def dashboard():
    # Tải nhanh tất cả để thống kê (kèm guard)
    try:
        all_rows = CongVan.query.order_by(CongVan.id.desc()).all()
    except Exception:
        all_rows = []

    chua_lay = [r for r in all_rows if (r.tinh_trang or "") == CHUA_LAY]
    dang_giai_quyet = [r for r in all_rows if (r.tinh_trang or "") == DANG_GQ]

    # Hoàn thành theo loại – tháng gần nhất
    completed = [r for r in all_rows if (r.tinh_trang or "") == HOAN_T and r.thang and r.nam]
    latest_thang = latest_nam = None
    hoan_thanh_by_loai = []
    month_total = month_completed = 0
    if completed:
        try:
            latest_nam, latest_thang = max(((int(r.nam), int(r.thang)) for r in completed))
            month_rows = [r for r in all_rows if r.nam == latest_nam and r.thang == latest_thang]
            month_total = len(month_rows)
            month_completed = sum(1 for r in month_rows if (r.tinh_trang or "") == HOAN_T)
            counter = {}
            for r in month_rows:
                if (r.tinh_trang or "") == HOAN_T:
                    key = (r.loai_don_thu or "(không ghi)")
                    counter[key] = counter.get(key, 0) + 1
            hoan_thanh_by_loai = sorted(counter.items(), key=lambda x: x[0])
        except Exception:
            latest_thang = latest_nam = None
            hoan_thanh_by_loai = []
            month_total = month_completed = 0

    # Bảng chi tiết có filter + paginate (guard)
    try:
        q = apply_filters(CongVan.query.order_by(CongVan.id.desc()))
        page = _int(request.args.get("page"), 1)
        per_page = 20
        pager = q.paginate(page=page, per_page=per_page, error_out=False)
        items = pager.items
    except Exception:
        class _P:
            page = 1; pages = 1; has_prev = False; has_next = False
            prev_num = 1; next_num = 1
        pager = _P()
        items = []

    return render_template(
        "dashboard.html",
        chua_lay=chua_lay,
        dang_giai_quyet=dang_giai_quyet,
        latest_thang=latest_thang,
        latest_nam=latest_nam,
        hoan_thanh_by_loai=hoan_thanh_by_loai,
        month_total=month_total,
        month_completed=month_completed,
        items=items,
        pager=pager,
    )

# -----------------------------------------------------------------------------
# Thêm / Sửa / Xem công văn
# -----------------------------------------------------------------------------
@app.route("/cong-van/add", methods=["GET", "POST"])
@login_required
def congvan_add():
    if request.method == "POST":
        f = request.form
        cv = CongVan(
            loai_don_thu=(f.get("loai_don_thu") or "").strip() or None,
            thang=_int(f.get("thang")),
            nam=_int(f.get("nam")),
            ma_kh=(f.get("ma_kh") or "").strip() or None,
            ten=(f.get("ten") or "").strip() or None,
            dia_chi=(f.get("dia_chi") or "").strip() or None,
            nhan_vien=(f.get("nhan_vien") or "").strip() or None,
            noi_dung=(f.get("noi_dung") or "").strip() or None,
            ngay_nhan=_date(f.get("ngay_nhan")) or date.today(),
            tinh_trang=f.get("tinh_trang") if f.get("tinh_trang") in STATUS_CHOICES else CHUA_LAY,
            ghi_chu=(f.get("ghi_chu") or "").strip() or None,
        )
        db.session.add(cv)
        db.session.commit()
        flash("Đã thêm công văn", "success")
        return redirect(url_for("dashboard"))

    # GET
    today = date.today().strftime("%Y-%m-%d")
    return render_template("congvan_form.html", today=today, STATUS_CHOICES=STATUS_CHOICES)


@app.get("/cong-van/<int:cid>")
@login_required
def congvan_detail(cid: int):
    cv = CongVan.query.get_or_404(cid)
    return render_template("congvan_detail.html", r=cv)


@app.route("/cong-van/<int:cid>/edit", methods=["GET", "POST"])
@login_required
def congvan_edit(cid: int):
    cv = CongVan.query.get_or_404(cid)
    if request.method == "POST":
        f = request.form
        cv.loai_don_thu = (f.get("loai_don_thu") or "").strip() or None
        cv.thang = _int(f.get("thang"))
        cv.nam = _int(f.get("nam"))
        cv.ma_kh = (f.get("ma_kh") or "").strip() or None
        cv.ten = (f.get("ten") or "").strip() or None
        cv.dia_chi = (f.get("dia_chi") or "").strip() or None
        cv.nhan_vien = (f.get("nhan_vien") or "").strip() or None
        cv.noi_dung = (f.get("noi_dung") or "").strip() or None
        cv.ngay_nhan = _date(f.get("ngay_nhan")) or cv.ngay_nhan
        cv.tinh_trang = f.get("tinh_trang") if f.get("tinh_trang") in STATUS_CHOICES else cv.tinh_trang
        cv.ghi_chu = (f.get("ghi_chu") or "").strip() or None
        db.session.commit()
        flash("Đã cập nhật", "success")
        return redirect(url_for("congvan_detail", cid=cv.id))

    today = (cv.ngay_nhan or date.today()).strftime("%Y-%m-%d")
    return render_template("congvan_form.html", r=cv, today=today, STATUS_CHOICES=STATUS_CHOICES)


@app.post("/cong-van/<int:cid>/delete")
@login_required
def congvan_delete(cid: int):
    cv = CongVan.query.get_or_404(cid)
    db.session.delete(cv)
    db.session.commit()
    flash("Đã xóa công văn", "success")
    return redirect(url_for("dashboard"))

# -----------------------------------------------------------------------------
# Xuất Excel (áp dụng bộ lọc hiện tại)
# -----------------------------------------------------------------------------
@app.get("/export-excel")
@login_required
def export_excel():
    q = apply_filters(CongVan.query.order_by(CongVan.id.desc()))
    rows = q.all()

    wb = Workbook()
    ws = wb.active
    ws.title = "CongVan"

    headers = [
        "Mã", "Tháng", "Năm", "Loại đơn thư", "Mã KH", "Tên", "Địa chỉ",
        "Nội dung", "Nhân viên", "Ngày nhận", "Tình trạng", "Ghi chú"
    ]
    ws.append(headers)

    def _v(s): return "" if s is None else s

    for r in rows:
        ws.append([
            r.id, r.thang, r.nam, _v(r.loai_don_thu), _v(r.ma_kh), _v(r.ten),
            _v(r.dia_chi), _v(r.noi_dung), _v(r.nhan_vien),
            (r.ngay_nhan.strftime("%d/%m/%Y") if r.ngay_nhan else ""),
            _v(r.tinh_trang), _v(r.ghi_chu)
        ])

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    fname = f"congvan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        bio, as_attachment=True, download_name=fname,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -----------------------------------------------------------------------------
# Quản lý người dùng (admin)
# -----------------------------------------------------------------------------
def _admin_required():
    if not current_user.is_authenticated or current_user.role != "admin":
        flash("Chỉ admin được phép truy cập", "error")
        return False
    return True

@app.get("/users")
@login_required
def users():
    if not _admin_required():
        return redirect(url_for("dashboard"))
    return render_template("users.html", users=User.query.order_by(User.id).all(), ROLES=ROLES)

@app.post("/users")
@login_required
def users_create():
    if not _admin_required():
        return redirect(url_for("dashboard"))
    u = (request.form.get("username") or "").strip()
    p = request.form.get("password") or ""
    role = request.form.get("role") if request.form.get("role") in ROLES else "staff"
    if not u or not p:
        flash("Thiếu username/mật khẩu", "error")
        return redirect(url_for("users"))
    if User.query.filter_by(username=u).first():
        flash("Username đã tồn tại", "error")
        return redirect(url_for("users"))
    user = User(username=u, role=role)
    user.set_password(p)
    db.session.add(user)
    db.session.commit()
    flash("Đã tạo tài khoản", "success")
    return redirect(url_for("users"))

@app.post("/users/<int:uid>/reset")
@login_required
def users_reset(uid: int):
    if not _admin_required():
        return redirect(url_for("dashboard"))
    user = User.query.get_or_404(uid)
    p = request.form.get("password") or ""
    if not p:
        flash("Chưa nhập mật khẩu mới", "error")
        return redirect(url_for("users"))
    user.set_password(p)
    db.session.commit()
    flash("Đã đặt lại mật khẩu", "success")
    return redirect(url_for("users"))

@app.post("/users/<int:uid>/update")
@login_required
def users_update(uid: int):
    if not _admin_required():
        return redirect(url_for("dashboard"))
    user = User.query.get_or_404(uid)
    role = request.form.get("role")
    name = (request.form.get("username") or "").strip()
    if role in ROLES:
        user.role = role
    if name:
        if User.query.filter(User.username == name, User.id != uid).first():
            flash("Username đã có", "error")
            return redirect(url_for("users"))
        user.username = name
    db.session.commit()
    flash("Đã cập nhật người dùng", "success")
    return redirect(url_for("users"))

@app.post("/users/<int:uid>/delete")
@login_required
def users_delete(uid: int):
    if not _admin_required():
        return redirect(url_for("dashboard"))
    user = User.query.get_or_404(uid)
    if user.id == current_user.id:
        flash("Không thể tự xóa chính mình", "error")
        return redirect(url_for("users"))
    db.session.delete(user)
    db.session.commit()
    flash("Đã xóa tài khoản", "success")
    return redirect(url_for("users"))

# -----------------------------------------------------------------------------
# Error pages (ghi log để thấy trên Render)
# -----------------------------------------------------------------------------
@app.errorhandler(404)
def _404(e):
    return render_template("error.html", code=404, msg="Không tìm thấy"), 404

@app.errorhandler(500)
def _500(e):
    try:
        logger.exception("Internal Server Error: %s", e)
    except Exception:
        pass
    return render_template("error.html", code=500, msg="Lỗi máy chủ"), 500

# -----------------------------------------------------------------------------
# Healthcheck
# -----------------------------------------------------------------------------
@app.get("/healthz")
def healthz():
    return "ok", 200

# Cho gunicorn: `server:app`
__all__ = ["app", "db", "User", "CongVan"]
