from datetime import datetime
import os
from io import BytesIO

from flask import (
    Flask, render_template, request, redirect, url_for, flash, session, send_file
)
from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    LoginManager, login_user, logout_user, current_user, login_required, UserMixin
)
from sqlalchemy import text
from openpyxl import Workbook

# -------------------------------------------------------
# App & DB config (sửa postgres:// -> postgresql://)
# -------------------------------------------------------
db_url = os.getenv("DATABASE_URL", "sqlite:///data.db")
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)

app = Flask(__name__)
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "dev-key")
app.config["SQLALCHEMY_DATABASE_URI"] = db_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

# -------------------------------------------------------
# Models
# -------------------------------------------------------
class User(db.Model, UserMixin):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(20), default="staff")  # 'admin' | 'staff'

class CongVan(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ma = db.Column(db.String(50))
    loai_don_thu = db.Column(db.String(120))
    thang = db.Column(db.Integer)
    nam = db.Column(db.Integer)
    ma_kh = db.Column(db.String(50))
    ten = db.Column(db.String(120))
    dia_chi = db.Column(db.String(200))
    nhan_vien = db.Column(db.String(120))
    noi_dung = db.Column(db.Text)
    ngay_nv_nhan = db.Column(db.Date)
    tinh_trang = db.Column(db.String(50))  # 'NV chưa lấy đơn', 'Đang giải quyết', 'Hoàn Thành' (có thể có giá trị cũ khác)

# -------------------------------------------------------
# Login
# -------------------------------------------------------
login_manager = LoginManager(app)
login_manager.login_view = "login"

@login_manager.user_loader
def load_user(uid):
    try:
        return User.query.get(int(uid))
    except Exception:
        return None

# -------------------------------------------------------
# UI constants
# -------------------------------------------------------
APP_NAME = "quản lý đơn thư đội 3"
STATUS_CHOICES = ["NV chưa lấy đơn", "Đang giải quyết", "Hoàn Thành"]  # bỏ 'PKD trả về NV' khỏi UI thêm mới

@app.context_processor
def inject_globals():
    return {"APP_NAME": APP_NAME, "STATUS_CHOICES": STATUS_CHOICES}

def page_url(num: int):
    args = dict(request.args)
    args["page"] = num
    return url_for("dashboard", **args)

# -------------------------------------------------------
# Health check / DB check
# -------------------------------------------------------
@app.route("/healthz")
def healthz():
    return "ok", 200

@app.route("/dbz")
def dbz():
    try:
        db.session.execute(text("SELECT 1"))
        return "db-ok", 200
    except Exception as e:
        app.logger.exception("DB check failed")
        return f"db-fail: {e}", 500

# -------------------------------------------------------
# Lazy init DB (Flask 3.1)
# -------------------------------------------------------
_init_done = False
def ensure_seed_users():
    if not User.query.first():
        admin_u = os.getenv("DEFAULT_ADMIN_USERNAME", "admin")
        admin_p = os.getenv("DEFAULT_ADMIN_PASSWORD", "admin123")
        staff_u = os.getenv("DEFAULT_STAFF_USERNAME", "nhanvien")
        staff_p = os.getenv("DEFAULT_STAFF_PASSWORD", "nhanvien123")
        db.session.add(User(username=admin_u, password=admin_p, role="admin"))
        db.session.add(User(username=staff_u, password=staff_p, role="staff"))
        db.session.commit()

@app.before_request
def _lazy_init_db_once():
    global _init_done
    if _init_done:
        return
    try:
        db.create_all()
        ensure_seed_users()
        _init_done = True
    except Exception:
        app.logger.exception("Init DB failed (ignored)")
        _init_done = True

# -------------------------------------------------------
# Auth routes
# -------------------------------------------------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        u = request.form.get("username", "").strip()
        p = request.form.get("password", "").strip()
        user = User.query.filter_by(username=u, password=p).first()
        if user:
            login_user(user)
            return redirect(url_for("dashboard"))
        flash("Sai tài khoản hoặc mật khẩu", "error")
    return render_template("login.html")

@app.route("/quick_login")
def quick_login():
    staff_u = os.getenv("DEFAULT_STAFF_USERNAME", "nhanvien")
    staff_p = os.getenv("DEFAULT_STAFF_PASSWORD", "nhanvien123")
    user = User.query.filter_by(username=staff_u, password=staff_p, role="staff").first()
    if not user:
        user = User(username=staff_u, password=staff_p, role="staff")
        db.session.add(user)
        db.session.commit()
    login_user(user)
    session["quick"] = True
    return redirect(url_for("dashboard"))

@app.route("/logout")
def logout():
    if current_user.is_authenticated:
        logout_user()
    return redirect(url_for("login"))

# -------------------------------------------------------
# Helper: build filtered query from request (reused by dashboard & export)
# -------------------------------------------------------
def build_filtered_query():
    q = CongVan.query
    ten = request.args.get("ten", "").strip()
    ma_kh = request.args.get("ma_kh", "").strip()
    thang = request.args.get("thang", "").strip()
    nam = request.args.get("nam", "").strip()
    dia_chi = request.args.get("dia_chi", "").strip()
    loai = request.args.get("loai", "").strip()
    tinh_trang = request.args.get("tinh_trang", "").strip()
    kw = request.args.get("q", "").strip()

    if ten:     q = q.filter(CongVan.ten.ilike(f"%{ten}%"))
    if ma_kh:   q = q.filter(CongVan.ma_kh.ilike(f"%{ma_kh}%"))
    if thang:
        try: q = q.filter(CongVan.thang == int(thang))
        except ValueError: pass
    if nam:
        try: q = q.filter(CongVan.nam == int(nam))
        except ValueError: pass
    if dia_chi: q = q.filter(CongVan.dia_chi.ilike(f"%{dia_chi}%"))
    if loai:
        if loai == "__EMPTY__":
            q = q.filter((CongVan.loai_don_thu == "") | (CongVan.loai_don_thu.is_(None)))
        else:
            q = q.filter(CongVan.loai_don_thu == loai)
    if tinh_trang:
        q = q.filter(CongVan.tinh_trang.ilike(f"%{tinh_trang}%"))
    if kw:
        q = q.filter(CongVan.noi_dung.ilike(f"%{kw}%"))
    return q

# -------------------------------------------------------
# Dashboard (thống kê mềm + fallback, export hint)
# -------------------------------------------------------
@app.route("/")
@login_required
def dashboard():
    items = []; pagination = None
    chua_lay = []; dang_giai_quyet = []
    latest_thang = None; latest_nam = None
    hoan_thanh_by_loai = []; loai_options = []
    month_total = 0; month_completed = 0

    try:
        q = build_filtered_query()

        # Pagination
        try:
            page = max(int(request.args.get("page", 1)), 1)
        except ValueError:
            page = 1
        per_page = 20
        pagination = q.order_by(CongVan.id.desc()).paginate(page=page, per_page=per_page, error_out=False)
        items = pagination.items

        # Thống kê mềm
        chua_lay = (CongVan.query
                    .filter(CongVan.tinh_trang.ilike("%chưa lấy%"))
                    .order_by(CongVan.id.desc()).all())
        dang_giai_quyet = (CongVan.query
                           .filter(CongVan.tinh_trang.ilike("%đang giải%"))
                           .order_by(CongVan.id.desc()).all())

        # Hoàn thành theo loại — ưu tiên tháng/năm gần nhất; nếu thiếu thì đếm tổng
        last = (db.session.query(CongVan.nam, CongVan.thang)
                .filter(CongVan.tinh_trang.ilike("%hoàn thành%"),
                        CongVan.nam.isnot(None), CongVan.thang.isnot(None))
                .order_by(CongVan.nam.desc(), CongVan.thang.desc())
                .first())
        if last:
            latest_nam, latest_thang = int(last[0]), int(last[1])
            # tổng hồ sơ của tháng đó (mọi trạng thái)
            month_total = db.session.query(db.func.count()).filter(
                CongVan.nam == latest_nam, CongVan.thang == latest_thang
            ).scalar() or 0
            # số hoàn thành theo loại
            rows = (db.session.query(CongVan.loai_don_thu, db.func.count())
                    .filter(CongVan.tinh_trang.ilike("%hoàn thành%"),
                            CongVan.nam == latest_nam, CongVan.thang == latest_thang)
                    .group_by(CongVan.loai_don_thu).all())
            hoan_thanh_by_loai = [(r[0] or "(không ghi)", int(r[1])) for r in rows]
            # tổng hoàn thành tháng đó
            month_completed = sum(cnt for _, cnt in hoan_thanh_by_loai)
        else:
            # fallback đếm tổng mọi thời kỳ (nếu chưa có tháng/năm)
            rows = (db.session.query(CongVan.loai_don_thu, db.func.count())
                    .filter(CongVan.tinh_trang.ilike("%hoàn thành%"))
                    .group_by(CongVan.loai_don_thu).all())
            hoan_thanh_by_loai = [(r[0] or "(không ghi)", int(r[1])) for r in rows]

        # Options cho "Loại đơn thư"
        loai_rows = db.session.query(CongVan.loai_don_thu).distinct().all()
        loai_options = [r[0] for r in loai_rows if r[0]]

    except Exception:
        app.logger.exception("dashboard query failed; rendering empty dataset")
        pagination = None

    # Điều kiện hiển thị nút xuất Excel theo yêu cầu:
    has_exportable = bool(request.args.get("loai") and request.args.get("tinh_trang") and pagination and pagination.total > 0)

    return render_template(
        "dashboard.html",
        items=items, pager=pagination, page_url=page_url,
        chua_lay=chua_lay, dang_giai_quyet=dang_giai_quyet,
        hoan_thanh_by_loai=hoan_thanh_by_loai,
        latest_thang=latest_thang, latest_nam=latest_nam,
        month_total=month_total, month_completed=month_completed,
        loai_options=loai_options, STATUS_CHOICES=STATUS_CHOICES,
        has_exportable=has_exportable
    )

# -------------------------------------------------------
# EXPORT EXCEL theo bộ lọc (yêu cầu: có 'loai' & 'tinh_trang' & có ≥1 kết quả)
# -------------------------------------------------------
@app.route("/export")
@login_required
def export_excel():
    q = build_filtered_query()
    need_loai = request.args.get("loai")
    need_tt = request.args.get("tinh_trang")
    rows = q.order_by(CongVan.id.desc()).all()
    if not (need_loai and need_tt and len(rows) > 0):
        flash("Cần chọn cả Loại đơn thư và Tình trạng, và có ít nhất 1 kết quả để xuất Excel.", "error")
        return redirect(url_for("dashboard", **request.args))

    wb = Workbook()
    ws = wb.active
    ws.title = "CongVan"
    headers = [
        "ID", "Mã", "Loại đơn thư", "Tháng", "Năm", "Mã KH", "Tên", "Địa chỉ",
        "Nội dung", "Nhân viên", "Ngày NV nhận", "Tình trạng"
    ]
    ws.append(headers)
    for r in rows:
        ws.append([
            r.id, r.ma or "", r.loai_don_thu or "", r.thang or "", r.nam or "",
            r.ma_kh or "", r.ten or "", r.dia_chi or "", r.noi_dung or "",
            r.nhan_vien or "", (r.ngay_nv_nhan.strftime("%Y-%m-%d") if r.ngay_nv_nhan else ""),
            r.tinh_trang or ""
        ])
    for col in ws.columns:
        max_len = max(len(str(c.value)) if c.value else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(12, max_len + 2), 60)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    fname = f"congvan_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        bio,
        as_attachment=True,
        download_name=fname,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -------------------------------------------------------
# CRUD Công văn
# -------------------------------------------------------
@app.route("/congvan/<int:id>")
@login_required
def congvan_detail(id):
    cv = CongVan.query.get_or_404(id)
    return render_template("congvan_detail.html", r=cv)

@app.route("/congvan/new", methods=["GET", "POST"])
@login_required
def congvan_new():
    if current_user.role != "admin":
        flash("Bạn không có quyền.", "error")
        return redirect(url_for("dashboard"))
    if request.method == "POST":
        cv = CongVan(
            ma=request.form.get("ma") or None,
            loai_don_thu=request.form.get("loai_don_thu") or None,
            thang=(int(request.form.get("thang")) if request.form.get("thang") else None),
            nam=(int(request.form.get("nam")) if request.form.get("nam") else None),
            ma_kh=request.form.get("ma_kh") or None,
            ten=request.form.get("ten") or None,
            dia_chi=request.form.get("dia_chi") or None,
            nhan_vien=request.form.get("nhan_vien") or None,
            noi_dung=request.form.get("noi_dung") or None,
            ngay_nv_nhan=(datetime.strptime(request.form.get("ngay_nv_nhan"), "%Y-%m-%d").date()
                          if request.form.get("ngay_nv_nhan") else None),
            tinh_trang=request.form.get("tinh_trang") or None,
        )
        db.session.add(cv)
        db.session.commit()
        return redirect(url_for("dashboard"))
    return render_template("congvan_form.html")

@app.route("/congvan/<int:id>/edit", methods=["GET", "POST"])
@login_required
def congvan_edit(id):
    if current_user.role != "admin":
        flash("Bạn không có quyền.", "error")
        return redirect(url_for("dashboard"))
    cv = CongVan.query.get_or_404(id)
    if request.method == "POST":
        for f in ["ma","loai_don_thu","ma_kh","ten","dia_chi","nhan_vien","noi_dung","tinh_trang"]:
            setattr(cv, f, request.form.get(f) or None)
        cv.thang = int(request.form.get("thang")) if request.form.get("thang") else None
        cv.nam = int(request.form.get("nam")) if request.form.get("nam") else None
        cv.ngay_nv_nhan = (datetime.strptime(request.form.get("ngay_nv_nhan"), "%Y-%m-%d").date()
                           if request.form.get("ngay_nv_nhan") else None)
        db.session.commit()
        return redirect(url_for("dashboard"))
    return render_template("congvan_form.html", r=cv)

@app.route("/congvan/<int:id>/delete", methods=["POST"])
@login_required
def congvan_delete(id):
    if current_user.role != "admin":
        flash("Bạn không có quyền.", "error")
        return redirect(url_for("dashboard"))
    cv = CongVan.query.get_or_404(id)
    db.session.delete(cv)
    db.session.commit()
    return redirect(url_for("dashboard"))

# -------------------------------------------------------
# Quản lý người dùng (admin)
# -------------------------------------------------------
@app.route("/users")
@login_required
def users():
    if current_user.role != "admin":
        flash("Bạn không có quyền.", "error")
        return redirect(url_for("dashboard"))
    lst = User.query.order_by(User.id.asc()).all()
    return render_template("users.html", users=lst)

@app.post("/users/create")
@login_required
def users_create():
    if current_user.role != "admin":
        flash("Bạn không có quyền.", "error")
        return redirect(url_for("dashboard"))
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "").strip()
    role = request.form.get("role", "staff")
    if not username or not password or role not in ("admin","staff"):
        flash("Thiếu dữ liệu hợp lệ.", "error")
        return redirect(url_for("users"))
    if User.query.filter_by(username=username).first():
        flash("Tài khoản đã tồn tại.", "error")
        return redirect(url_for("users"))
    db.session.add(User(username=username, password=password, role=role))
    db.session.commit()
    flash("Tạo tài khoản thành công.", "ok")
    return redirect(url_for("users"))

@app.post("/users/<int:uid>/password")
@login_required
def users_password(uid):
    if current_user.role != "admin":
        flash("Bạn không có quyền.", "error")
        return redirect(url_for("dashboard"))
    u = User.query.get_or_404(uid)
    new_pw = request.form.get("new_password", "").strip()
    if not new_pw:
        flash("Mật khẩu mới không hợp lệ.", "error")
        return redirect(url_for("users"))
    u.password = new_pw
    db.session.commit()
    flash("Đổi mật khẩu thành công.", "ok")
    return redirect(url_for("users"))

@app.post("/users/<int:uid>/delete")
@login_required
def users_delete(uid):
    if current_user.role != "admin":
        flash("Bạn không có quyền.", "error")
        return redirect(url_for("dashboard"))
    if current_user.id == uid:
        flash("Không thể tự xoá tài khoản đang đăng nhập.", "error")
        return redirect(url_for("users"))
    u = User.query.get_or_404(uid)
    db.session.delete(u)
    db.session.commit()
    flash("Đã xoá tài khoản.", "ok")
    return redirect(url_for("users"))

# -------------------------------------------------------
# Error pages
# -------------------------------------------------------
@app.errorhandler(404)
def _404(e):
    return render_template("error.html", code=404, message="Không tìm thấy"), 404

@app.errorhandler(500)
def _500(e):
    return render_template("error.html", code=500, message="Lỗi máy chủ"), 500

# -------------------------------------------------------
# Run local
# -------------------------------------------------------
if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.getenv("PORT", 5000)))
