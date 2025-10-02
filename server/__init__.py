import os
from io import BytesIO
from datetime import datetime, date

from flask import (
    Flask, render_template, request, redirect, url_for,
    flash, send_file, session
)
from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    LoginManager, login_user, logout_user,
    current_user, login_required, UserMixin
)
from sqlalchemy import text
from werkzeug.exceptions import Forbidden
from openpyxl import Workbook

# =================== App & DB ===================
APP_NAME = "Quản lý đơn thư đội 3"

def _normalize_pg_url(url: str) -> str:
    if url.startswith("postgres://"):
        return url.replace("postgres://", "postgresql://", 1)
    return url

db_url = _normalize_pg_url(os.getenv("DATABASE_URL", "sqlite:///data.db"))

app = Flask(__name__)
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "dev-key")
app.config["SQLALCHEMY_DATABASE_URI"] = db_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)

# =================== Models ===================
class User(db.Model, UserMixin):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(20), default="staff")  # 'admin' | 'staff'

class CongVan(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ma = db.Column(db.String(50))                 # Mã nội bộ (tùy chọn)
    loai_don_thu = db.Column(db.String(120))
    thang = db.Column(db.Integer)
    nam = db.Column(db.Integer)
    ma_kh = db.Column(db.String(50))
    ten = db.Column(db.String(120))
    dia_chi = db.Column(db.String(200))
    nhan_vien = db.Column(db.String(120))
    noi_dung = db.Column(db.Text)
    ngay_nv_nhan = db.Column(db.Date)
    tinh_trang = db.Column(db.String(50))         # NV chưa lấy đơn | Đang giải quyết | Hoàn Thành

# =================== Login ===================
login_manager = LoginManager(app)
login_manager.login_view = "login"

@login_manager.user_loader
def load_user(uid):
    try:
        return User.query.get(int(uid))
    except Exception:
        return None

STATUS_CHOICES = ["NV chưa lấy đơn", "Đang giải quyết", "Hoàn Thành"]

@app.context_processor
def inject_globals():
    return {"APP_NAME": APP_NAME, "STATUS_CHOICES": STATUS_CHOICES}

# =================== Health & Init ===================
@app.route("/healthz")
def healthz(): return "ok", 200

@app.route("/dbz")
def dbz():
    try:
        db.session.execute(text("SELECT 1"))
        return "db-ok", 200
    except Exception as e:
        app.logger.exception("db check error")
        return f"db-fail: {e}", 500

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
def _lazy_init():
    global _init_done
    if _init_done:
        return
    try:
        db.create_all()
        ensure_seed_users()
    finally:
        _init_done = True

# =================== Helpers ===================
def admin_required():
    if not (current_user.is_authenticated and current_user.role == "admin"):
        raise Forbidden("Admin only")

def page_url(num: int):
    args = dict(request.args)
    args["page"] = num
    return url_for("dashboard", **args)

def build_filtered_query():
    """Ghép điều kiện từ form tìm kiếm."""
    q = CongVan.query
    ten = request.args.get("ten", "").strip()
    ma_kh = request.args.get("ma_kh", "").strip()
    thang = request.args.get("thang", "").strip()
    nam = request.args.get("nam", "").strip()
    dia_chi = request.args.get("dia_chi", "").strip()
    loai = request.args.get("loai", "").strip()
    tinh_trang = request.args.get("tinh_trang", "").strip()
    kw = request.args.get("q", "").strip()

    if ten:       q = q.filter(CongVan.ten.ilike(f"%{ten}%"))
    if ma_kh:     q = q.filter(CongVan.ma_kh.ilike(f"%{ma_kh}%"))
    if thang:
        try: q = q.filter(CongVan.thang == int(thang))
        except ValueError: pass
    if nam:
        try: q = q.filter(CongVan.nam == int(nam))
        except ValueError: pass
    if dia_chi:   q = q.filter(CongVan.dia_chi.ilike(f"%{dia_chi}%"))
    if loai:
        if loai == "__EMPTY__":
            q = q.filter((CongVan.loai_don_thu == "") | (CongVan.loai_don_thu.is_(None)))
        else:
            q = q.filter(CongVan.loai_don_thu == loai)
    if tinh_trang: q = q.filter(CongVan.tinh_trang.ilike(f"%{tinh_trang}%"))
    if kw:         q = q.filter(CongVan.noi_dung.ilike(f"%{kw}%"))
    return q

def _make_excel(rows, prefix):
    """Tạo file Excel từ list CongVan."""
    wb = Workbook(); ws = wb.active; ws.title = "CongVan"
    headers = ["ID","Mã","Loại đơn thư","Tháng","Năm","Mã KH","Tên","Địa chỉ","Nội dung",
               "Nhân viên","Ngày NV nhận","Tình trạng"]
    ws.append(headers)
    for r in rows:
        ws.append([
            r.id, r.ma or "", r.loai_don_thu or "", r.thang or "", r.nam or "",
            r.ma_kh or "", r.ten or "", r.dia_chi or "", r.noi_dung or "",
            r.nhan_vien or "", (r.ngay_nv_nhan.strftime("%Y-%m-%d") if r.ngay_nv_nhan else ""),
            r.tinh_trang or ""
        ])
    for col in ws.columns:
        w = max(len(str(c.value)) if c.value else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(12, w+2), 60)
    bio = BytesIO(); wb.save(bio); bio.seek(0)
    fname = f"{prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(
        bio, as_attachment=True, download_name=fname,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def _redir_back():
    qs = request.query_string.decode()
    return redirect(url_for("dashboard") + (f"?{qs}" if qs else ""))

# =================== Auth ===================
@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        u = request.form.get("username","").strip()
        p = request.form.get("password","").strip()
        user = User.query.filter_by(username=u, password=p).first()
        if user:
            login_user(user)
            return redirect(url_for("dashboard"))
        flash("Sai tài khoản hoặc mật khẩu", "error")
    return render_template("login.html")

@app.route("/quick_login")
def quick_login():
    staff_u = os.getenv("DEFAULT_STAFF_USERNAME","nhanvien")
    staff_p = os.getenv("DEFAULT_STAFF_PASSWORD","nhanvien123")
    user = User.query.filter_by(username=staff_u, password=staff_p, role="staff").first()
    if not user:
        user = User(username=staff_u, password=staff_p, role="staff")
        db.session.add(user); db.session.commit()
    login_user(user); session["quick"]=True
    return redirect(url_for("dashboard"))

@app.route("/logout")
def logout():
    if current_user.is_authenticated:
        logout_user()
    return redirect(url_for("login"))

# =================== Dashboard ===================
@app.route("/")
@login_required
def dashboard():
    items=[]; pagination=None
    chua_lay=[]; dang_giai_quyet=[]
    latest_thang=None; latest_nam=None
    hoan_thanh_by_loai=[]; loai_options=[]
    month_total=0; month_completed=0

    try:
        q = build_filtered_query()
        # paging
        try: page = max(int(request.args.get("page",1)), 1)
        except ValueError: page = 1
        pagination = q.order_by(CongVan.id.desc()).paginate(page=page, per_page=20, error_out=False)
        items = pagination.items

        # 2 thống kê chính
        chua_lay = (CongVan.query
                    .filter(CongVan.tinh_trang.ilike("%chưa lấy%"))
                    .order_by(CongVan.id.desc()).all())
        dang_giai_quyet = (CongVan.query
                           .filter(CongVan.tinh_trang.ilike("%đang giải%"))
                           .order_by(CongVan.id.desc()).all())

        # Hoàn thành theo loại - tháng gần nhất
        last = (db.session.query(CongVan.nam, CongVan.thang)
                .filter(CongVan.tinh_trang.ilike("%hoàn thành%"),
                        CongVan.nam.isnot(None), CongVan.thang.isnot(None))
                .order_by(CongVan.nam.desc(), CongVan.thang.desc())
                .first())
        if last:
            latest_nam, latest_thang = int(last[0]), int(last[1])
            month_total = db.session.query(db.func.count()).filter(
                CongVan.nam==latest_nam, CongVan.thang==latest_thang
            ).scalar() or 0
            rows = (db.session.query(CongVan.loai_don_thu, db.func.count())
                    .filter(CongVan.tinh_trang.ilike("%hoàn thành%"),
                            CongVan.nam==latest_nam, CongVan.thang==latest_thang)
                    .group_by(CongVan.loai_don_thu).all())
            hoan_thanh_by_loai = [(r[0] or "(không ghi)", int(r[1])) for r in rows]
            month_completed = sum(cnt for _,cnt in hoan_thanh_by_loai)
        else:
            rows = (db.session.query(CongVan.loai_don_thu, db.func.count())
                    .filter(CongVan.tinh_trang.ilike("%hoàn thành%"))
                    .group_by(CongVan.loai_don_thu).all())
            hoan_thanh_by_loai = [(r[0] or "(không ghi)", int(r[1])) for r in rows]

        loai_rows = db.session.query(CongVan.loai_don_thu).distinct().all()
        loai_options = [r[0] for r in loai_rows if r[0]]

    except Exception:
        app.logger.exception("dashboard query failed")

    has_exportable = bool(request.args.get("loai") and request.args.get("tinh_trang")
                          and pagination and pagination.total>0)

    return render_template(
        "dashboard.html",
        items=items, pager=pagination, page_url=page_url,
        chua_lay=chua_lay, dang_giai_quyet=dang_giai_quyet,
        hoan_thanh_by_loai=hoan_thanh_by_loai,
        latest_thang=latest_thang, latest_nam=latest_nam,
        month_total=month_total, month_completed=month_completed,
        loai_options=loai_options, has_exportable=has_exportable
    )

# =================== Export Excel ===================
@app.route("/export")
@login_required
def export_excel():
    """Xuất Excel theo bộ lọc — yêu cầu Loại + Tình trạng và có ≥1 kết quả."""
    q = build_filtered_query()
    need_loai = request.args.get("loai")
    need_tt = request.args.get("tinh_trang")
    rows = q.order_by(CongVan.id.desc()).all()
    if not (need_loai and need_tt and rows):
        flash("Chọn Loại đơn thư + Tình trạng (và có kết quả) để xuất Excel.", "error")
        return _redir_back()
    return _make_excel(rows, "congvan_filter")

@app.route("/export_table")
@login_required
def export_table():
    """Xuất toàn bộ theo bộ lọc hiện tại (không điều kiện)."""
    rows = build_filtered_query().order_by(CongVan.id.desc()).all()
    if not rows:
        flash("Không có dữ liệu để xuất.", "error")
        return _redir_back()
    return _make_excel(rows, "congvan_all")

# =================== CRUD Công văn ===================
@app.route("/congvan/<int:id>")
@login_required
def congvan_detail(id):
    r = CongVan.query.get_or_404(id)
    return render_template("congvan_detail.html", r=r)

@app.route("/congvan/new", methods=["GET","POST"])
@login_required
def congvan_new():
    admin_required()
    if request.method == "POST":
        r = CongVan(
            ma=request.form.get("ma") or None,
            loai_don_thu=request.form.get("loai_don_thu") or None,
            thang=int(request.form.get("thang") or 0) or None,
            nam=int(request.form.get("nam") or 0) or None,
            ma_kh=request.form.get("ma_kh") or None,
            ten=request.form.get("ten") or None,
            dia_chi=request.form.get("dia_chi") or None,
            nhan_vien=request.form.get("nhan_vien") or None,
            noi_dung=request.form.get("noi_dung") or None,
            ngay_nv_nhan=(datetime.strptime(request.form.get("ngay_nv_nhan"), "%Y-%m-%d").date()
                          if request.form.get("ngay_nv_nhan") else None),
            tinh_trang=request.form.get("tinh_trang") or None
        )
        db.session.add(r); db.session.commit()
        flash("Đã thêm công văn", "ok")
        return redirect(url_for("dashboard"))
    # mặc định tháng/năm hiện tại
    today = date.today()
    return render_template("congvan_form.html", mode="new",
                           default_thang=today.month, default_nam=today.year)

@app.route("/congvan/<int:id>/edit", methods=["GET","POST"])
@login_required
def congvan_edit(id):
    admin_required()
    r = CongVan.query.get_or_404(id)
    if request.method == "POST":
        r.ma = request.form.get("ma") or None
        r.loai_don_thu = request.form.get("loai_don_thu") or None
        r.thang = int(request.form.get("thang") or 0) or None
        r.nam = int(request.form.get("nam") or 0) or None
        r.ma_kh = request.form.get("ma_kh") or None
        r.ten = request.form.get("ten") or None
        r.dia_chi = request.form.get("dia_chi") or None
        r.nhan_vien = request.form.get("nhan_vien") or None
        r.noi_dung = request.form.get("noi_dung") or None
        r.ngay_nv_nhan = (datetime.strptime(request.form.get("ngay_nv_nhan"), "%Y-%m-%d").date()
                          if request.form.get("ngay_nv_nhan") else None)
        r.tinh_trang = request.form.get("tinh_trang") or None
        db.session.commit()
        flash("Đã cập nhật công văn", "ok")
        return redirect(url_for("congvan_detail", id=r.id))
    return render_template("congvan_form.html", mode="edit", r=r,
                           default_thang=r.thang, default_nam=r.nam)

@app.route("/congvan/<int:id>/delete", methods=["POST"])
@login_required
def congvan_delete(id):
    admin_required()
    r = CongVan.query.get_or_404(id)
    db.session.delete(r); db.session.commit()
    flash("Đã xoá công văn", "ok")
    return redirect(url_for("dashboard"))

# =================== Users (Admin) ===================
@app.route("/users")
@login_required
def users():
    admin_required()
    us = User.query.order_by(User.id.asc()).all()
    return render_template("users.html", users=us)

@app.route("/users/create", methods=["POST"])
@login_required
def users_create():
    admin_required()
    u = request.form.get("username","").strip()
    p = request.form.get("password","").strip()
    role = request.form.get("role","staff")
    if not u or not p:
        flash("Thiếu username/password", "error"); return redirect(url_for("users"))
    if User.query.filter_by(username=u).first():
        flash("Username đã tồn tại", "error"); return redirect(url_for("users"))
    db.session.add(User(username=u, password=p, role=role)); db.session.commit()
    flash("Đã tạo tài khoản", "ok")
    return redirect(url_for("users"))

@app.route("/users/<int:id>/password", methods=["POST"])
@login_required
def users_password(id):
    admin_required()
    user = User.query.get_or_404(id)
    p = request.form.get("password","").strip()
    if not p:
        flash("Mật khẩu trống", "error"); return redirect(url_for("users"))
    user.password = p; db.session.commit()
    flash("Đã đổi mật khẩu", "ok")
    return redirect(url_for("users"))

@app.route("/users/<int:id>/delete", methods=["POST"])
@login_required
def users_delete(id):
    admin_required()
    user = User.query.get_or_404(id)
    if user.role == "admin" and user.id == current_user.id:
        flash("Không tự xoá admin đang đăng nhập", "error")
        return redirect(url_for("users"))
    db.session.delete(user); db.session.commit()
    flash("Đã xoá tài khoản", "ok")
    return redirect(url_for("users"))

# =================== Errors ===================
@app.errorhandler(404)
def _404(e): return render_template("error.html", code=404, msg="Không tìm thấy"), 404

@app.errorhandler(403)
def _403(e): return render_template("error.html", code=403, msg="Không đủ quyền"), 403

@app.errorhandler(500)
def _500(e):
    app.logger.exception("server error")
    return render_template("error.html", code=500, msg="Lỗi máy chủ"), 500

# =================== Main ===================
if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.getenv("PORT", "5000")))
