import os
from io import BytesIO
from datetime import datetime

from flask import (
    Flask, render_template, request, redirect,
    url_for, flash, send_file, abort
)
from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    LoginManager, login_user, logout_user,
    login_required, current_user, UserMixin
)
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import or_, func

# ============ App & DB ============
app = Flask(__name__)
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "dev-secret")

db_url = os.getenv("DATABASE_URL", "sqlite:///data.db")
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)
app.config["SQLALCHEMY_DATABASE_URI"] = db_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

# ============ Login ============
login_manager = LoginManager(app)
login_manager.login_view = "login"

# ============ HẰNG ============
STATUS_CHOICES = ["NV chưa lấy đơn", "Đang giải quyết", "Hoàn Thành"]
CHUA_LAY, DANG_GQ, HOAN_T = STATUS_CHOICES

APP_NAME = "Quản lý đơn thư - Đội 3"

# ============ Models ============
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(64), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(16), default="staff")  # admin | staff
    def set_password(self, raw): self.password_hash = generate_password_hash(raw)
    def check_password(self, raw): return check_password_hash(self.password_hash, raw)

class CongVan(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    loai_don_thu = db.Column(db.String(120))
    thang = db.Column(db.Integer)
    nam = db.Column(db.Integer)
    ma_kh = db.Column(db.String(64))
    ten = db.Column(db.String(120))
    dia_chi = db.Column(db.String(255))
    nhan_vien = db.Column(db.String(64))
    noi_dung = db.Column(db.Text)
    ngay_nv_nhan = db.Column(db.String(20))  # YYYY-MM-DD
    tinh_trang = db.Column(db.String(120))
    ghi_chu = db.Column(db.Text)
    ket_qua = db.Column(db.String(120))

@login_manager.user_loader
def load_user(uid): return User.query.get(int(uid))

# Khởi tạo DB + tài khoản mặc định
with app.app_context():
    db.create_all()
    if User.query.count() == 0:
        a = User(username=os.getenv("DEFAULT_ADMIN_USERNAME","admin"), role="admin")
        a.set_password(os.getenv("DEFAULT_ADMIN_PASSWORD","admin123"))
        s = User(username=os.getenv("DEFAULT_STAFF_USERNAME","nhanvien"), role="staff")
        s.set_password(os.getenv("DEFAULT_STAFF_PASSWORD","nhanvien123"))
        db.session.add_all([a, s]); db.session.commit()

# ============ Helpers ============
def admin_required():
    if not current_user.is_authenticated or current_user.role != "admin":
        abort(403)

def apply_filters(q):
    """Áp bộ lọc theo query string – dùng chung cho bảng & export."""
    ma_kh = request.args.get("ma_kh","").strip()
    ten = request.args.get("ten","").strip()
    dia_chi = request.args.get("dia_chi","").strip()
    loai = request.args.get("loai","").strip()
    thang = request.args.get("thang","").strip()
    nam = request.args.get("nam","").strip()
    tinh_trang = request.args.get("tinh_trang","").strip()
    keyword = request.args.get("q","").strip()

    if ma_kh:      q = q.filter(CongVan.ma_kh.ilike(f"%{ma_kh}%"))
    if ten:        q = q.filter(CongVan.ten.ilike(f"%{ten}%"))
    if dia_chi:    q = q.filter(CongVan.dia_chi.ilike(f"%{dia_chi}%"))
    if loai == "__EMPTY__":
        q = q.filter(or_(CongVan.loai_don_thu==None, CongVan.loai_don_thu==""))  # noqa
    elif loai:
        q = q.filter(CongVan.loai_don_thu == loai)
    if thang:      q = q.filter(CongVan.thang == int(thang))
    if nam:        q = q.filter(CongVan.nam == int(nam))
    if tinh_trang: q = q.filter(CongVan.tinh_trang == tinh_trang)  # so sánh BẰNG
    if keyword:    q = q.filter(CongVan.noi_dung.ilike(f"%{keyword}%"))
    return q

@app.context_processor
def _ctx():
    def page_url(p):
        args = request.args.to_dict(flat=True); args["page"] = p
        return url_for("dashboard", **args)
    loai_options = [r[0] for r in db.session.query(CongVan.loai_don_thu).distinct().all() if r[0]]
    return dict(page_url=page_url, STATUS_CHOICES=STATUS_CHOICES, loai_options=loai_options, APP_NAME=APP_NAME)

# ============ Auth ============
@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        u = User.query.filter_by(username=request.form.get("username","").strip()).first()
        if u and u.check_password(request.form.get("password","")):
            login_user(u); return redirect(url_for("dashboard"))
        flash("Sai tài khoản/mật khẩu","error")
    return render_template("login.html")

@app.get("/quick_login")
def quick_login():
    u = User.query.filter_by(role="staff").first()
    login_user(u); return redirect(url_for("dashboard"))

@app.get("/logout")
def logout():
    logout_user(); return redirect(url_for("login"))

# Health check cho Render
@app.get("/healthz")
def healthz():
    return "ok", 200

# ============ Dashboard & Stats ============
@app.get("/")
@login_required
def dashboard():
    all_rows = CongVan.query.order_by(CongVan.id.desc()).all()

    chua_lay = [r for r in all_rows if (r.tinh_trang or "") == CHUA_LAY]
    dang_giai_quyet = [r for r in all_rows if (r.tinh_trang or "") == DANG_GQ]
    # hoàn thành theo tháng gần nhất
    completed = [r for r in all_rows if (r.tinh_trang or "") == HOAN_T and r.thang and r.nam]

    latest_thang = latest_nam = None
    hoan_thanh_by_loai = []
    month_total = month_completed = 0

    if completed:
        latest_nam, latest_thang = max(((r.nam, r.thang) for r in completed))
        month_rows = [r for r in all_rows if r.nam==latest_nam and r.thang==latest_thang]
        month_total = len(month_rows)
        month_completed = sum(1 for r in month_rows if (r.tinh_trang or "") == HOAN_T)
        counter = {}
        for r in month_rows:
            if (r.tinh_trang or "") == HOAN_T:
                key = (r.loai_don_thu or "(không ghi)")
                counter[key] = counter.get(key,0)+1
        hoan_thanh_by_loai = sorted(counter.items(), key=lambda x: x[0])

    # Bảng chi tiết (admin mới thấy)
    q = apply_filters(CongVan.query.order_by(CongVan.id.desc()))
    page = int(request.args.get("page", 1)); per_page = 20
    pager = q.paginate(page=page, per_page=per_page, error_out=False)
    items = pager.items

    return render_template("dashboard.html",
        chua_lay=chua_lay,
        dang_giai_quyet=dang_giai_quyet,
        latest_thang=latest_thang, latest_nam=latest_nam,
        hoan_thanh_by_loai=hoan_thanh_by_loai,
        month_total=month_total, month_completed=month_completed,
        items=items, pager=pager
    )

# ============ CRUD Công văn ============
@app.route("/congvan/new", methods=["GET","POST"])
@login_required
def congvan_new():
    admin_required()
    if request.method == "POST":
        r = CongVan(
            loai_don_thu = request.form.get("loai_don_thu") or None,
            thang        = int(request.form.get("thang") or 0) or None,
            nam          = int(request.form.get("nam") or 0) or None,
            ma_kh        = request.form.get("ma_kh") or None,
            ten          = request.form.get("ten") or None,
            dia_chi      = request.form.get("dia_chi") or None,
            nhan_vien    = request.form.get("nhan_vien") or None,
            noi_dung     = request.form.get("noi_dung") or None,
            ngay_nv_nhan = request.form.get("ngay_nv_nhan") or None,
            tinh_trang   = request.form.get("tinh_trang") or CHUA_LAY,
            ghi_chu      = request.form.get("ghi_chu") or None,
            ket_qua      = request.form.get("ket_qua") or None
        )
        db.session.add(r); db.session.commit()
        flash("Đã thêm công văn","ok"); return redirect(url_for("dashboard"))
    today = datetime.now()
    return render_template("congvan_form.html", mode="new",
                           default_thang=today.month, default_nam=today.year)

@app.route("/congvan/<int:id>/edit", methods=["GET","POST"])
@login_required
def congvan_edit(id):
    admin_required()
    r = CongVan.query.get_or_404(id)
    if request.method == "POST":
        r.loai_don_thu = request.form.get("loai_don_thu") or None
        r.thang        = int(request.form.get("thang") or 0) or None
        r.nam          = int(request.form.get("nam") or 0) or None
        r.ma_kh        = request.form.get("ma_kh") or None
        r.ten          = request.form.get("ten") or None
        r.dia_chi      = request.form.get("dia_chi") or None
        r.nhan_vien    = request.form.get("nhan_vien") or None
        r.noi_dung     = request.form.get("noi_dung") or None
        r.ngay_nv_nhan = request.form.get("ngay_nv_nhan") or None
        r.tinh_trang   = request.form.get("tinh_trang") or CHUA_LAY
        r.ghi_chu      = request.form.get("ghi_chu") or None
        r.ket_qua      = request.form.get("ket_qua") or None
        db.session.commit()
        flash("Đã cập nhật công văn","ok"); return redirect(url_for("congvan_detail", id=r.id))
    return render_template("congvan_form.html", mode="edit", r=r,
                           default_thang=r.thang, default_nam=r.nam)

@app.post("/congvan/<int:id>/delete")
@login_required
def congvan_delete(id):
    admin_required()
    r = CongVan.query.get_or_404(id)
    db.session.delete(r); db.session.commit()
    flash("Đã xoá công văn","ok")
    return redirect(url_for("dashboard"))

# Staff cũng được xem chi tiết
@app.get("/congvan/<int:id>")
@login_required
def congvan_detail(id):
    r = CongVan.query.get_or_404(id)
    return render_template("congvan_detail.html", r=r)

# ============ Export Excel ============
@app.get("/export")
@login_required
def export_table():
    q = apply_filters(CongVan.query.order_by(CongVan.id.desc()))
    rows = q.all()

    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active; ws.title = "CongVan"
    headers = ["STT","Tháng","Năm","Loại","Mã KH","Tên","Địa chỉ","Nội dung",
               "Nhân viên","Ngày nhận","Tình trạng","Kết quả","Ghi chú"]
    ws.append(headers)
    for r in rows:
        ws.append([
            r.id, r.thang, r.nam, r.loai_don_thu, r.ma_kh, r.ten, r.dia_chi,
            (r.noi_dung or "")[:32760],
            r.nhan_vien, r.ngay_nv_nhan, r.tinh_trang, r.ket_qua, r.ghi_chu
        ])
    bio = BytesIO(); wb.save(bio); bio.seek(0)
    filename = f"congvan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(bio,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True, download_name=filename)

# ============ Error pages ============
@app.errorhandler(404)
def _404(e): return render_template("error.html", code=404, msg="Không tìm thấy trang"), 404

@app.errorhandler(403)
def _403(e): return render_template("error.html", code=403, msg="Không có quyền truy cập"), 403

@app.errorhandler(500)
def _500(e): return render_template("error.html", code=500, msg="Lỗi máy chủ"), 500

if __name__ == "__main__":
    app.run(debug=True)
