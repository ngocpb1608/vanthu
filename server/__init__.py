import os
import unicodedata
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

# ================= App & DB =================
app = Flask(__name__)
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "dev-secret")

db_url = os.getenv("DATABASE_URL", "sqlite:///data.db")
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)
app.config["SQLALCHEMY_DATABASE_URI"] = db_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

# ================= Login =================
login_manager = LoginManager(app)
login_manager.login_view = "login"

# ================= Models =================
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
    ngay_nv_nhan = db.Column(db.String(20))  # YYYY-MM-DD (chuỗi)
    tinh_trang = db.Column(db.String(120))
    ghi_chu = db.Column(db.Text)
    ket_qua = db.Column(db.String(120))

@login_manager.user_loader
def load_user(uid): return User.query.get(int(uid))

# Khởi tạo DB + tài khoản mẫu
with app.app_context():
    db.create_all()
    if User.query.count() == 0:
        a_user = os.getenv("DEFAULT_ADMIN_USERNAME", "admin")
        a_pass = os.getenv("DEFAULT_ADMIN_PASSWORD", "admin123")
        s_user = os.getenv("DEFAULT_STAFF_USERNAME", "nhanvien")
        s_pass = os.getenv("DEFAULT_STAFF_PASSWORD", "nhanvien123")
        a = User(username=a_user, role="admin"); a.set_password(a_pass)
        s = User(username=s_user, role="staff"); s.set_password(s_pass)
        db.session.add_all([a, s]); db.session.commit()

# ================= Helpers =================
def admin_required():
    if not current_user.is_authenticated or current_user.role != "admin":
        abort(403)

STATUS_CHOICES = ["NV chưa lấy đơn", "Đang giải quyết", "Hoàn Thành"]

def _ilike_any(col, keywords):
    patterns = [f"%{kw.lower()}%" for kw in keywords]
    return or_(*[func.lower(col).like(p) for p in patterns])

def apply_filters(q):
    """Áp bộ lọc theo tham số GET (dùng cho bảng & export)."""
    ma_kh = request.args.get("ma_kh", "").strip()
    ten = request.args.get("ten", "").strip()
    dia_chi = request.args.get("dia_chi", "").strip()
    loai = request.args.get("loai", "").strip()
    thang = request.args.get("thang", "").strip()
    nam = request.args.get("nam", "").strip()
    tinh_trang = request.args.get("tinh_trang", "").strip()
    keyword = request.args.get("q", "").strip()

    if ma_kh:      q = q.filter(CongVan.ma_kh.ilike(f"%{ma_kh}%"))
    if ten:        q = q.filter(CongVan.ten.ilike(f"%{ten}%"))
    if dia_chi:    q = q.filter(CongVan.dia_chi.ilike(f"%{dia_chi}%"))
    if loai == "__EMPTY__":
        q = q.filter(or_(CongVan.loai_don_thu == None, CongVan.loai_don_thu == ""))  # noqa
    elif loai:
        q = q.filter(CongVan.loai_don_thu == loai)
    if thang:      q = q.filter(CongVan.thang == int(thang))
    if nam:        q = q.filter(CongVan.nam == int(nam))
    if tinh_trang: q = q.filter(CongVan.tinh_trang.ilike(f"%{tinh_trang}%"))
    if keyword:    q = q.filter(CongVan.noi_dung.ilike(f"%{keyword}%"))
    return q

# ---- Chuẩn hoá tiếng Việt (bỏ dấu, lower, gọn khoảng trắng) ----
def vn_norm(s: str) -> str:
    if not s: return ""
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = s.lower().strip()
    s = " ".join(s.split())
    return s

def is_chua_lay(s: str) -> bool:
    v = vn_norm(s)
    return ("chua lay" in v) or ("nv chua lay" in v)

def is_dang_gq(s: str) -> bool:
    v = vn_norm(s)
    return ("dang giai quyet" in v) or ("dang xu ly" in v)

def is_hoan_thanh(s: str) -> bool:
    v = vn_norm(s)
    return ("hoan thanh" in v) or ("da hoan thanh" in v) or ("hoan tat" in v)

# ================= Auth =================
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        u = User.query.filter_by(username=request.form.get("username","").strip()).first()
        if u and u.check_password(request.form.get("password","")):
            login_user(u); return redirect(url_for("dashboard"))
        flash("Sai tài khoản/mật khẩu", "error")
    return render_template("login.html")

@app.route("/quick_login")
def quick_login():
    u = User.query.filter_by(role="staff").first()
    login_user(u); return redirect(url_for("dashboard"))

@app.route("/logout")
def logout():
    logout_user(); return redirect(url_for("login"))

# ================= Dashboard & Stats =================
@app.route("/")
@login_required
def dashboard():
    # Lấy toàn bộ để phân loại robust bằng Python (tránh miss do khác chính tả/dấu)
    all_rows = CongVan.query.order_by(CongVan.id.desc()).all()

    chua_lay = [r for r in all_rows if is_chua_lay(r.tinh_trang)]
    dang_giai_quyet = [r for r in all_rows if is_dang_gq(r.tinh_trang)]

    # Hoàn thành theo loại – xác định tháng gần nhất có "hoàn thành"
    completed = [r for r in all_rows if is_hoan_thanh(r.tinh_trang) and r.thang and r.nam]
    latest_thang = latest_nam = None
    hoan_thanh_by_loai = []
    month_total = month_completed = 0

    if completed:
        # tìm (nam, thang) lớn nhất
        latest_nam, latest_thang = max(((r.nam, r.thang) for r in completed))
        # tổng tất cả đơn của tháng đó
        month_rows = [r for r in all_rows if r.nam == latest_nam and r.thang == latest_thang]
        month_total = len(month_rows)
        # số hoàn thành trong tháng đó
        month_completed = sum(1 for r in month_rows if is_hoan_thanh(r.tinh_trang))
        # nhóm theo loại
        counter = {}
        for r in month_rows:
            if is_hoan_thanh(r.tinh_trang):
                key = (r.loai_don_thu or "(không ghi)")
                counter[key] = counter.get(key, 0) + 1
        # sắp xếp theo tên loại
        hoan_thanh_by_loai = sorted(counter.items(), key=lambda x: x[0])

    # Bảng (admin xem); dữ liệu lọc theo tham số
    loai_options = [r[0] for r in db.session.query(CongVan.loai_don_thu).distinct().all() if r[0]]
    q = apply_filters(CongVan.query.order_by(CongVan.id.desc()))
    page = int(request.args.get("page", 1)); per_page = 20
    pager = q.paginate(page=page, per_page=per_page, error_out=False)
    items = pager.items

    return render_template(
        "dashboard.html",
        chua_lay=chua_lay,
        dang_giai_quyet=dang_giai_quyet,
        latest_thang=latest_thang,
        latest_nam=latest_nam,
        hoan_thanh_by_loai=hoan_thanh_by_loai,
        month_total=month_total,
        month_completed=month_completed,
        items=items, pager=pager,
        STATUS_CHOICES=STATUS_CHOICES,
        loai_options=loai_options,
        APP_NAME="Đơn thư đội 3"
    )

# ================= CRUD Công văn =================
@app.route("/congvan/new", methods=["GET", "POST"])
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
            tinh_trang   = request.form.get("tinh_trang") or None,
            ghi_chu      = request.form.get("ghi_chu") or None,
            ket_qua      = request.form.get("ket_qua") or None
        )
        db.session.add(r); db.session.commit()
        flash("Đã thêm công văn", "ok")
        return redirect(url_for("dashboard"))
    today = datetime.now()
    return render_template("congvan_form.html",
                           mode="new",
                           default_thang=today.month,
                           default_nam=today.year)

@app.route("/congvan/<int:id>/edit", methods=["GET", "POST"])
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
        r.tinh_trang   = request.form.get("tinh_trang") or None
        r.ghi_chu      = request.form.get("ghi_chu") or None
        r.ket_qua      = request.form.get("ket_qua") or None
        db.session.commit()
        flash("Đã cập nhật công văn", "ok")
        return redirect(url_for("congvan_detail", id=r.id))
    return render_template("congvan_form.html",
                           mode="edit", r=r,
                           default_thang=r.thang, default_nam=r.nam)

@app.route("/congvan/<int:id>/delete", methods=["POST"])
@login_required
def congvan_delete(id):
    admin_required()
    r = CongVan.query.get_or_404(id)
    db.session.delete(r); db.session.commit()
    flash("Đã xoá công văn", "ok")
    return redirect(url_for("dashboard"))

# ================= Detail (CHO PHÉP STAFF XEM) =================
@app.route("/congvan/<int:id>")
@login_required
def congvan_detail(id):
    r = CongVan.query.get_or_404(id)
    return render_template("congvan_detail.html", r=r)

# ================= Export Excel =================
@app.route("/export")
@login_required
def export_table():
    # Nếu muốn cấm staff export: uncomment
    # admin_required()
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
    return send_file(
        bio,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename
    )

# ================= Jinja helpers =================
@app.context_processor
def _ctx():
    def page_url(p):
        args = request.args.to_dict(flat=True); args["page"] = p
        return url_for("dashboard", **args)
    return dict(page_url=page_url, APP_NAME="Đơn thư đội 3", STATUS_CHOICES=STATUS_CHOICES)

if __name__ == "__main__":
    app.run(debug=True)
