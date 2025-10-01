import os
from datetime import datetime
from functools import wraps

from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash

from openpyxl import Workbook
import io
from datetime import datetime as _dt

app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'dev-secret-change-me')

db_url = os.getenv('DATABASE_URL', 'sqlite:///data.db')
if db_url.startswith('postgres://'):
    db_url = db_url.replace('postgres://', 'postgresql://', 1)
app.config['SQLALCHEMY_DATABASE_URI'] = db_url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

# ================= Models =================
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(20), default='staff')  # 'admin' | 'staff'
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def set_password(self, pwd: str):
        self.password_hash = generate_password_hash(pwd, method='pbkdf2:sha256')

    def check_password(self, pwd: str) -> bool:
        return check_password_hash(self.password_hash, pwd)

class CongVan(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ma = db.Column(db.String(50), unique=True, nullable=False)
    loai_don_thu = db.Column(db.String(120))
    thang = db.Column(db.Integer)
    nam = db.Column(db.Integer)
    ma_kh = db.Column(db.String(50))
    ten = db.Column(db.String(200), nullable=False)
    dia_chi = db.Column(db.String(300))
    nhan_vien = db.Column(db.String(120))
    noi_dung = db.Column(db.Text)
    ngay_nv_nhan = db.Column(db.Date)
    tinh_trang = db.Column(db.String(50), default='NV chưa lấy đơn')  # NV chưa lấy đơn | Đang giải quyết | Hoàn Thành
    ghi_chu = db.Column(db.Text)
    created_by_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not current_user.is_authenticated or current_user.role != 'admin':
            flash('Chức năng này chỉ dành cho ADMIN.', 'error')
            return redirect(url_for('dashboard'))
        return view(*args, **kwargs)
    return wrapped

def bootstrap():
    db.create_all()
    if User.query.count() == 0:
        admin = User(username=os.getenv('DEFAULT_ADMIN_USERNAME', 'admin'))
        admin.set_password(os.getenv('DEFAULT_ADMIN_PASSWORD', 'admin123'))
        admin.role = 'admin'
        staff = User(username=os.getenv('DEFAULT_STAFF_USERNAME', 'nhanvien'))
        staff.set_password(os.getenv('DEFAULT_STAFF_PASSWORD', 'nhanvien123'))
        staff.role = 'staff'
        db.session.add_all([admin, staff]); db.session.commit()

with app.app_context():
    bootstrap()

# ================= Auth =================
@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        u = User.query.filter_by(username=request.form.get('username','').strip()).first()
        if u and u.check_password(request.form.get('password','')):
            login_user(u); flash('Đăng nhập thành công.','success'); return redirect(url_for('dashboard'))
        flash('Sai tài khoản hoặc mật khẩu.','error')
    return render_template('login.html')

@app.route('/quick-login')
def quick_login():
    u = User.query.filter_by(username='quick_staff').first()
    if not u:
        u = User(username='quick_staff', role='staff'); u.set_password('quick'); db.session.add(u); db.session.commit()
    login_user(u); flash('Đăng nhập nhanh (quyền xem).','success'); return redirect(url_for('dashboard'))

@app.route('/logout')
@login_required
def logout():
    logout_user(); flash('Đã đăng xuất.','success'); return redirect(url_for('login'))

# ================= Query helper (có thêm lọc: loai, q) =================
def _query_congvan_from_request():
    q = CongVan.query
    ma = request.args.get('ma','').strip()
    ten = request.args.get('ten','').strip()
    dia_chi = request.args.get('dia_chi','').strip()
    thang = request.args.get('thang','').strip()
    nam = request.args.get('nam','').strip()
    tinh_trang = request.args.get('tinh_trang','').strip()
    loai = request.args.get('loai','').strip()   # NEW
    kw = request.args.get('q','').strip()        # NEW

    if ma: q = q.filter(CongVan.ma.ilike(f"%{ma}%"))
    if ten: q = q.filter(CongVan.ten.ilike(f"%{ten}%"))
    if dia_chi: q = q.filter(CongVan.dia_chi.ilike(f"%{dia_chi}%"))
    if thang.isdigit(): q = q.filter_by(thang=int(thang))
    if nam.isdigit(): q = q.filter_by(nam=int(nam))
    if tinh_trang: q = q.filter_by(tinh_trang=tinh_trang)
    if loai: q = q.filter(CongVan.loai_don_thu.ilike(f"%{loai}%"))
    if kw: q = q.filter(CongVan.noi_dung.ilike(f"%{kw}%"))
    return q

# ================= Dashboard (root) =================
@app.route('/')
@login_required
def dashboard():
    dang_giai_quyet = CongVan.query.filter_by(tinh_trang='Đang giải quyết').all()
    chua_lay = CongVan.query.filter_by(tinh_trang='NV chưa lấy đơn').all()
    items = _query_congvan_from_request().order_by(CongVan.created_at.desc()).all()
    return render_template('dashboard.html', dang_giai_quyet=dang_giai_quyet, chua_lay=chua_lay, items=items)

# ================= CRUD (chỉ ADMIN) =================
@app.route('/congvan/new', methods=['GET','POST'])
@admin_required
def congvan_new():
    if request.method == 'POST':
        try:
            ngay = request.form.get('ngay_nv_nhan') or None
            ngay = _dt.strptime(ngay, "%Y-%m-%d").date() if ngay else None
            cv = CongVan(
                ma=request.form['ma'].strip(),
                loai_don_thu=request.form.get('loai_don_thu','').strip(),
                thang=int(request.form.get('thang')) if request.form.get('thang') else None,
                nam=int(request.form.get('nam')) if request.form.get('nam') else None,
                ma_kh=request.form.get('ma_kh','').strip(),
                ten=request.form.get('ten','').strip(),
                dia_chi=request.form.get('dia_chi','').strip(),
                nhan_vien=request.form.get('nhan_vien','').strip(),
                noi_dung=request.form.get('noi_dung','').strip(),
                ngay_nv_nhan=ngay,
                tinh_trang=request.form.get('tinh_trang','NV chưa lấy đơn'),
                ghi_chu=request.form.get('ghi_chu','').strip(),
                created_by_id=current_user.id
            )
            db.session.add(cv); db.session.commit(); flash('Đã thêm công văn.','success'); return redirect(url_for('dashboard'))
        except Exception as e:
            db.session.rollback(); flash(f'Lỗi: {e}','error')
    return render_template('congvan_form.html', mode='new', item=None, today=datetime.today())

@app.route('/congvan/<int:id>/edit', methods=['GET','POST'])
@admin_required
def congvan_edit(id):
    item = CongVan.query.get_or_404(id)
    if request.method == 'POST':
        try:
            item.ma = request.form['ma'].strip()
            item.loai_don_thu = request.form.get('loai_don_thu','').strip()
            item.thang = int(request.form.get('thang')) if request.form.get('thang') else None
            item.nam = int(request.form.get('nam')) if request.form.get('nam') else None
            item.ma_kh = request.form.get('ma_kh','').strip()
            item.ten = request.form.get('ten','').strip()
            item.dia_chi = request.form.get('dia_chi','').strip()
            item.nhan_vien = request.form.get('nhan_vien','').strip()
            item.noi_dung = request.form.get('noi_dung','').strip()
            ngay = request.form.get('ngay_nv_nhan') or None
            item.ngay_nv_nhan = _dt.strptime(ngay, "%Y-%m-%d").date() if ngay else None
            item.tinh_trang = request.form.get('tinh_trang','NV chưa lấy đơn')
            item.ghi_chu = request.form.get('ghi_chu','').strip()
            db.session.commit(); flash('Đã cập nhật công văn.','success'); return redirect(url_for('dashboard'))
        except Exception as e:
            db.session.rollback(); flash(f'Lỗi: {e}','error')
    return render_template('congvan_form.html', mode='edit', item=item, today=datetime.today())

@app.route('/congvan/<int:id>/delete', methods=['POST'])
@admin_required
def congvan_delete(id):
    item = CongVan.query.get_or_404(id)
    db.session.delete(item); db.session.commit(); flash('Đã xoá công văn.','success'); return redirect(url_for('dashboard'))

# ================= Chi tiết công văn =================
@app.route('/congvan/<int:id>')
@login_required
def congvan_detail(id):
    item = CongVan.query.get_or_404(id)
    return render_template('congvan_detail.html', item=item)

# ================= Export Excel =================
@app.route('/export')
@login_required
def export_excel():
    rows = _query_congvan_from_request().order_by(CongVan.created_at.desc()).all()

    headers = [
        'Mã','Loại đơn thư','Tháng','Năm','Mã KH','Tên','Địa chỉ','Nhân viên',
        'Nội dung đơn','Ngày NV nhận đơn','Tình trạng xử lý','Ghi chú',
        'Ngày tạo','Ngày sửa'
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "CongVan"
    ws.append(headers)

    for r in rows:
        ws.append([
            r.ma, r.loai_don_thu, r.thang, r.nam, r.ma_kh, r.ten, r.dia_chi, r.nhan_vien,
            r.noi_dung,
            r.ngay_nv_nhan.isoformat() if r.ngay_nv_nhan else '',
            r.tinh_trang, r.ghi_chu,
            r.created_at.strftime('%Y-%m-%d %H:%M:%S') if r.created_at else '',
            r.updated_at.strftime('%Y-%m-%d %H:%M:%S') if r.updated_at else ''
        ])

    for col in ws.columns:
        max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(8, max_len + 2), 60)

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name='congvan.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.errorhandler(404)
def e404(e): return render_template('error.html', code=404, message='Không tìm thấy trang'), 404

@app.errorhandler(500)
def e500(e): return render_template('error.html', code=500, message='Lỗi máy chủ'), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.getenv('PORT', 5050)))
