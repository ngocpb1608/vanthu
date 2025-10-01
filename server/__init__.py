# server/__init__.py
import os
import io
from datetime import datetime as _dt, date
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for, flash,
    send_file, g
)
from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    LoginManager, UserMixin, login_user, login_required,
    logout_user, current_user
)
from werkzeug.security import generate_password_hash, check_password_hash
from sqlalchemy import or_, func, desc
from openpyxl import Workbook

# ===================== APP & DB =====================
app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'dev-secret-change-me')

db_url = os.getenv('DATABASE_URL', 'sqlite:///data.db')
if db_url.startswith('postgres://'):
    db_url = db_url.replace('postgres://', 'postgresql://', 1)
app.config['SQLALCHEMY_DATABASE_URI'] = db_url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

PAGE_SIZE = int(os.getenv('PAGE_SIZE', '50'))  # phân trang danh sách

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

# ===================== MODELS =====================
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(20), default='staff')  # 'admin' | 'staff'
    created_at = db.Column(db.DateTime, default=_dt.utcnow)

    def set_password(self, pwd: str):
        self.password_hash = generate_password_hash(pwd, method='pbkdf2:sha256')

    def check_password(self, pwd: str) -> bool:
        return check_password_hash(self.password_hash, pwd)


class CongVan(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    ma = db.Column(db.String(50), unique=True, nullable=False, index=True)
    loai_don_thu = db.Column(db.String(120), index=True)
    thang = db.Column(db.Integer, index=True)
    nam = db.Column(db.Integer, index=True)
    ma_kh = db.Column(db.String(50), index=True)
    ten = db.Column(db.String(200), nullable=False, index=True)
    dia_chi = db.Column(db.String(300))
    nhan_vien = db.Column(db.String(120), index=True)
    noi_dung = db.Column(db.Text)
    ngay_nv_nhan = db.Column(db.Date)
    tinh_trang = db.Column(db.String(50), default='NV chưa lấy đơn', index=True)
    ghi_chu = db.Column(db.Text)
    created_by_id = db.Column(db.Integer, db.ForeignKey('user.id'))
    created_at = db.Column(db.DateTime, default=_dt.utcnow)
    updated_at = db.Column(db.DateTime, default=_dt.utcnow, onupdate=_dt.utcnow)

# ===================== LOGIN/ACCESS =====================
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

# ===================== BOOTSTRAP =====================
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

    if os.getenv('SEED_DEMO', '0') == '1' and CongVan.query.count() == 0:
        samples = [
            CongVan(ma='CV-202501-0001', loai_don_thu='Khiếu nại', thang=1, nam=2025, ma_kh='KH01',
                    ten='Nguyễn Văn A', dia_chi='Hà Nội', nhan_vien='Lan',
                    noi_dung='Khiếu nại thủ tục cấp sổ.', ngay_nv_nhan=date(2025,1,5),
                    tinh_trang='Đang giải quyết'),
            CongVan(ma='CV-202501-0002', loai_don_thu='Đề nghị', thang=1, nam=2025, ma_kh='KH02',
                    ten='Trần Thị B', dia_chi='Đà Nẵng', nhan_vien='Minh',
                    noi_dung='Đề nghị xác nhận tạm trú.', ngay_nv_nhan=date(2025,1,10),
                    tinh_trang='NV chưa lấy đơn'),
            CongVan(ma='CV-202412-0001', loai_don_thu='Phản ánh', thang=12, nam=2024, ma_kh='KH03',
                    ten='Phạm Văn C', dia_chi='TP.HCM', nhan_vien='Hùng',
                    noi_dung='Phản ánh thái độ phục vụ.', ngay_nv_nhan=date(2024,12,28),
                    tinh_trang='Hoàn Thành', ghi_chu='Đã phản hồi bằng văn bản.')
        ]
        db.session.add_all(samples); db.session.commit()

with app.app_context():
    bootstrap()

# ===================== GLOBALS =====================
@app.before_request
def _set_ctx():
    g.is_admin = current_user.is_authenticated and current_user.role == 'admin'

@app.context_processor
def inject_globals():
    def page_url(p):
        args = request.args.to_dict()
        args['page'] = p
        return url_for('dashboard', **args)
    return dict(
        APP_NAME='Đơn Thư & Công Văn',
        now=_dt.utcnow,
        is_admin=(current_user.is_authenticated and current_user.role == 'admin'),
        STATUS_CHOICES=['NV chưa lấy đơn','Đang giải quyết','Hoàn Thành'],
        PAGE_SIZE=PAGE_SIZE,
        page_url=page_url,
    )

# ===================== HELPERS =====================
def _safe_int(val, default=None):
    try: return int(val)
    except Exception: return default

def _distinct_loai_options():
    return [
        r[0] for r in db.session.query(CongVan.loai_don_thu)
        .filter(CongVan.loai_don_thu.isnot(None), CongVan.loai_don_thu!='')
        .distinct().order_by(CongVan.loai_don_thu.asc()).all()
    ]

def _generate_ma(thang=None, nam=None) -> str:
    """Tự sinh mã CV-YYYYMM-#### (duy nhất)."""
    today = _dt.today()
    thang = thang or today.month
    nam = nam or today.year
    prefix = f"CV-{nam:04d}{thang:02d}-"
    last = db.session.query(CongVan.ma).filter(CongVan.ma.like(prefix + "%")) \
           .order_by(CongVan.ma.desc()).first()
    next_num = 1
    if last and last[0].startswith(prefix):
        try: next_num = int(last[0].split('-')[-1]) + 1
        except: pass
    return f"{prefix}{next_num:04d}"

def _query_congvan_from_request():
    """Áp bộ lọc từ form tìm kiếm."""
    q = CongVan.query
    ma = request.args.get('ma','').strip()
    ten = request.args.get('ten','').strip()
    dia_chi = request.args.get('dia_chi','').strip()
    thang = request.args.get('thang','').strip()
    nam = request.args.get('nam','').strip()
    tinh_trang = request.args.get('tinh_trang','').strip()
    loai = request.args.get('loai','').strip()
    kw = request.args.get('q','').strip()

    if ma: q = q.filter(CongVan.ma.ilike(f'%{ma}%'))
    if ten: q = q.filter(CongVan.ten.ilike(f'%{ten}%'))
    if dia_chi: q = q.filter(CongVan.dia_chi.ilike(f'%{dia_chi}%'))
    if thang.isdigit(): q = q.filter_by(thang=int(thang))
    if nam.isdigit(): q = q.filter_by(nam=int(nam))
    if tinh_trang: q = q.filter_by(tinh_trang=tinh_trang)

    if loai == '__EMPTY__':
        q = q.filter(or_(CongVan.loai_don_thu == None, CongVan.loai_don_thu == ''))  # noqa: E711
    elif loai:
        q = q.filter(CongVan.loai_don_thu == loai)

    if kw: q = q.filter(CongVan.noi_dung.ilike(f'%{kw}%'))
    return q

def _latest_done_month():
    row = db.session.query(CongVan.nam, CongVan.thang) \
        .filter(CongVan.tinh_trang=='Hoàn Thành') \
        .order_by(desc(CongVan.nam), desc(CongVan.thang)).first()
    if row: return int(row[0]), int(row[1])
    return None, None

def _done_by_loai(nam, thang):
    rows = db.session.query(CongVan.loai_don_thu, func.count(CongVan.id)) \
        .filter(CongVan.tinh_trang=='Hoàn Thành',
                CongVan.nam==nam, CongVan.thang==thang) \
        .group_by(CongVan.loai_don_thu).order_by(func.count(CongVan.id).desc()).all()
    return [((loai or '(Trống)'), cnt) for loai, cnt in rows]

# ===================== AUTH =====================
@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username','').strip()
        password = request.form.get('password','').strip()
        u = User.query.filter_by(username=username).first()
        if u and u.check_password(password):
            login_user(u); flash('Đăng nhập thành công.','success')
            return redirect(url_for('dashboard'))
        flash('Sai tài khoản hoặc mật khẩu.','error')
    return render_template('login.html')

@app.route('/quick-login')
def quick_login():
    u = User.query.filter_by(username='quick_staff').first()
    if not u:
        u = User(username='quick_staff', role='staff')
        u.set_password('quick'); db.session.add(u); db.session.commit()
    login_user(u); flash('Đăng nhập nhanh (quyền xem).','success')
    return redirect(url_for('dashboard'))

@app.route('/logout')
@login_required
def logout():
    logout_user(); flash('Đã đăng xuất.','success')
    return redirect(url_for('login'))

# ===================== DASHBOARD =====================
@app.route('/')
@login_required
def dashboard():
    loai_options = _distinct_loai_options()
    # 2 nhóm thống kê nhanh
    dang_giai_quyet = CongVan.query.filter_by(tinh_trang='Đang giải quyết').all()
    chua_lay = CongVan.query.filter_by(tinh_trang='NV chưa lấy đơn').all()
    # Tháng gần nhất có 'Hoàn Thành' + thống kê theo Loại
    latest_nam, latest_thang = _latest_done_month()
    hoan_thanh_by_loai = _done_by_loai(latest_nam, latest_thang) if latest_nam else []

    # Danh sách có phân trang
    page = request.args.get('page', 1, type=int)
    query = _query_congvan_from_request().order_by(CongVan.created_at.desc())
    pager = db.paginate(query, page=page, per_page=PAGE_SIZE, error_out=False)
    items = pager.items

    return render_template(
        'dashboard.html',
        dang_giai_quyet=dang_giai_quyet,
        chua_lay=chua_lay,
        latest_nam=latest_nam, latest_thang=latest_thang,
        hoan_thanh_by_loai=hoan_thanh_by_loai,
        items=items, loai_options=loai_options, pager=pager
    )

# ===================== CRUD: CHỈ ADMIN =====================
@app.route('/congvan/new', methods=['GET','POST'])
@admin_required
def congvan_new():
    if request.method == 'POST':
        try:
            # Không nhận 'ma' -> tự sinh
            thang_val = _safe_int(request.form.get('thang')) or _dt.today().month
            nam_val = _safe_int(request.form.get('nam')) or _dt.today().year
            auto_ma = _generate_ma(thang_val, nam_val)

            ngay_raw = request.form.get('ngay_nv_nhan') or None
            ngay = _dt.strptime(ngay_raw, '%Y-%m-%d').date() if ngay_raw else None

            cv = CongVan(
                ma=auto_ma,
                loai_don_thu=request.form.get('loai_don_thu','').strip(),
                thang=thang_val, nam=nam_val,
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
            db.session.add(cv); db.session.commit()
            flash(f'Đã thêm công văn. Mã tự sinh: {cv.ma}','success')
            return redirect(url_for('dashboard'))
        except Exception as e:
            db.session.rollback(); flash(f'Lỗi: {e}','error')

    loai_options = _distinct_loai_options()
    return render_template('congvan_form.html', mode='new', item=None, today=_dt.today(), loai_options=loai_options)

@app.route('/congvan/<int:id>/edit', methods=['GET','POST'])
@admin_required
def congvan_edit(id):
    item = CongVan.query.get_or_404(id)
    if request.method == 'POST':
        try:
            # Mã không cho sửa
            item.loai_don_thu = request.form.get('loai_don_thu','').strip()
            item.thang = _safe_int(request.form.get('thang'))
            item.nam = _safe_int(request.form.get('nam'))
            item.ma_kh = request.form.get('ma_kh','').strip()
            item.ten = request.form.get('ten','').strip()
            item.dia_chi = request.form.get('dia_chi','').strip()
            item.nhan_vien = request.form.get('nhan_vien','').strip()
            item.noi_dung = request.form.get('noi_dung','').strip()
            ngay_raw = request.form.get('ngay_nv_nhan') or None
            item.ngay_nv_nhan = _dt.strptime(ngay_raw, '%Y-%m-%d').date() if ngay_raw else None
            item.tinh_trang = request.form.get('tinh_trang','NV chưa lấy đơn')
            item.ghi_chu = request.form.get('ghi_chu','').strip()
            db.session.commit()
            flash('Đã cập nhật công văn.','success')
            return redirect(url_for('dashboard'))
        except Exception as e:
            db.session.rollback(); flash(f'Lỗi: {e}','error')

    loai_options = _distinct_loai_options()
    return render_template('congvan_form.html', mode='edit', item=item, today=_dt.today(), loai_options=loai_options)

@app.route('/congvan/<int:id>/delete', methods=['POST'])
@admin_required
def congvan_delete(id):
    item = CongVan.query.get_or_404(id)
    db.session.delete(item); db.session.commit()
    flash('Đã xoá công văn.','success')
    return redirect(url_for('dashboard'))

# ===================== DETAIL =====================
@app.route('/congvan/<int:id>')
@login_required
def congvan_detail(id):
    item = CongVan.query.get_or_404(id)
    return render_template('congvan_detail.html', item=item)

# ===================== USERS (ADMIN) =====================
@app.route('/users', methods=['GET','POST'])
@admin_required
def users():
    if request.method == 'POST':
        action = request.form.get('action')
        try:
            if action == 'create':
                username = request.form.get('username','').strip()
                password = request.form.get('password','').strip()
                role = request.form.get('role','staff')
                if not username or not password:
                    flash('Nhập tài khoản và mật khẩu.','error')
                elif User.query.filter_by(username=username).first():
                    flash('Tài khoản đã tồn tại.','error')
                else:
                    u = User(username=username, role=role)
                    u.set_password(password)
                    db.session.add(u); db.session.commit()
                    flash('Đã tạo người dùng.','success')
            elif action == 'reset':
                uid = int(request.form['user_id'])
                pwd = request.form.get('password','').strip()
                if not pwd:
                    flash('Nhập mật khẩu mới.','error')
                else:
                    u = User.query.get_or_404(uid)
                    u.set_password(pwd); db.session.commit()
                    flash('Đã đặt lại mật khẩu.','success')
            elif action == 'delete':
                uid = int(request.form['user_id'])
                if current_user.id == uid:
                    flash('Không thể xoá tài khoản đang đăng nhập.','error')
                else:
                    u = User.query.get_or_404(uid)
                    db.session.delete(u); db.session.commit()
                    flash('Đã xoá người dùng.','success')
        except Exception as e:
            db.session.rollback(); flash(f'Lỗi: {e}','error')
        return redirect(url_for('users'))

    users = User.query.order_by(User.created_at.desc()).all()
    return render_template('users.html', users=users)

# ===================== EXPORT EXCEL =====================
@app.route('/export')
@login_required
def export_excel():
    rows = _query_congvan_from_request().order_by(CongVan.created_at.desc()).all()
    headers = ['Mã','Loại đơn thư','Tháng','Năm','Mã KH','Tên','Địa chỉ','Nhân viên',
               'Nội dung đơn','Ngày NV nhận đơn','Tình trạng xử lý','Ghi chú',
               'Ngày tạo','Ngày sửa']
    wb = Workbook(); ws = wb.active; ws.title = 'CongVan'; ws.append(headers)
    for r in rows:
        ws.append([r.ma, r.loai_don_thu, r.thang, r.nam, r.ma_kh, r.ten, r.dia_chi, r.nhan_vien,
                   r.noi_dung, r.ngay_nv_nhan.isoformat() if r.ngay_nv_nhan else '',
                   r.tinh_trang, r.ghi_chu,
                   r.created_at.strftime('%Y-%m-%d %H:%M:%S') if r.created_at else '',
                   r.updated_at.strftime('%Y-%m-%d %H:%M:%S') if r.updated_at else ''])
    for col in ws.columns:
        max_len = max(len(str(c.value)) if c.value is not None else 0 for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(10, max_len + 2), 60)
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name='congvan.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ===================== MISC =====================
@app.route('/healthz')
def healthz():
    return {'status':'ok'}, 200

@app.errorhandler(404)
def e404(e):
    return render_template('error.html', code=404, message='Không tìm thấy trang'), 404

@app.errorhandler(500)
def e500(e):
    return render_template('error.html', code=500, message='Lỗi máy chủ'), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.getenv('PORT', 5050)))
