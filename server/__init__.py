import os
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager

app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY','dev-secret-change-me')

db_url = os.getenv('DATABASE_URL', 'sqlite:///data.db')
if db_url.startswith('postgres://'):
    db_url = db_url.replace('postgres://', 'postgresql://', 1)
app.config['SQLALCHEMY_DATABASE_URI'] = db_url
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

@app.route('/')
@login_required
def dashboard():
    # dropdown Loại
    loai_options = [
        row[0] for row in db.session.query(CongVan.loai_don_thu)
        .filter(CongVan.loai_don_thu.isnot(None))
        .filter(CongVan.loai_don_thu != '')
        .distinct()
        .order_by(CongVan.loai_don_thu.asc())
        .all()
    ]

    # danh sách cho 2 box thống kê
    dang_giai_quyet = CongVan.query.filter_by(tinh_trang='Đang giải quyết').all()
    chua_lay = CongVan.query.filter_by(tinh_trang='NV chưa lấy đơn').all()

    # thêm đếm Hoàn Thành để làm card
    hoan_thanh_count = db.session.query(CongVan).filter_by(tinh_trang='Hoàn Thành').count()

    # danh sách kết quả theo bộ lọc
    items = _query_congvan_from_request().order_by(CongVan.created_at.desc()).all()

    return render_template(
        'dashboard.html',
        dang_giai_quyet=dang_giai_quyet,
        chua_lay=chua_lay,
        hoan_thanh_count=hoan_thanh_count,
        items=items,
        loai_options=loai_options
    )
