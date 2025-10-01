# ----- Quản lý người dùng (ADMIN) -----
@app.route('/users', methods=['GET', 'POST'])
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
                    flash('Nhập tài khoản và mật khẩu.', 'error')
                elif User.query.filter_by(username=username).first():
                    flash('Tài khoản đã tồn tại.', 'error')
                else:
                    u = User(username=username, role=role)
                    u.set_password(password)
                    db.session.add(u); db.session.commit()
                    flash('Đã tạo người dùng.', 'success')

            elif action == 'reset':
                uid = int(request.form['user_id'])
                pwd = request.form.get('password','').strip()
                if not pwd:
                    flash('Nhập mật khẩu mới.', 'error')
                else:
                    u = User.query.get_or_404(uid)
                    u.set_password(pwd); db.session.commit()
                    flash('Đã đặt lại mật khẩu.', 'success')

            elif action == 'delete':
                uid = int(request.form['user_id'])
                if current_user.id == uid:
                    flash('Không thể xoá tài khoản đang đăng nhập.', 'error')
                else:
                    u = User.query.get_or_404(uid)
                    db.session.delete(u); db.session.commit()
                    flash('Đã xoá người dùng.', 'success')
        except Exception as e:
            db.session.rollback(); flash(f'Lỗi: {e}', 'error')
        return redirect(url_for('users'))

    users = User.query.order_by(User.created_at.desc()).all()
    return render_template('users.html', users=users)
