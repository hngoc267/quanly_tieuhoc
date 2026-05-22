import json
from datetime import datetime, timedelta
from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash
from supabase import create_client
import pandas as pd
from io import BytesIO
import random
import re
from openpyxl import Workbook
import io
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = 'chuoi-bao-mat-bat-ky'

# 🔗 KẾT NỐI SUPABASE
url = "https://xamrdpqetfqefbrrhryr.supabase.co"
key = "sb_publishable_N_m7lCFseorDvmFWU70qKA_5yIQg2SC"
supabase = create_client(url, key)

# ======================
# ROUTE CƠ BẢN
# ======================
@app.route('/')
def index():
    return redirect(url_for('login'))


# Thêm đoạn này vào app.py để mọi trang đều tự có user_name
@app.context_processor
def inject_user():
    return dict(user_name=session.get('user_name'))


# ======================
# LOGIN
# ======================
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        input_phone = request.form.get('phone')
        input_password = request.form.get('password')

        res = supabase.table("users") \
            .select("*") \
            .eq("username", input_phone) \
            .eq("is_active", True) \
            .execute()

        user = res.data[0] if res.data else None

        if user and user["password_hash"] == input_password:
            session['user_id'] = user['user_id']
            session['user_name'] = user['username']

            # lấy role
            role_res = supabase.table("roles") \
                .select("*") \
                .eq("role_id", user["role_id"]) \
                .execute()

            role = role_res.data[0] if role_res.data else None
            session['role_name'] = role['role_name'] if role else ""

            if session['role_name'] == 'Admin':
                return redirect(url_for('admin_dashboard'))
            elif session['role_name'] == 'Teacher':
                return redirect(url_for('teacher_dashboard'))

        return render_template('auth/login.html', error="Sai tài khoản hoặc mật khẩu!")

    return render_template('auth/login.html')


# ======================
# QUÊN MẬT KHẨU
# ======================
@app.route('/forgot-password', methods=['GET', 'POST'])
def forgot_password():
    if request.method == 'GET':
        step = session.get('step', 1)
        error = session.pop('error', None)
        return render_template('auth/forgot.html', step=step, error=error)
    if request.method == 'POST':

        # STEP 1: nhập SĐT
        if 'send_otp' in request.form:
            phone = request.form.get('phone') or session.get('reset_phone')

            res = supabase.table("users") \
                .select("*") \
                .eq("username", phone) \
                .execute()

            user = res.data[0] if res.data else None

            if not user:
                return render_template('auth/forgot.html', error="Không tìm thấy tài khoản!")

            otp = str(random.randint(100000, 999999))
            print("OTP:", otp)
            session['reset_phone'] = phone
            session['otp'] = otp

            session['step'] = 2
            return redirect(url_for('forgot_password'))

        # STEP 2: nhập OTP + đổi pass
        elif 'reset_password' in request.form:
            input_otp = request.form.get('otp')
            new_password = request.form.get('new_password')
            confirm_password = request.form.get('confirm_password')

            if new_password != confirm_password:
                if new_password != confirm_password:
                    session['step'] = 2
                    session['error'] = "Mật khẩu xác nhận không khớp!"
                    return redirect(url_for('forgot_password'))

            # ❌ sai OTP
            # ❌ sai OTP
            if input_otp != session.get('otp'):
                new_otp = str(random.randint(100000, 999999))
                session['otp'] = new_otp

                session['step'] = 2
                session['error'] = "OTP không đúng!"
                return redirect(url_for('forgot_password'))

            # 🔥 validate password
            error_msg = validate_password(new_password)
            if error_msg:
                # 👉 tạo OTP mới
                new_otp = str(random.randint(100000, 999999))
                session['otp'] = new_otp

                session['step'] = 2
                session['error'] = error_msg
                return redirect(url_for('forgot_password'))

            phone = session.get('reset_phone')

            res = supabase.table("users") \
                .select("*") \
                .eq("username", phone) \
                .execute()

            user = res.data[0]

            supabase.table("users") \
                .update({"password_hash": new_password}) \
                .eq("user_id", user["user_id"]) \
                .execute()

            session.pop('otp', None)
            session.pop('reset_phone', None)
            session.pop('step', None)

            return redirect(url_for('login'))

    return render_template('auth/forgot.html', step=1)


def validate_password(password):
    if len(password) < 8:
        return "Mật khẩu phải ≥ 8 ký tự"

    if not re.search(r"[A-Z]", password):
        return "Phải có chữ in HOA"

    if not re.search(r"[a-z]", password):
        return "Phải có chữ thường"

    if not re.search(r"\d", password):
        return "Phải có số"

    if not re.search(r"[!@#$%^&*(),.?\":{}|<>]", password):
        return "Phải có ký tự đặc biệt"

    return None


# ======================
# ADMIN DASHBOARD
# ======================
@app.route('/admin/dashboard')
def admin_dashboard():
    if 'user_id' not in session or session.get('role_name') != 'Admin':
        return redirect(url_for('login'))

    # Đếm giáo viên
    res = supabase.table("teachers").select("*").execute()
    total_teachers = len(res.data)

    # Lấy user hiện tại
    user_res = supabase.table("users") \
        .select("*") \
        .eq("user_id", session['user_id']) \
        .execute()

    current_user = user_res.data[0] if user_res.data else None

    # Lấy học kỳ hiện tại
    sem_res = supabase.table("semesters") \
        .select("*") \
        .eq("is_current", True) \
        .execute()

    current_sem = sem_res.data[0] if sem_res.data else None

    return render_template(
        'admin/dashboard.html',
        user_name=current_user['username'] if current_user else "",
        semester_name=current_sem['semester_name'] if current_sem else "N/A",
        school_name="SỞ GIÁO DỤC VÀ ĐÀO TẠO KHÁNH HÒA",
        teacher_count=total_teachers
    )


# ======================
# Đăng xuất
# ======================
@app.route('/logout')
def logout():
    session.clear()  # 🔥 xoá toàn bộ session
    return redirect(url_for('login'))


# ======================
# TEACHER DASHBOARD
# ======================
@app.route('/teacher/dashboard')
def teacher_dashboard():
    if 'user_id' not in session or session.get('role_name') != 'Teacher':
        return redirect(url_for('login'))

    # Lấy thông tin giáo viên
    teacher_res = supabase.table("teachers") \
        .select("*") \
        .eq("user_id", session['user_id']) \
        .execute()

    teacher = teacher_res.data[0] if teacher_res.data else None
    display_name = teacher['full_name'] if teacher else "Giáo viên"

    # Học kỳ
    sem_res = supabase.table("semesters") \
        .select("*") \
        .eq("is_current", True) \
        .execute()

    current_sem = sem_res.data[0] if sem_res.data else None

    return render_template(
        'teacher/dashboard.html',
        user_name=display_name,
        semester_name=current_sem['semester_name'] if current_sem else "N/A",
        school_name="SỞ GIÁO DỤC VÀ ĐÀO TẠO KHÁNH HÒA"
    )


# ======================
# QUẢN LÝ NĂM HỌC
# ======================
@app.route('/admin/namhoc', methods=['GET', 'POST'])
def quanly_namhoc():
    if 'user_id' not in session or session.get('role_name') != 'Admin':
        return redirect(url_for('login'))

    success_msg = request.args.get('msg')
    error_msg = request.args.get('error_msg')

    # === 1. XỬ LÝ CẬP NHẬT TỪ FORM ===
    if request.method == 'POST':
        year_id = request.form.get('year_id')
        term1_id = request.form.get('term1_id')
        term2_id = request.form.get('term2_id')

        t1_mid, t1_fin = request.form.get('term1_midterm'), request.form.get('term1_final')
        t2_mid, t2_fin = request.form.get('term2_midterm'), request.form.get('term2_final')

        # Logic chặn lỗi ngày tháng
        if (t1_mid and t1_fin and int(t1_mid) >= int(t1_fin)) or \
                (t2_mid and t2_fin and int(t2_mid) >= int(t2_fin)):
            return redirect(
                url_for('quanly_namhoc', year_id=year_id, error_msg='Lỗi: Tuần giữa kỳ phải trước cuối kỳ!'))

        try:
            # FIX LỖI OVERWRITE NULL: Chỉ cập nhật khi HTML có gửi data (không bị disabled)
            if term1_id:
                t1_data = {}
                if t1_mid: t1_data['midterm_start'] = int(t1_mid)
                if t1_fin: t1_data['final_start'] = int(t1_fin)
                if t1_data:  # Có dữ liệu mới gọi update
                    supabase.table('semesters').update(t1_data).eq('semester_id', term1_id).execute()

            if term2_id:
                t2_data = {}
                if t2_mid: t2_data['midterm_start'] = int(t2_mid)
                if t2_fin: t2_data['final_start'] = int(t2_fin)
                if t2_data:
                    supabase.table('semesters').update(t2_data).eq('semester_id', term2_id).execute()

            return redirect(url_for('quanly_namhoc', year_id=year_id, msg='Đã cập nhật thành công!'))
        except Exception as e:
            return redirect(url_for('quanly_namhoc', year_id=year_id, error_msg=f'Lỗi DB: {str(e)}'))

    # === 2. LẤY DỮ LIỆU & TÌM KIẾM ===
    search_kw = request.args.get('q', '')
    query = supabase.table("academic_years").select("*").order("year_name", desc=True)
    if search_kw: query = query.ilike("year_name", f"%{search_kw}%")
    years_list = query.execute().data

    selected_year_id = request.args.get('year_id')
    selected_year = next((y for y in years_list if str(y['year_id']) == str(selected_year_id)),
                         None) if selected_year_id else \
        (next((y for y in years_list if y.get('is_current')), years_list[0]) if years_list else None)

    term1, term2 = None, None
    term1_weeks, term2_weeks = [], []

    if selected_year:
        terms_list = supabase.table("semesters").select("*").eq("year_id", selected_year['year_id']).order(
            "semester_name").execute().data

        def format_vn_date(d_str):
            try:
                return datetime.strptime(d_str, '%Y-%m-%d').strftime('%d/%m/%Y')
            except:
                return "(chưa có)"

        def get_weeks(start_date_str, total_weeks, start_week_num=1):
            if not start_date_str: return []
            try:
                start = datetime.strptime(start_date_str, '%Y-%m-%d')
                return [
                    {'week_num': start_week_num + i, 'start_str': (start + timedelta(days=i * 7)).strftime('%d/%m/%Y'),
                     'end_str': (start + timedelta(days=i * 7 + 6)).strftime('%d/%m/%Y')} for i in range(total_weeks)]
            except:
                return []

        # FIX LỖI LỘN KỲ: Bắt chữ "kỳ 1" thay vì chỉ tìm số "1" để tránh nhầm năm học (VD: 2021)
        term1 = next((t for t in terms_list if 'kỳ 1' in str(t.get('semester_name', '')).lower()), None)
        term2 = next((t for t in terms_list if 'kỳ 2' in str(t.get('semester_name', '')).lower()), None)

        if not term1 and len(terms_list) > 0: term1 = terms_list[0]
        if not term2 and len(terms_list) > 1: term2 = terms_list[1]

        year_start = selected_year.get('start_date')
        year_end = selected_year.get('end_date')

        if term1:
            term1['start_vn'] = format_vn_date(year_start) if year_start else "(chưa có)"
            term1['end_vn'] = "Giữa năm"
            term1_weeks = get_weeks(year_start, 18, 1)

        if term2:
            term2['start_vn'] = "Giữa năm"
            term2['end_vn'] = format_vn_date(year_end) if year_end else "(chưa có)"
            start_w2 = len(term1_weeks) + 1 if term1_weeks else 19

            term2_start_date = None
            if year_start:
                try:
                    term2_start_date = (datetime.strptime(year_start, '%Y-%m-%d') + timedelta(weeks=18)).strftime(
                        '%Y-%m-%d')
                except:
                    pass
            term2_weeks = get_weeks(term2_start_date, 17, start_w2)

    return render_template('admin/namhoc.html', years_list=years_list, selected_year=selected_year, term1=term1,
                           term2=term2, term1_weeks=term1_weeks, term2_weeks=term2_weeks, search_kw=search_kw,
                           success_msg=success_msg, error_msg=error_msg, user_name=session.get('user_name'))


# ======================
# API: TẢI FILE EXCEL MẪU
# ======================
@app.route('/admin/khoilop/taimau')
def tai_mau_khoilop():
    # Tạo dữ liệu mẫu chuẩn
    data = {
        'Khối (*)': ['Khối 1', 'Khối 1', 'Khối 2'],
        'Tên lớp (*)': ['1A', '1B', '2A'],
        'Phòng học': ['Phòng 101', 'Phòng 102', '']
    }
    df = pd.DataFrame(data)

    # Ghi ra file ảo (trên RAM) để gửi cho user mà không cần lưu ổ cứng
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Danh_sach_lop')
    output.seek(0)

    return send_file(output, download_name="Mau_Nhap_Lop.xlsx", as_attachment=True)


# ======================
# DANH SÁCH KHỐI & LỚP (CẬP NHẬT LOGIC IMPORT)
# ======================
@app.route('/admin/khoilop', methods=['GET', 'POST'])
def danhsach_khoilop():
    if 'user_id' not in session or session.get('role_name') != 'Admin':
        return redirect(url_for('login'))

    success_msg = request.args.get('msg')
    error_msg = request.args.get('error_msg')

    # Lấy thông tin năm học hiện tại
    current_year_data = supabase.table('academic_years').select('year_id, year_name').eq('is_current', True).execute().data
    current_year_id = current_year_data[0]['year_id'] if current_year_data else None
    current_year_name = current_year_data[0]['year_name'] if current_year_data else "Chưa xác định"

    # Lấy sẵn danh sách Khối để dùng chung
    grades_db = supabase.table('grades').select('*').execute().data
    grade_map = {g['grade_name'].strip().lower(): g['grade_id'] for g in
                 grades_db}  # Map tên khối ra ID để tra cứu nhanh

    if request.method == 'POST':
        action = request.form.get('action')

        # --- LOGIC THÊM LỚP MỚI ---
        if action == 'add':
            if not current_year_id:
                return redirect(url_for('danhsach_khoilop', error_msg='Chưa thiết lập năm học hiện tại!'))

            grade_id = request.form.get('grade_id')
            class_name = request.form.get('class_name', '').strip().upper()
            room = request.form.get('room', '').strip()

            existing = supabase.table('classes').select('class_id').eq('grade_id', grade_id).eq('class_name',
                                                                                                class_name).eq('academic_year_id', current_year_id).execute().data
            if existing: return redirect(url_for('danhsach_khoilop', error_msg=f'Lớp {class_name} đã tồn tại trong năm học này!'))

            supabase.table('classes').insert(
                {'grade_id': grade_id, 'class_name': class_name, 'room': room, 'student_count': 0, 'academic_year_id': current_year_id}).execute()
            return redirect(url_for('danhsach_khoilop', msg='Đã thêm lớp mới thành công!'))

        # --- LOGIC CẬP NHẬT (SỬA) LỚP ---
        elif action == 'edit':
            class_id = request.form.get('class_id')
            grade_id = request.form.get('grade_id')
            class_name = request.form.get('class_name', '').strip().upper()
            room = request.form.get('room', '').strip()

            # Kiểm tra xem sửa tên lớp có bị trùng với lớp KHÁC đang có sẵn không
            existing = supabase.table('classes').select('class_id').eq('grade_id', grade_id).eq('class_name',
                                                                                                class_name).eq('academic_year_id', current_year_id).neq(
                'class_id', class_id).execute().data
            if existing: return redirect(
                url_for('danhsach_khoilop', error_msg=f'Tên lớp {class_name} bị trùng với lớp khác trong năm học này!'))

            # Cập nhật vào DB
            supabase.table('classes').update({'grade_id': grade_id, 'class_name': class_name, 'room': room}).eq(
                'class_id', class_id).execute()
            return redirect(url_for('danhsach_khoilop', msg='Cập nhật thông tin lớp thành công!'))

        # --- LOGIC XÓA LỚP ---
        elif action == 'delete':
            class_ids_str = request.form.get('class_ids', '')
            if class_ids_str:
                # Chuyển chuỗi "1,2,3" thành mảng số nguyên [1, 2, 3]
                class_ids = [int(x) for x in class_ids_str.split(',') if x.isdigit()]
                try:
                    # Lệnh xóa nhiều dòng cùng lúc bằng .in_
                    supabase.table('classes').delete().in_('class_id', class_ids).execute()
                    return redirect(url_for('danhsach_khoilop', msg='Đã xóa các lớp thành công!'))
                except Exception as e:
                    return redirect(
                        url_for('danhsach_khoilop', error_msg='Không thể xóa do lớp đang có ràng buộc dữ liệu!'))

        # --- LOGIC NHẬP TỪ EXCEL ---
        elif action == 'preview_import':
            file = request.files.get('excel_file')
            if not file or file.filename == '':
                return {"status": "error", "message": "Vui lòng chọn file!"}, 400

            try:
                import pandas as pd
                df = pd.read_excel(file)

                if 'Khối (*)' not in df.columns or 'Tên lớp (*)' not in df.columns:
                    return {"status": "error", "message": "File không đúng định dạng mẫu!"}, 400

                grades_db = supabase.table('grades').select('*').execute().data
                grade_map = {g['grade_name'].strip().lower(): g['grade_id'] for g in grades_db}

                valid_rows = []
                error_rows = []

                for index, row in df.iterrows():
                    row_num = index + 2  # Dòng trong Excel (bỏ qua header)
                    grade_name_excel = str(row['Khối (*)']).strip()
                    class_name_excel = str(row['Tên lớp (*)']).strip().upper()
                    room_excel = str(row.get('Phòng học', '')).strip()
                    if room_excel == 'nan': room_excel = ''

                    # Lỗi 1: Bỏ trống
                    if not grade_name_excel or grade_name_excel == 'nan' or not class_name_excel or class_name_excel == 'nan':
                        error_rows.append(
                            {"row": row_num, "class_name": class_name_excel if class_name_excel != 'nan' else '(Trống)',
                             "error": "Thiếu Khối hoặc Tên lớp"})
                        continue

                    # Lỗi 2: Sai tên Khối
                    grade_id = grade_map.get(grade_name_excel.lower())
                    if not grade_id:
                        error_rows.append({"row": row_num, "class_name": class_name_excel,
                                           "error": f"Khối '{grade_name_excel}' không tồn tại"})
                        continue

                    # Lỗi 3: Trùng lớp đã có
                    existing = supabase.table('classes').select('class_id').eq('grade_id', grade_id).eq('class_name',
                                                                                                        class_name_excel).eq('academic_year_id', current_year_id).execute().data
                    if existing:
                        error_rows.append(
                            {"row": row_num, "class_name": class_name_excel, "error": f"Lớp đã tồn tại trong năm học {current_year_name}"})
                        continue

                    # Nếu qua hết 3 ải trên thì là dòng Hợp lệ
                    valid_rows.append({
                        'grade_id': grade_id,
                        'class_name': class_name_excel,
                        'room': room_excel
                    })

                # Trả cục dữ liệu này về cho Javascript hiện Popup
                return {
                    "status": "success", "filename": file.filename,
                    "total": len(df), "valid_count": len(valid_rows), "error_count": len(error_rows),
                    "valid_data": valid_rows, "error_data": error_rows
                }

            except Exception as e:
                return {"status": "error", "message": "Lỗi định dạng file Excel!"}, 400

            # --- 5. LƯU EXCEL THẬT SỰ (KHI BẤM "TIẾP TỤC" TRÊN POPUP) ---
        elif action == 'confirm_import':
            valid_data_str = request.form.get('valid_data')

            try:
                valid_data = json.loads(valid_data_str)
                if len(valid_data) == 0:
                    return redirect(url_for('danhsach_khoilop', error_msg='Không có dòng nào hợp lệ để nhập!'))

                # Nhét vào Database
                for item in valid_data:
                    item['student_count'] = 0
                    item['academic_year_id'] = current_year_id
                    supabase.table('classes').insert(item).execute()

                return redirect(url_for('danhsach_khoilop', msg=f'Đã nhập thành công {len(valid_data)} lớp mới!'))
            except Exception as e:
                return redirect(url_for('danhsach_khoilop', error_msg='Lỗi khi lưu vào Database!'))
    # === 2. LẤY DỮ LIỆU HIỂN THỊ (GET) ===

    # 2.1. Lấy danh sách khối (Sắp xếp chuẩn theo grade_number)
    grades = supabase.table('grades').select('*').order('grade_number').execute().data
    selected_grade = request.args.get('grade_filter', '')

    # 2.2. Tìm ID của Học kỳ hiện tại (is_current = true)
    current_sem_data = supabase.table('semesters').select('semester_id').eq('is_current', True).execute().data
    current_sem_id = current_sem_data[0]['semester_id'] if current_sem_data else None

    # 2.3. Lấy toàn bộ danh sách lớp kèm tên Khối
    query = supabase.table('classes').select('*, grades(grade_name)')
    if selected_grade:
        query = query.eq('grade_id', selected_grade)
    if current_year_id:
        query = query.eq('academic_year_id', current_year_id)
    classes_raw = query.order('class_name').execute().data

    # 2.4. Nếu có Học kỳ hiện tại -> Lấy danh sách GVCN
    gvcn_map = {}
    if current_sem_id:
        # Lấy dữ liệu phân công chủ nhiệm
        hr_assignments = supabase.table('homeroom_assignments') \
            .select('class_id, teachers(full_name)') \
            .execute().data

        # Tạo Map nhanh (Dictionary) để tra cứu: class_id -> Tên GVCN
        for hr in hr_assignments:
            if hr.get('teachers'):
                gvcn_map[hr['class_id']] = hr['teachers'].get('full_name', '')

    # 2.5. Ép phẳng dữ liệu (Merge Data) để đẩy ra HTML
    classes = []
    for c in classes_raw:
        c['grade_name'] = c.get('grades', {}).get('grade_name', '') if c.get('grades') else ''
        c['homeroom_teacher'] = gvcn_map.get(c['class_id'], '')  # Lấy tên GVCN từ Map
        classes.append(c)

    return render_template(
        'admin/khoilop.html',
        grades=grades,
        classes=classes,
        selected_grade=selected_grade,
        current_year_name=current_year_name,
        success_msg=success_msg,
        error_msg=error_msg,
        user_name=session.get('user_name')
    )


@app.route('/admin/khoitao_hocsinh', methods=['GET', 'POST'])
def khoitao_hocsinh():
    if 'user_id' not in session or session.get('role_name') != 'Admin':
        return redirect(url_for('login'))

    # Lấy năm học hiện tại
    year_res = supabase.table("academic_years").select("year_id").eq("is_current", True).execute()
    current_year_id = year_res.data[0]['year_id'] if year_res.data else None

    if request.method == 'POST':
        action = request.form.get('action')

        if not current_year_id:
            if action == 'preview_import':
                return {"status": "error", "message": "Lỗi: Chưa thiết lập năm học hiện tại!"}, 400
            flash('Lỗi: Chưa thiết lập năm học hiện tại!', 'error')
            return redirect(url_for('khoitao_hocsinh'))

        if action == 'preview_import':
            file = request.files.get('excel_file')
            if not file:
                return {"status": "error", "message": "Vui lòng chọn file!"}, 400
            try:
                df = pd.read_excel(file)
                valid_rows, error_rows = [], []

                # Lấy map lớp của năm học hiện tại để kiểm tra
                classes_db = supabase.table('classes').select('class_id, class_name').eq('academic_year_id', current_year_id).execute().data
                class_map = {c['class_name'].strip().upper(): c['class_id'] for c in classes_db}

                for index, row in df.iterrows():
                    ho_ten = str(row.get('Họ tên', '')).strip()
                    lop = str(row.get('Tên lớp', '')).strip().upper()

                    if not ho_ten or ho_ten == 'nan' or not lop or lop == 'nan':
                        error_rows.append({"row": index + 2, "name": ho_ten or '(Trống)', "error": "Thiếu thông tin Họ tên hoặc Tên lớp"})
                        continue
                    
                    if lop not in class_map:
                        error_rows.append({"row": index + 2, "name": ho_ten, "error": f"Lớp '{lop}' không tồn tại trong năm học hiện tại"})
                        continue
                    
                    valid_rows.append({
                        "class_id": class_map[lop],
                        "ho_ten": ho_ten,
                        "ngay_sinh": str(row.get('Ngày sinh', ''))[:10] if pd.notna(row.get('Ngày sinh')) else None,
                        "gioi_tinh": str(row.get('Giới tính', 'Nam')),
                        "dan_toc": str(row.get('Dân tộc', 'Kinh')),
                        "tinh_thuong_tru": str(row.get('Tỉnh/Thành TT', '')),
                        "xa_thuong_tru": str(row.get('Phường/Xã TT', '')),
                        "dia_chi": str(row.get('Địa chỉ', ''))
                    })
                return {"status": "success", "filename": file.filename, "total": len(df),
                        "valid_count": len(valid_rows), "error_count": len(error_rows), "valid_data": valid_rows,
                        "error_data": error_rows}
            except Exception as e:
                return {"status": "error", "message": f"Lỗi đọc file Excel: {str(e)}"}, 400

        elif action == 'confirm_import':
            valid_data_str = request.form.get('valid_data')
            try:
                valid_data = json.loads(valid_data_str)
                if not valid_data:
                    flash('Không có dữ liệu hợp lệ để nhập!', 'error')
                    return redirect(url_for('khoitao_hocsinh'))

                for item in valid_data:
                    class_id = item.pop('class_id')
                    item["ma_dinh_danh"] = f"HS{datetime.now().year}{random.randint(10000, 99999)}"
                    
                    res_hs = supabase.table("students").insert(item).execute()
                    
                    if res_hs.data:
                        supabase.table("student_enrollments").insert({
                            "student_id": res_hs.data[0]['student_id'],
                            "class_id": class_id,
                            "academic_year_id": current_year_id
                        }).execute()
                
                flash(f'Đã nhập thành công {len(valid_data)} học sinh!', 'success')
            except Exception as e:
                flash(f'Lỗi khi lưu vào DB: {str(e)}', 'error')
            return redirect(url_for('khoitao_hocsinh'))

    # --- LOGIC GET ---
    f_grade = request.args.get('grade_id')
    f_class = request.args.get('class_id')

    grades = supabase.table("grades").select("*").order("grade_number").execute().data
    
    classes_query = supabase.table("classes").select("*").order("class_name")
    if current_year_id:
        classes_query = classes_query.eq("academic_year_id", current_year_id)
    classes = classes_query.execute().data

    students = []
    if f_class:
        enrollment_query = supabase.table("student_enrollments").select("student_id, students(*)").eq("class_id", f_class)
        if current_year_id:
            enrollment_query = enrollment_query.eq("academic_year_id", current_year_id)
        
        raw_data = enrollment_query.execute().data
        if raw_data:
            students = [item['students'] for item in raw_data if item.get('students')]

    return render_template(
        'admin/danhsach_hocsinh.html',
        grades=grades,
        classes=classes,
        students=students,
        f_grade=f_grade,
        f_class=f_class
    )


@app.route('/admin/khoitao_hocsinh/export')
def export_khoitao_hocsinh():
    if 'user_id' not in session or session.get('role_name') != 'Admin':
        return redirect(url_for('login'))

    f_class = request.args.get('class_id')

    year_res = supabase.table("academic_years").select("year_id").eq("is_current", True).execute()
    current_year_id = year_res.data[0]['year_id'] if year_res.data else None

    students = []
    if f_class:
        enrollment_query = supabase.table("student_enrollments").select("student_id, students(*)").eq("class_id", f_class)
        if current_year_id:
            enrollment_query = enrollment_query.eq("academic_year_id", current_year_id)
        raw_data = enrollment_query.execute().data
        if raw_data:
            students = [item['students'] for item in raw_data if item.get('students')]

    export_data = [{
        "Mã định danh": s.get("ma_dinh_danh"), "Họ tên": s.get("ho_ten"),
        "Giới tính": s.get("gioi_tinh"), "Ngày sinh": s.get("ngay_sinh"),
    } for s in students]

    df = pd.DataFrame(export_data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='DanhSachHocSinh')
    output.seek(0)

    return send_file(output, download_name="Danh_sach_hoc_sinh_khoi_tao.xlsx", as_attachment=True)


# =========================================================
# PHÂN PHỐI MÔN HỌC
# =========================================================
@app.route('/admin/phanphoi_monhoc', methods=['GET', 'POST'])
def phanphoi_monhoc():
    if 'user_id' not in session or session.get('role_name') != 'Admin':
        return redirect(url_for('login'))

    msg = request.args.get('msg')
    error_msg = request.args.get('error_msg')

    # --- 1. KHAI BÁO & TÍNH TOÁN QUỸ THỜI GIAN (LUÔN CHẠY) ---
    # Lấy danh sách 5 Khối
    grades = supabase.table("grades").select("*").order("grade_number").execute().data

    # Giả định tổng số tiết tối đa theo quy định là 35 tiết/tuần
    TOTAL_MAX_PERIODS = 35
    grade_remaining_slots = {}

    for g in grades:
        # Tính tổng tiết đã dùng của các môn hiện có trong khối này
        used_res = supabase.table("subjects").select("periods_per_week").eq("grade_id", g['grade_id']).execute().data
        total_used = sum(item['periods_per_week'] for item in used_res)
        # Quỹ còn lại = Tổng quy định - Đã dùng (tối thiểu là 0)
        grade_remaining_slots[g['grade_id']] = max(0, TOTAL_MAX_PERIODS - total_used)

    # Lấy Khối đang chọn ở bộ lọc (mặc định khối 1)
    selected_grade_id = request.args.get('grade_id')
    if not selected_grade_id and grades:
        selected_grade_id = str(grades[0]['grade_id'])

    # --- 2. XỬ LÝ CÁC HÀNH ĐỘNG POST ---
    if request.method == 'POST':
        action = request.form.get('action')

        # Đồng bộ môn học chuẩn
        if action == 'sync_subjects':
            for g in grades:
                g_id = g['grade_id']
                check = supabase.table("subjects").select("subject_id").eq("grade_id", g_id).eq("subject_name",
                                                                                                "Toán").execute().data
                if not check:
                    std_subs = [
                        {"grade_id": g_id, "subject_code": f"TOAN{g_id}", "subject_name": "Toán", "type": "Bắt buộc",
                         "evaluation_type": "Chấm điểm", "periods_per_week": 5, "status": "Đang áp dụng"},
                        {"grade_id": g_id, "subject_code": f"TV{g_id}", "subject_name": "Tiếng Việt",
                         "type": "Bắt buộc", "evaluation_type": "Chấm điểm", "periods_per_week": 10,
                         "status": "Đang áp dụng"}
                    ]
                    supabase.table("subjects").insert(std_subs).execute()
            return redirect(url_for('phanphoi_monhoc', grade_id=selected_grade_id, msg="Đồng bộ thành công!"))

        # Thêm môn tự chọn cho nhiều khối
        elif action == 'add_custom_subject':
            sub_name = request.form.get('subject_name')
            eval_type = request.form.get('evaluation_type')
            status = request.form.get('status')
            sel_grades = request.form.getlist('selected_grades')

            for g_id in sel_grades:
                periods = request.form.get(f'periods_grade_{g_id}')
                if periods and int(periods) > 0:
                    supabase.table("subjects").insert({
                        "subject_name": sub_name, "grade_id": g_id, "type": "Tự chọn",
                        "evaluation_type": eval_type, "periods_per_week": int(periods),
                        "status": status, "subject_code": f"TC_{g_id}_{datetime.now().strftime('%S%f')}"
                    }).execute()
            return redirect(url_for('phanphoi_monhoc', grade_id=selected_grade_id, msg="Đã thêm môn tự chọn!"))

    # --- 3. LẤY DỮ LIỆU BẢNG HIỂN THỊ ---
    subjects = []
    if selected_grade_id:
        subjects = supabase.table("subjects").select("*").eq("grade_id", selected_grade_id).order("type").execute().data

    return render_template(
        'admin/phanphoi_monhoc.html',
        grades=grades,
        subjects=subjects,
        selected_grade_id=selected_grade_id,
        grade_remaining_slots=grade_remaining_slots,  # Biến này giờ chắc chắn đã có giá trị
        msg=msg, error_msg=error_msg
    )


# =========================================================
# API: TẢI FILE EXCEL MẪU HỌC SINH
# =========================================================
@app.route('/admin/hocsinh/taimau')
def tai_mau_hocsinh():
    # Tạo dữ liệu mẫu chuẩn để người dùng nhập liệu[cite: 8]
    data = {
        'Khối': ['Khối 1', 'Khối 1'],
        'Tên lớp': ['1A', '1B'],
        'Họ tên': ['Nguyễn Văn A', 'Trần Thị B'],
        'Ngày sinh': ['2019-01-15', '2019-05-20'],
        'Giới tính': ['Nam', 'Nữ'],
        'Dân tộc': ['Kinh', 'Kinh'],
        'Tỉnh/Thành TT': ['Khánh Hòa', 'Hà Nội'],
        'Phường/Xã TT': ['Phường Lộc Thọ', 'Phường Tràng Tiền'],
        'Địa chỉ': ['Số 1 Trần Phú', 'Số 2 Đinh Tiên Hoàng']
    }
    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data_Mau')
    output.seek(0)
    return send_file(output, download_name="Mau_Nhap_HocSinh.xlsx", as_attachment=True)


# =========================================================
# QUẢN LÝ HỒ SƠ HỌC SINH TỔNG HỢP
# =========================================================
@app.route('/admin/hocsinh', methods=['GET', 'POST'])
def quanly_hocsinh():
    if 'user_id' not in session or session.get('role_name') != 'Admin':
        return redirect(url_for('login'))

    # BỔ SUNG: Lấy ID của năm học hiện hành ngay từ đầu
    year_res = supabase.table("academic_years").select("year_id").eq("is_current", True).execute()
    current_year_id = year_res.data[0]['year_id'] if year_res.data else None

    # 1. TÍNH NĂNG GET: HIỂN THỊ & BỘ LỌC
    grades = supabase.table("grades").select("*").order("grade_number").execute().data
    
    # FIX LỖI: Chỉ lấy các lớp thuộc năm học hiện tại để đưa vào ô Dropdown chọn lớp
    classes_query = supabase.table("classes").select("*").order("class_name")
    if current_year_id:
        classes_query = classes_query.eq("academic_year_id", current_year_id)
    classes = classes_query.execute().data

    f_grade = request.args.get('grade_id')
    f_class = request.args.get('class_id')
    f_name = request.args.get('search_name', '')

    # 2. TÍNH NĂNG POST: XỬ LÝ CÁC NGHIỆP VỤ[cite: 8]
    if request.method == 'POST':
        action = request.form.get('action')

        # --- A. THÊM MỚI & CẬP NHẬT ---[cite: 8]
        if action == 'save':
            try:
                s_id = request.form.get('student_id')
                class_id = request.form.get('class_id')
                data = {
                    "ho_ten": request.form.get('ho_ten'),
                    "ngay_sinh": request.form.get('ngay_sinh') or None,
                    "gioi_tinh": request.form.get('gioi_tinh'),
                    "dan_toc": request.form.get('dan_toc', 'Kinh'),
                    "tinh_thuong_tru": request.form.get('tinh_thuong_tru'),
                    "xa_thuong_tru": request.form.get('xa_thuong_tru'),
                    "dia_chi": request.form.get('dia_chi'),
                    "tinh_que_quan": request.form.get('tinh_que_quan'),
                    "xa_que_quan": request.form.get('xa_que_quan'),
                    "ten_phu_huynh": request.form.get('ten_phu_huynh'),
                    "quan_he_ph": request.form.get('quan_he_ph'),
                    "nghe_nghiep_ph": request.form.get('nghe_nghiep_ph'),
                    "sdt_ph": request.form.get('sdt_ph')
                }

                if s_id:  # Cập nhật[cite: 8]
                    supabase.table("students").update(data).eq("student_id", s_id).execute()
                    if class_id:
                        supabase.table("student_enrollments").update({"class_id": class_id}).eq("student_id",
                                                                                                s_id).execute()
                    flash('Đã cập nhật hồ sơ thành công!', 'success')
                else:  # Thêm mới[cite: 8]
                    data["ma_dinh_danh"] = f"HS2026{random.randint(10000, 99999)}"
                    res_hs = supabase.table("students").insert(data).execute()
                    if not res_hs.data: raise Exception("Lỗi tạo hồ sơ học sinh")

                    new_hs_id = res_hs.data[0]['student_id']
                    year_res = supabase.table("academic_years").select("year_id").eq("is_current", True).execute()
                    if not year_res.data: raise Exception("Chưa thiết lập năm học hiện hành")

                    supabase.table("student_enrollments").insert({
                        "student_id": new_hs_id, "class_id": class_id, "academic_year_id": year_res.data[0]['year_id']
                    }).execute()
                    flash('Đã thêm hồ sơ thành công!', 'success')
            except Exception as e:
                flash(f"Lỗi nghiệp vụ: {str(e)}", 'error')
            return redirect(url_for('quanly_hocsinh'))

        # --- B. XÓA HỌC SINH ---[cite: 8]
        elif action == 'delete':
            ids_str = request.form.get('student_ids')
            if not ids_str:
                flash("Vui lòng tick chọn ít nhất một học sinh!", 'warning')
            else:
                ids = ids_str.split(',')
                try:
                    supabase.table("student_enrollments").delete().in_("student_id", ids).execute()
                    supabase.table("students").delete().in_("student_id", ids).execute()
                    flash(f'Đã xóa vĩnh viễn {len(ids)} học sinh!', 'success')
                except Exception as e:
                    flash(f"Lỗi khi xoá: {str(e)}", 'error')
            return redirect(url_for('quanly_hocsinh'))

        # --- C. PREVIEW DATA TỪ EXCEL ---[cite: 7, 8]
        elif action == 'preview_import':
            file = request.files.get('excel_file')
            if not file: return {"status": "error", "message": "Vui lòng chọn file!"}, 400
            try:
                df = pd.read_excel(file)
                valid_rows, error_rows = [], []
                for index, row in df.iterrows():
                    ho_ten = str(row.get('Họ tên', '')).strip()
                    lop = str(row.get('Tên lớp', '')).strip().upper()
                    if not ho_ten or ho_ten == 'nan' or not lop or lop == 'nan':
                        error_rows.append({"row": index + 2, "name": ho_ten, "error": "Thiếu thông tin"})
                        continue
                    class_data = supabase.table('classes').select('class_id').eq('class_name', lop).execute().data
                    if not class_data:
                        error_rows.append({"row": index + 2, "name": ho_ten, "error": f"Lớp {lop} không tồn tại"})
                        continue
                    valid_rows.append({
                        "class_id": class_data[0]['class_id'], "ho_ten": ho_ten,
                        "ngay_sinh": str(row.get('Ngày sinh', ''))[:10] if str(
                            row.get('Ngày sinh', '')) != 'nan' else None,
                        "gioi_tinh": str(row.get('Giới tính', 'Nam')), "dan_toc": str(row.get('Dân tộc', 'Kinh')),
                        "tinh_thuong_tru": str(row.get('Tỉnh/Thành TT', '')),
                        "xa_thuong_tru": str(row.get('Phường/Xã TT', '')), "dia_chi": str(row.get('Địa chỉ', ''))
                    })
                return {"status": "success", "filename": file.filename, "total": len(df),
                        "valid_count": len(valid_rows), "error_count": len(error_rows), "valid_data": valid_rows,
                        "error_data": error_rows}
            except Exception as e:
                return {"status": "error", "message": str(e)}, 400

        # --- D. CONFIRM LƯU DATA EXCEL VÀO DB ---[cite: 7, 8]
        elif action == 'confirm_import':
            valid_data_str = request.form.get('valid_data')
            try:
                valid_data = json.loads(valid_data_str)
                year_res = supabase.table("academic_years").select("year_id").eq("is_current", True).execute()
                for item in valid_data:
                    c_id = item.pop('class_id')
                    item["ma_dinh_danh"] = f"HS2026{random.randint(10000, 99999)}"
                    res_hs = supabase.table("students").insert(item).execute()
                    if res_hs.data:
                        supabase.table("student_enrollments").insert(
                            {"student_id": res_hs.data[0]['student_id'], "class_id": c_id,
                             "academic_year_id": year_res.data[0]['year_id']}).execute()
                flash(f'Đã nhập thành công {len(valid_data)} học sinh!', 'success')
            except Exception as e:
                flash(f'Lỗi khi lưu: {str(e)}', 'error')
            return redirect(url_for('quanly_hocsinh'))

    # TRUY VẤN DANH SÁCH HIỂN THỊ[cite: 8]
    # 1. Lấy ID của năm học hiện hành đang active
    year_res = supabase.table("academic_years").select("year_id").eq("is_current", True).execute()
    current_year_id = year_res.data[0]['year_id'] if year_res.data else None

    # 2. Bổ sung trường academic_year_id vào chuỗi select của bảng phân lớp
    query = supabase.table("students").select(
        "*, student_enrollments(class_id, academic_year_id, classes(class_name, grade_id, grades(grade_name)))"
    )
    if f_name: 
        query = query.ilike("ho_ten", f"%{f_name}%")
        
    raw_data = query.execute().data
    students = []

    # 3. Lọc học sinh theo năm học hiện tại
    for s in raw_data:
        enroll = s.get('student_enrollments', [])
        if not enroll: 
            continue
            
        # Chỉ giữ lại các bản ghi xếp lớp trùng với ID của năm học hiện hành
        current_enroll = [e for e in enroll if e.get('academic_year_id') == current_year_id]
        if not current_enroll: 
            continue  # Nếu học sinh không có lịch sử phân lớp ở năm hiện hành -> Bỏ qua không hiển thị

        s_class_id = str(current_enroll[0]['class_id'])
        s_grade_id = str(current_enroll[0]['classes']['grade_id']) if current_enroll[0].get('classes') else ""
        
        # Tiếp tục áp dụng các bộ lọc Khối/Lớp tự chọn trên thanh công cụ UI
        if f_grade and s_grade_id != f_grade: 
            continue
        if f_class and s_class_id != f_class: 
            continue
            
        students.append(s)

    return render_template('admin/hocsinh.html', grades=grades, classes=classes, students=students, f_grade=f_grade,
                           f_class=f_class, f_name=f_name)


# =========================================================
# API: XUẤT FILE EXCEL THEO BỘ LỌC ĐẸP[cite: 8]
# =========================================================
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill


@app.route('/admin/hocsinh/export')
def export_hocsinh():
    if 'user_id' not in session or session.get('role_name') != 'Admin':
        return redirect(url_for('login'))

    f_grade = request.args.get('grade_id')
    f_class = request.args.get('class_id')
    f_name = request.args.get('search_name', '')

    # Truy vấn dữ liệu hệt như bộ lọc trên Web
    query = supabase.table("students").select(
        "*, student_enrollments(class_id, classes(class_name, grade_id, grades(grade_name)))")
    if f_name: query = query.ilike("ho_ten", f"%{f_name}%")
    raw_data = query.execute().data

    filtered_data = []
    title_text = "DANH SÁCH HỌC SINH TOÀN TRƯỜNG"

    for s in raw_data:
        enroll = s.get('student_enrollments', [])
        if not enroll: continue

        c_info = enroll[0].get('classes')
        s_class_id = str(enroll[0]['class_id'])
        s_grade_id = str(c_info['grade_id']) if c_info else ""
        s_class_name = c_info['class_name'] if c_info else ""
        s_grade_name = c_info['grades']['grade_name'] if c_info and c_info.get('grades') else ""

        if f_grade and s_grade_id != f_grade: continue
        if f_class and s_class_id != f_class: continue

        if f_class:
            title_text = f"DANH SÁCH HỌC SINH LỚP {s_class_name.upper()}"
        elif f_grade:
            title_text = f"DANH SÁCH HỌC SINH {s_grade_name.upper()}"

        filtered_data.append({
            "Mã định danh": s.get("ma_dinh_danh"),
            "Họ tên": s.get("ho_ten"),
            "Ngày sinh": s.get("ngay_sinh"),
            "Giới tính": s.get("gioi_tinh"),
            "Dân tộc": s.get("dan_toc"),
            "Khối": s_grade_name,
            "Lớp": s_class_name
        })

    df = pd.DataFrame(filtered_data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, startrow=2, sheet_name='HocSinh')
        ws = writer.sheets['HocSinh']

        # Trang trí tiêu đề (Dòng 1)
        ws.merge_cells('A1:G1')
        ws['A1'] = title_text
        ws['A1'].font = Font(size=14, bold=True, color="FF0000")
        ws['A1'].alignment = Alignment(horizontal="center")

        # Trang trí Header bảng (Dòng 3)
        header_fill = PatternFill(start_color="1162A8", end_color="1162A8", fill_type="solid")
        for cell in ws[3]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")

    output.seek(0)
    return send_file(output, download_name="Danh_sach_hoc_sinh.xlsx", as_attachment=True)


# ======================
# HỒ SƠ NHÂN SỰ
# ======================
@app.route('/admin/hosonhansu')
def hosonhansu():
    gioi_tinh = request.args.get('gioi_tinh')

    res = supabase.table("teachers") \
        .select("*") \
        .order("teacher_id", desc=False) \
        .execute()

    nhan_su_list = res.data

    if gioi_tinh:
        nhan_su_list = [
            ns for ns in nhan_su_list
            if ns.get("gioi_tinh") == gioi_tinh
        ]

    return render_template(
        'admin/hosonhansu.html',
        nhan_su_list=nhan_su_list,
        user_name=session.get('user_name')
    )

# ======================
# EXPORT EXCEL
# ======================
@app.route('/admin/export-excel')
def export_excel():
    res = supabase.table("teachers").select("*").execute()
    data = res.data

    result = []

    for ns in data:
        result.append({
            "Mã số nhân sự": ns.get("teacher_id"),
            "Họ tên": ns.get("full_name"),
            "Giới tính": ns.get("gioi_tinh"),
            "Ngày sinh": ns.get("ngay_sinh") or '',
            "SĐT": ns.get("phone_number"),
            "Email": ns.get("email"),
            "Địa chỉ": ns.get("dia_chi"),
            "CCCD": ns.get("cccd"),
            "Ngày vào làm": ns.get("ngay_vao_lam") or '',
            "Trạng thái": ns.get("trang_thai"),
            "Trình độ": ns.get("trinh_do"),
            "Lương": ns.get("he_so_luong"),
            "Ngày ký HĐ": ns.get("ngay_ky_hop_dong") or ''
        })

    df = pd.DataFrame(result)

    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return send_file(
        output,
        download_name="nhan_su.xlsx",
        as_attachment=True
    )

# =========================================================
# GÁN GIÁO VIÊN CHỦ NHIỆM (HIỂN THỊ VÀ CẬP NHẬT)
# =========================================================
@app.route('/admin/phancongchunhiem', methods=['GET', 'POST'])
def phancongchunhiem():

    # 1. Lấy tham số lọc từ Form hoặc URL
    if request.method == 'POST':
        selected_grade = request.form.get('grade_id')
        selected_year = request.form.get('year_id')
    else:
        selected_grade = request.args.get('grade_id')
        selected_year = request.args.get('year_id')

    # Nếu chọn "Tất cả" thì đưa về None
    if selected_grade == "":
        selected_grade = None

    if selected_year == "":
        selected_year = None

    # Nếu chưa có năm học, tự động lấy năm học hiện tại
    if not selected_year:
        current_year = supabase.table("academic_years") \
            .select("*") \
            .eq("is_current", True) \
            .execute().data

        if current_year:
            selected_year = current_year[0]["year_id"]

    error_msg = None

    # 2. XỬ LÝ LƯU DỮ LIỆU POST
    if request.method == 'POST':
        is_update_action = any(key.startswith("teacher_") for key in request.form)

        if is_update_action:
            try:
                target_year = int(request.form.get("year_id"))

                for key, teacher_id in request.form.items():
                    if key.startswith("teacher_"):
                        class_id = int(key.split("_")[1])

                        supabase.table("homeroom_assignments") \
                            .delete() \
                            .eq("class_id", class_id) \
                            .eq("year_id", target_year) \
                            .execute()

                        if teacher_id:
                            supabase.table("homeroom_assignments").insert({
                                "class_id": class_id,
                                "teacher_id": int(teacher_id),
                                "year_id": target_year
                            }).execute()

                return redirect(url_for(
                    'phancongchunhiem',
                    year_id=target_year,
                    grade_id=selected_grade if selected_grade else "",
                    msg="Cập nhật thành công!"
                ))

            except Exception as e:
                error_msg = f"Lỗi hệ thống: {str(e)}"

    # 3. TRUY VẤN DỮ LIỆU ĐỂ HIỂN THỊ
    grades = supabase.table("grades") \
        .select("*") \
        .order("grade_number") \
        .execute().data

    academic_years = supabase.table("academic_years") \
        .select("*") \
        .order("year_id") \
        .execute().data

    teachers = supabase.table("teachers") \
        .select("teacher_id, full_name") \
        .order("teacher_id") \
        .execute().data

    class_query = supabase.table("classes").select("*")

    if selected_grade:
        class_query = class_query.eq("grade_id", int(selected_grade))

    if selected_year:
        class_query = class_query.eq("academic_year_id", int(selected_year))

    classes = class_query.order("class_name").execute().data

    assignments = []

    if selected_year:
        assignments = supabase.table("homeroom_assignments") \
            .select("*") \
            .eq("year_id", int(selected_year)) \
            .execute().data

    return render_template(
        "admin/phancongchunhiem.html",
        classes=classes,
        teachers=teachers,
        academic_years=academic_years,
        grades=grades,
        assignments=assignments,
        selected_grade=int(selected_grade) if selected_grade else None,
        selected_year=int(selected_year) if selected_year else None,
        error=error_msg,
        user_name=session.get('user_name')
    )

# =========================================================
# XUẤT FILE EXCEL PHÂN CÔNG CHỦ NHIỆM
# =========================================================
@app.route('/admin/export-phancong')
def export_excel_phancong():
    # 1. Lấy tham số lọc từ URL (Học kỳ và Khối đang chọn)
    selected_year = request.args.get('year_id', type=int)
    selected_grade = request.args.get('grade_id', type=int)

    # 2. Xây dựng câu truy vấn bảng homeroom_assignments
    query = supabase.table("homeroom_assignments").select("*, classes(class_name, grade_id), teachers(teacher_id, full_name)")

    # Lọc theo năm học
    if selected_year:
        query = query.eq("year_id", selected_year)

    data = query.execute().data

    # 3. Lọc tiếp theo khối bằng Python
    if selected_grade:
        data = [row for row in data if row.get("classes") and row["classes"].get("grade_id") == selected_grade]

    # 4. Trích xuất dữ liệu
    result = []
    for row in data:
        class_info = row.get("classes") or {}
        teacher_info = row.get("teachers") or {}

        result.append({
            "Lớp": class_info.get("class_name", "Không rõ"),
            "Mã GV": teacher_info.get("teacher_id", ""),
            "Giáo viên chủ nhiệm": teacher_info.get("full_name", "Chưa phân công")
        })

    # 5. Chuyển thành DataFrame và sắp xếp theo tên lớp (A-Z)
    df = pd.DataFrame(result)
    if not df.empty:
        df.sort_values(by="Lớp", inplace=True)

    # 6. Xuất file
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)

    return send_file(output, download_name="phan_cong_chu_nhiem.xlsx", as_attachment=True)

# ======================
# PHÂN CÔNG CHUYÊN MÔN (MÀN 1)
# ======================
@app.route('/admin/phancongchuyenmon')
def phancongchuyenmon():

    selected_semester = request.args.get("semester_id", type=int)

    # lấy năm học hiện tại
    current_year = supabase.table("academic_years") \
        .select("year_id") \
        .eq("is_current", True) \
        .single() \
        .execute().data

    current_year_id = current_year["year_id"]

    # nếu chưa chọn học kỳ → lấy học kỳ hiện tại
    if selected_semester:
        semester_id = selected_semester
    else:
        semester_id = get_current_semester()

    # lấy danh sách giáo viên THEO THỨ TỰ teacher_id
    teachers = supabase.table("teachers") \
        .select("teacher_id, full_name") \
        .order("teacher_id") \
        .execute().data

    # lấy phân công chuyên môn theo học kỳ đang chọn
    assignments = supabase.table("teaching_assignments") \
        .select("*, classes(class_name), subjects(subject_name)") \
        .eq("semester_id", semester_id) \
        .execute().data

    # gom dữ liệu môn dạy theo giáo viên
    teacher_map = {}

    for row in assignments:
        tid = row["teacher_id"]
        text = f'{row["subjects"]["subject_name"]} ({row["classes"]["class_name"]})'
        teacher_map.setdefault(tid, []).append(text)

    # chỉ lấy danh sách học kỳ của năm học hiện tại
    semesters = supabase.table("semesters") \
        .select("*") \
        .eq("year_id", current_year_id) \
        .order("semester_id") \
        .execute().data

    return render_template(
        "admin/phancongchuyenmon.html",
        teachers=teachers,
        teacher_map=teacher_map,
        semesters=semesters,
        selected_semester=semester_id
    )

# ======================
# PHÂN CÔNG CHUYÊN MÔN (MÀN 2)
# ======================
@app.route('/admin/phancongchuyenmon/edit/<int:teacher_id>', methods=['GET', 'POST'])
def edit_pcm(teacher_id):
    teachers = supabase.table("teachers").select("teacher_id, full_name").execute().data
    teacher = supabase.table("teachers").select("*").eq("teacher_id", teacher_id).execute().data
    teacher = teacher[0] if teacher else None

    semester_id = request.args.get("semester_id", type=int)
    if not semester_id:
        semester_id = get_current_semester()

    # ================= POST (LƯU DỮ LIỆU) =================
    if request.method == 'POST':
        # ❌ xoá phân công cũ
        supabase.table("teaching_assignments") \
            .delete() \
            .eq("teacher_id", teacher_id) \
            .eq("semester_id", semester_id) \
            .execute()

        # ✅ thêm mới
        for key in request.form:
            if key.startswith("assign_"):
                _, class_id, subject_id = key.split("_")
                supabase.table("teaching_assignments").insert({
                    "teacher_id": teacher_id,
                    "class_id": int(class_id),
                    "subject_id": int(subject_id),
                    "semester_id": semester_id
                }).execute()

        return redirect(url_for('phancongchuyenmon', semester_id=semester_id))

    # ================= GET (HIỂN THỊ DỮ LIỆU) =================

    # 1. Lấy Khối (Nếu để trống thì là Tất cả)
    selected_grade_str = request.args.get("grade_id")
    selected_grade = int(selected_grade_str) if selected_grade_str and selected_grade_str.isdigit() else None

    # Lấy danh sách Khối
    all_grades = supabase.table("grades").select("*").order("grade_number").execute().data

    # Lọc khối nếu có chọn, nếu không thì lấy tất cả
    grades_to_process = [g for g in all_grades if g["grade_id"] == selected_grade] if selected_grade else all_grades

    # 2. Gom dữ liệu theo từng khối để đẩy ra HTML không bị lỗi
    grade_data = []
    for g in grades_to_process:
        # Lấy academic_year_id từ học kỳ đang chọn
        sem_info = supabase.table("semesters").select("year_id").eq("semester_id", semester_id).execute().data
        year_id = sem_info[0]['year_id'] if sem_info else None

        c_list = supabase.table("classes").select("*") \
            .eq("grade_id", g["grade_id"]) \
            .eq("academic_year_id", year_id) \
            .order("class_name").execute().data if year_id else []
        s_list = supabase.table("subjects").select("*").eq("grade_id", g["grade_id"]).execute().data

        # Chỉ hiển thị khối nào có lớp và môn học
        if c_list and s_list:
            grade_data.append({
                "grade": g,
                "classes": c_list,
                "subjects": s_list
            })

    # 3. Lấy phân công để bôi đỏ (Giữ nguyên logic cũ)
    all_assignments = supabase.table("teaching_assignments") \
        .select("*, teachers(full_name)") \
        .eq("semester_id", semester_id) \
        .execute().data

    assigned_map = {}
    if all_assignments:
        for a in all_assignments:
            key = f"{a['class_id']}_{a['subject_id']}"
            assigned_map[key] = {
                "teacher_id": a["teacher_id"],
                "teacher_name": a["teachers"]["full_name"] if a.get("teachers") else "Không rõ"
            }

    return render_template(
        "admin/phancongchuyenmon_edit.html",
        grade_data=grade_data,  # <-- Truyền list đã gom nhóm
        teachers=teachers,
        teacher=teacher,
        assigned_map=assigned_map,
        semester_id=semester_id,
        selected_grade=selected_grade
    )


@app.route('/admin/phancongchuyenmon/by-class')
def pcm_by_class():
    selected_semester = request.args.get("semester_id", type=int)
    selected_grade = request.args.get("grade_id", type=int)
    selected_class = request.args.get("class_id", type=int)

    # lấy năm học hiện tại
    current_year = supabase.table("academic_years") \
        .select("year_id") \
        .eq("is_current", True) \
        .single() \
        .execute().data

    current_year_id = current_year["year_id"]

    # chỉ lấy danh sách học kỳ của năm học hiện tại
    semesters = supabase.table("semesters") \
        .select("*") \
        .eq("year_id", current_year_id) \
        .order("semester_id") \
        .execute().data

    # nếu chưa chọn học kỳ → lấy học kỳ hiện tại
    if not selected_semester:
        for sem in semesters:
            if sem.get("is_current"):
                selected_semester = sem["semester_id"]
                break

    # nếu năm hiện tại chưa có học kỳ is_current thì lấy học kỳ đầu tiên
    if not selected_semester and semesters:
        selected_semester = semesters[0]["semester_id"]

    # lấy khối
    grades = supabase.table("grades") \
        .select("*") \
        .order("grade_id") \
        .execute().data

    # lấy lớp theo khối và theo năm học hiện tại
    class_query = supabase.table("classes") \
        .select("*") \
        .eq("academic_year_id", current_year_id)

    if selected_grade:
        class_query = class_query.eq("grade_id", selected_grade)

    classes = class_query.order("class_name").execute().data

    data = []

    # truy vấn dữ liệu phân công chuyên môn
    if selected_class and selected_semester:
        rows = supabase.table("teaching_assignments") \
            .select("*, teachers(teacher_id, full_name), subjects(subject_name)") \
            .eq("class_id", selected_class) \
            .eq("semester_id", selected_semester) \
            .execute().data

        teacher_map = {}

        for r in rows:
            tid = r["teacher_id"]

            teacher_info = r.get("teachers") or {}
            subject_info = r.get("subjects") or {}

            name = teacher_info.get("full_name", "Không rõ")
            subject = subject_info.get("subject_name", "Không rõ")

            if tid not in teacher_map:
                teacher_map[tid] = {
                    "teacher_id": tid,
                    "name": name,
                    "subjects": []
                }

            teacher_map[tid]["subjects"].append(subject)

        for i, t in enumerate(teacher_map.values(), start=1):
            data.append({
                "stt": i,
                "teacher_id": t["teacher_id"],
                "name": t["name"],
                "subjects": ", ".join(t["subjects"])
            })

    return render_template(
        "admin/phancongchuyenmon_by_class.html",
        grades=grades,
        classes=classes,
        data=data,
        selected_grade=selected_grade,
        selected_class=selected_class,
        selected_semester=selected_semester,
        semesters=semesters
    )


@app.route('/admin/phancongchuyenmon/export')
def export_pcm():
    semester_id = request.args.get("semester_id", type=int) or get_current_semester()

    assignments = supabase.table("teaching_assignments") \
        .select("*, classes(class_name), subjects(subject_name)") \
        .eq("semester_id", semester_id) \
        .execute().data

    teachers = supabase.table("teachers") \
        .select("teacher_id, full_name") \
        .execute().data

    teacher_map = {}
    for row in assignments:
        tid = row["teacher_id"]
        text = f'{row["subjects"]["subject_name"]} ({row["classes"]["class_name"]})'
        teacher_map.setdefault(tid, []).append(text)

    # 🔥 FIX LỖI Ở ĐÂY: Khởi tạo Workbook và Worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Phân công chuyên môn"

    # Bây giờ các lệnh ws.append sẽ không còn dấu đỏ nữa
    ws.append(["STT", "Mã GV", "Tên giáo viên", "Môn dạy"])

    for i, t in enumerate(teachers, start=1):
        subjects = ", ".join(teacher_map.get(t["teacher_id"], []))
        ws.append([
            i,
            t["teacher_id"],
            t["full_name"],
            subjects
        ])

    # Lưu vào memory
    file_stream = io.BytesIO()
    wb.save(file_stream)  # Hết dấu đỏ ở wb
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name="phan_cong_chuyen_mon.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route('/admin/phancongchuyenmon/by-class/export')
def export_pcm_by_class():
    semester_id = request.args.get("semester_id", type=int) or get_current_semester()

    grade_id = request.args.get("grade_id")
    class_id = request.args.get("class_id")

    # ================= QUERY =================
    query = supabase.table("teaching_assignments") \
        .select("*, teachers(full_name), subjects(subject_name), classes(class_name)") \
        .eq("semester_id", semester_id)

    if class_id:
        query = query.eq("class_id", int(class_id))

    rows = query.execute().data

    # ================= GOM DATA =================
    teacher_map = {}

    for r in rows:
        tid = r["teacher_id"]
        name = r["teachers"]["full_name"]
        subject = f'{r["subjects"]["subject_name"]} ({r["classes"]["class_name"]})'

        if tid not in teacher_map:
            teacher_map[tid] = {
                "teacher_id": tid,
                "name": name,
                "subjects": []
            }

        teacher_map[tid]["subjects"].append(subject)

    # ================= TẠO EXCEL =================
    wb = Workbook()
    ws = wb.active
    ws.title = "PCCM"

    # header
    headers = ["STT", "Mã GV", "Họ tên", "Môn dạy"]
    ws.append(headers)

    # 🔥 in đậm header
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # data
    for i, t in enumerate(teacher_map.values(), start=1):
        ws.append([
            i,
            t["teacher_id"],
            t["name"],
            ", ".join(t["subjects"])
        ])

    # 🔥 auto width cột
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter

        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass

        ws.column_dimensions[col_letter].width = max_length + 2

    # ================= TÊN FILE =================
    class_name = "tat_ca"

    if class_id:
        cls = supabase.table("classes") \
            .select("class_name") \
            .eq("class_id", int(class_id)) \
            .execute().data

        if cls:
            class_name = cls[0]["class_name"]

    today = datetime.now().strftime("%d-%m-%Y")
    filename = f"pccm_{class_name}_{today}.xlsx"

    # ================= TRẢ FILE =================
    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


def get_current_semester():
    sem = supabase.table("semesters") \
        .select("*") \
        .eq("is_current", True) \
        .execute().data

    return sem[0]["semester_id"] if sem else None


@app.route('/admin/phancongchuyenmon/import')
def import_pcm_page():
    error_msg = request.args.get("error_msg")  # 🔥 THÊM

    grades = supabase.table("grades").select("*").execute().data
    subjects = supabase.table("subjects").select("*").execute().data

    subject_map = {}
    for s in subjects:
        subject_map.setdefault(s["grade_id"], []).append(s)

    return render_template(
        "admin/import_pcm.html",
        grades=grades,
        subject_map=subject_map,
        error_msg=error_msg  # 🔥 THÊM
    )


@app.route('/admin/phancongchuyenmon/template')
def download_template_pcm():
    wb = load_workbook("template_pcm.xlsx")
    ws = wb.active

    teachers = supabase.table("teachers").select("*").execute().data

    start_row = 7  # dòng bắt đầu bảng

    # 🔥 tạo border chuẩn
    thin = Side(style='thin')
    full_border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    # 🔥 XÓA dòng cũ
    ws.delete_rows(start_row, 100)

    # 🔥 TẠO DÒNG MỚI
    for idx, t in enumerate(teachers, start=1):

        row = start_row + idx - 1

        ws.cell(row=row, column=1).value = idx
        ws.cell(row=row, column=2).value = t["teacher_id"]
        ws.cell(row=row, column=3).value = t["full_name"]

        # 👉 kẻ bảng full
        for col in range(1, 5):
            ws.cell(row=row, column=col).border = full_border

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)

    return send_file(
        stream,
        as_attachment=True,
        download_name="template_pcm.xlsx"
    )


@app.route('/admin/phancongchuyenmon/upload', methods=['POST'])
def upload_pcm_excel():
    file = request.files.get('file')

    if not file or file.filename == '':
        return redirect(url_for('import_pcm_page', error_msg="Chưa chọn file hoặc file rỗng!"))

    try:
        # Đọc dữ liệu từ file Excel, bỏ qua 5 dòng đầu theo mẫu
        df = pd.read_excel(file, skiprows=5)
        df.columns = df.columns.astype(str).str.strip()

        if "Mã giáo viên" not in df.columns:
            return redirect(
                url_for('import_pcm_page', error_msg='File không đúng chuẩn mẫu: Không tìm thấy cột "Mã giáo viên".'))

        df = df.dropna(subset=['Mã giáo viên'])
    except Exception as e:
        return redirect(url_for('import_pcm_page', error_msg=f"Không thể đọc dữ liệu file: {str(e)}"))

    # 1. Lấy dữ liệu cơ bản từ Database
    subjects = supabase.table("subjects").select("*").execute().data
    classes = supabase.table("classes").select("*").execute().data
    sem = supabase.table("semesters").select("semester_id").eq("is_current", True).execute().data
    semester_id = sem[0]["semester_id"] if sem else None

    class_info_map = {c["class_name"]: {"class_id": c["class_id"], "grade_id": c["grade_id"]} for c in classes}
    subject_map = {(s["grade_id"], s["subject_name"]): s["subject_id"] for s in subjects}

    # 2. Kiểm tra trùng lặp trong Database
    existing_assignments = supabase.table("teaching_assignments") \
        .select("class_id, subject_id, teachers(full_name)") \
        .eq("semester_id", semester_id) \
        .execute().data

    assigned_db_map = {}
    for a in existing_assignments:
        if a.get("teachers"):
            assigned_db_map[(a["class_id"], a["subject_id"])] = a["teachers"]["full_name"]

    results = []
    valid_assignments = []
    current_file_map = {}

    # 3. Duyệt dữ liệu
    for _, row in df.iterrows():
        try:
            teacher_id = int(row.get("Mã giáo viên"))
        except:
            continue

        teacher_name = str(row.get("Họ & Tên", row.get("Họ tên giáo viên", "Không rõ"))).strip()
        subject_text = str(row.get("Môn dạy", "")).strip()

        if not subject_text or subject_text == 'nan' or "(" not in subject_text:
            results.append({
                "teacher_id": teacher_id, "teacher_name": teacher_name,
                "subject_name": subject_text if subject_text != 'nan' else 'Trống',
                "status": "error", "message": "Sai định dạng. VD: Toán (1A)"
            })
            continue

        pairs = re.findall(r"(.+?)\s*\((.+?)\)", subject_text)
        error_msg = ""
        row_parsed_data = []

        for sub_name, class_name in pairs:
            sub_name = sub_name.strip()
            class_name = class_name.strip()

            if class_name not in class_info_map:
                error_msg += f"Lớp '{class_name}' không tồn tại; "
                continue

            c_info = class_info_map[class_name]
            c_id = c_info["class_id"]
            grade_id = c_info["grade_id"]
            key_subject = (grade_id, sub_name)

            if key_subject not in subject_map:
                error_msg += f"Môn '{sub_name}' không thuộc khối của lớp {class_name}; "
            else:
                s_id = subject_map[key_subject]
                if (c_id, s_id) in assigned_db_map:
                    gv_cu = assigned_db_map[(c_id, s_id)]
                    error_msg += f"Trùng DB: Đã gán cho GV '{gv_cu}'; "
                elif (c_id, s_id) in current_file_map:
                    gv_truoc = current_file_map[(c_id, s_id)]
                    error_msg += f"Trùng File: Đã gán cho '{gv_truoc}'; "
                else:
                    current_file_map[(c_id, s_id)] = teacher_name
                    row_parsed_data.append({
                        "teacher_id": teacher_id, "class_id": c_id, "subject_id": s_id, "semester_id": semester_id
                    })

        if error_msg:
            results.append({"teacher_id": teacher_id, "teacher_name": teacher_name, "subject_name": subject_text,
                            "status": "error", "message": error_msg})
        else:
            results.append({"teacher_id": teacher_id, "teacher_name": teacher_name, "subject_name": subject_text,
                            "status": "success", "message": ""})
            valid_assignments.extend(row_parsed_data)

    # Lưu vào Database
    if valid_assignments and semester_id:
        try:
            supabase.table("teaching_assignments").insert(valid_assignments).execute()
        except Exception as e:
            return redirect(url_for('import_pcm_page', error_msg=f"Lỗi DB: {str(e)}"))

    # 🔥 LƯU KẾT QUẢ VÀO SESSION ĐỂ TẢI XUỐNG[cite: 7]
    session['temp_import_results'] = results

    return render_template("admin/import_result.html", results=results)


@app.route('/admin/import-result/export')
def export_import_result():
    results = session.get('temp_import_results', [])
    if not results:
        return "Không có dữ liệu!", 400

    df = pd.DataFrame(results)
    column_map = {"teacher_id": "Mã giáo viên", "teacher_name": "Họ & Tên", "subject_name": "Môn dạy",
                  "status": "Trạng thái", "message": "Ghi chú lỗi"}
    df = df.rename(columns=column_map)
    df['Trạng thái'] = df['Trạng thái'].map({'success': 'Thành công', 'error': 'Thất bại'})

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='KetQuaNhapLieu')
    output.seek(0)

    return send_file(output, download_name="Ket_qua_nhap_lieu_PCCM.xlsx", as_attachment=True)


# ======================
# HẾT PHẦN NHÂN SỰ
# ======================
# ======================
# TEMPLATE FILTERS
# ======================
@app.template_filter('format_date')
def format_date_filter(value):
    if not value:
        return ''
    try:
        from datetime import datetime
        return datetime.strptime(str(value)[:10], '%Y-%m-%d').strftime('%d/%m/%Y')
    except:
        return str(value)


# ======================
# THỜI KHÓA BIỂU
# ======================
# ======================
# KHAI BÁO TUẦN
# ======================
@app.route('/admin/khaibao_tuan', methods=['GET', 'POST'])
def khaibao_tuan():
    if 'user_id' not in session or session.get('role_name') != 'Admin':
        return redirect(url_for('login'))

    msg = request.args.get('msg')
    error_msg = request.args.get('error_msg')

    # ── LẤY HỌC KỲ & FILTER ──────────────────────────────────────
    semesters = supabase.table("semesters").select("*").order("semester_id").execute().data
    years_raw = supabase.table("academic_years").select("year_id, year_name").execute().data
    year_map = {y['year_id']: y['year_name'] for y in years_raw}
    for s in semesters:
        year_name = year_map.get(s.get('year_id'), '')
        s['display_name'] = f"{s['semester_name']} - {year_name}" if year_name else s['semester_name']

    # Học kỳ được chọn (ưu tiên query param → học kỳ hiện tại → đầu tiên)
    selected_semester_id = request.args.get('semester_id', type=int)
    if not selected_semester_id:
        cur = next((s for s in semesters if s.get('is_current')), None)
        selected_semester_id = cur['semester_id'] if cur else (semesters[0]['semester_id'] if semesters else None)

    # ── XỬ LÝ POST ────────────────────────────────────────────────
    if request.method == 'POST':
        action = request.form.get('action')
        sem_filter = request.form.get('semester_filter', '')

        redirect_url = url_for('khaibao_tuan', semester_id=sem_filter or selected_semester_id)

        # ── THÊM / SỬA TUẦN ──────────────────────────────────────
        if action in ('add', 'edit'):
            semester_id = request.form.get('semester_id', type=int)
            start_date = request.form.get('start_date')
            end_date = request.form.get('end_date')
            week_number = request.form.get('week_number', type=int)
            is_visible = bool(request.form.get('is_visible'))
            is_exam_week = bool(request.form.get('is_exam_week'))
            exam_week_name = request.form.get('exam_week_name', '').strip() or None

            if not (semester_id and start_date and end_date and week_number):
                return redirect(url_for('khaibao_tuan',
                                        semester_id=sem_filter,
                                        error_msg='Vui lòng điền đầy đủ các trường bắt buộc!'))

            payload = {
                'semester_id': semester_id,
                'start_date': start_date,
                'end_date': end_date,
                'week_number': week_number,
                'is_visible': is_visible,
                'is_exam_week': is_exam_week,
                'exam_week_name': exam_week_name,
            }

            try:
                if action == 'add':
                    supabase.table("school_weeks").insert(payload).execute()
                    return redirect(url_for('khaibao_tuan',
                                            semester_id=semester_id,
                                            msg='Thêm tuần học thành công!'))
                else:  # edit
                    week_id = request.form.get('week_id', type=int)
                    supabase.table("school_weeks").update(payload).eq('week_id', week_id).execute()
                    return redirect(url_for('khaibao_tuan',
                                            semester_id=sem_filter or semester_id,
                                            msg='Cập nhật tuần học thành công!'))
            except Exception as e:
                return redirect(url_for('khaibao_tuan',
                                        semester_id=sem_filter,
                                        error_msg=f'Lỗi DB: {str(e)}'))

        # ── XÓA TUẦN ─────────────────────────────────────────────
        elif action == 'delete':
            ids_str = request.form.get('week_ids', '')
            ids = [int(x) for x in ids_str.split(',') if x.strip().isdigit()]
            if not ids:
                return redirect(url_for('khaibao_tuan', semester_id=sem_filter,
                                        error_msg='Không có tuần nào được chọn!'))
            try:
                supabase.table("school_weeks").delete().in_('week_id', ids).execute()
                return redirect(url_for('khaibao_tuan', semester_id=sem_filter,
                                        msg=f'Đã xóa {len(ids)} tuần thành công!'))
            except Exception as e:
                return redirect(url_for('khaibao_tuan', semester_id=sem_filter,
                                        error_msg=f'Không thể xóa: {str(e)}'))

    # ── LẤY DANH SÁCH TUẦN (GET) ──────────────────────────────────
    per_page = request.args.get('per_page', 100, type=int)
    page = request.args.get('page', 1, type=int)

    query = supabase.table("school_weeks").select("*, semesters(semester_name)")
    if selected_semester_id:
        query = query.eq('semester_id', selected_semester_id)
    all_weeks = query.order('week_number').execute().data

    # Flatten semester_name
    for w in all_weeks:
        sem_obj = w.pop('semesters', None)
        w['semester_name'] = sem_obj['semester_name'] if sem_obj else ''

    # Phân trang thủ công
    total = len(all_weeks)
    total_pages = max(1, (total + per_page - 1) // per_page)
    page = max(1, min(page, total_pages))
    start_idx = (page - 1) * per_page
    weeks = all_weeks[start_idx: start_idx + per_page]

    return render_template(
        'admin/khaibao_tuan.html',
        semesters=semesters,
        selected_semester_id=selected_semester_id,
        weeks=weeks,
        total_pages=total_pages,
        current_page=page,
        msg=msg,
        error_msg=error_msg,
    )


# ======================
# QUẢN LÝ THỜI KHÓA BIỂU (TRANG DANH SÁCH)
# ======================
@app.route('/admin/thoikhoabieu')
def quanly_tkb():
    if 'user_id' not in session or session.get('role_name') != 'Admin':
        return redirect(url_for('login'))

    msg = request.args.get('msg')
    error_msg = request.args.get('error_msg')

    # ── BỘ LỌC ───────────────────────────────────────────────────
    grades = supabase.table("grades").select("*").order("grade_number").execute().data
    semesters = supabase.table("semesters").select("*").order("semester_id").execute().data
    years_raw = supabase.table("academic_years").select("year_id, year_name").execute().data
    year_map = {y['year_id']: y['year_name'] for y in years_raw}
    for s in semesters:
        year_name = year_map.get(s.get('year_id'), '')
        s['display_name'] = f"{s['semester_name']} - {year_name}" if year_name else s['semester_name']

    selected_grade = request.args.get('grade_id', type=int)
    selected_class = request.args.get('class_id', type=int)
    selected_semester = request.args.get('semester_id', type=int)

    # Mặc định khối 1, học kỳ hiện tại
    if not selected_grade and grades:
        selected_grade = grades[0]['grade_id']
    if not selected_semester:
        cur = next((s for s in semesters if s.get('is_current')), None)
        selected_semester = cur['semester_id'] if cur else (semesters[0]['semester_id'] if semesters else None)

    # ── LẤY academic_year_id TỪ HỌC KỲ ĐANG CHỌN ───────────────
    academic_year_id = None
    if selected_semester:
        sem_info = next((s for s in semesters if s['semester_id'] == selected_semester), None)
        if sem_info:
            academic_year_id = sem_info.get('year_id')

    # ── LẤY LỚP THEO KHỐI + NĂM HỌC ─────────────────────────────
    classes_q = supabase.table("classes").select("*").order("class_name")
    if selected_grade:
        classes_q = classes_q.eq("grade_id", selected_grade)
    if academic_year_id:
        classes_q = classes_q.eq("academic_year_id", academic_year_id)
    if selected_class:
        classes_q = classes_q.eq("class_id", selected_class)
    classes = classes_q.execute().data

    # ── LẤY TẤT CẢ TUẦN TRONG HỌC KỲ ────────────────────────────
    weeks_raw = []
    if selected_semester:
        weeks_raw = supabase.table("school_weeks") \
            .select("*") \
            .eq("semester_id", selected_semester) \
            .order("week_number") \
            .execute().data

    # Tạo label "Tuần X" cho mỗi tuần
    all_weeks = []
    for w in weeks_raw:
        exam_name = w.get('exam_week_name')
        if exam_name == 'null' or not exam_name:
            exam_name = None
        label = exam_name or f"Tuần {w['week_number']}"
        all_weeks.append({**w, 'week_label': label})

    week_ids = [w['week_id'] for w in all_weeks]

    # ── LẤY TKB ĐÃ CÓ ────────────────────────────────────────────
    # Map: {(class_id, week_id)} = True
    existing_map = set()
    if week_ids:
        existing_raw = supabase.table("timetable") \
            .select("class_id, week_id") \
            .in_("week_id", week_ids) \
            .execute().data
        for e in existing_raw:
            existing_map.add((e['class_id'], e['week_id']))

    # ── PHÂN TRANG LỚP ────────────────────────────────────────────
    per_page = request.args.get('per_page', 100, type=int)
    page = request.args.get('page', 1, type=int)
    total = len(classes)
    total_pages = max(1, (total + per_page - 1) // per_page)
    page = max(1, min(page, total_pages))
    classes_page = classes[(page - 1) * per_page: page * per_page]

    # ── XÂY DỰNG ROWS CHO BẢNG ────────────────────────────────────
    class_weeks = []
    for cls in classes_page:
        has_tkb = []
        no_tkb = []
        for w in all_weeks:
            if not w.get('is_visible', True):
                continue
            entry = {'week_id': w['week_id'], 'week_label': w['week_label']}
            if (cls['class_id'], w['week_id']) in existing_map:
                has_tkb.append(entry)
            else:
                no_tkb.append(entry)
        class_weeks.append({
            'class_id': cls['class_id'],
            'class_name': cls['class_name'],
            'has_tkb': has_tkb,
            'no_tkb': no_tkb,
        })

    # ── DATA CHO MODAL (JS) ────────────────────────────────────────
    # subjects theo grade_id → dict
    all_subjects = supabase.table("subjects").select("*").execute().data
    subjects_by_grade = {}
    for s in all_subjects:
        gid = str(s['grade_id'])
        subjects_by_grade.setdefault(gid, []).append({
            'subject_id': s['subject_id'],
            'subject_name': s['subject_name'],
        })

    all_teachers = supabase.table("teachers") \
        .select("teacher_id, full_name") \
        .order("full_name") \
        .execute().data
    # Lấy phân công chuyên môn: {class_id: {subject_id: teacher_id}}
    assignments_raw = supabase.table("teaching_assignments") \
        .select("class_id, subject_id, teacher_id") \
        .eq("semester_id", selected_semester) \
        .execute().data

    # Build map: "class_id_subject_id" -> teacher_id
    assignment_map = {}
    for a in assignments_raw:
        key = f"{a['class_id']}_{a['subject_id']}"
        assignment_map[key] = a['teacher_id']

    current_year = supabase.table("academic_years").select("year_id").eq("is_current",
                                                                         True).execute().data  # ← thêm dấu .
    current_year_id = current_year[0]['year_id'] if current_year else None

    all_classes_raw = supabase.table("classes").select("*").order("class_name").execute().data
    classes_json = [
        {'class_id': c['class_id'], 'class_name': c['class_name'], 'grade_id': c['grade_id']}
        for c in all_classes_raw
        if c.get('academic_year_id') == current_year_id
    ]

    return render_template(
        'admin/thoikhoabieu.html',
        grades=grades,
        classes=classes,
        semesters=semesters,
        selected_grade=selected_grade,
        selected_class=selected_class,
        selected_semester=selected_semester,
        class_weeks=class_weeks,
        all_weeks=all_weeks,
        total_pages=total_pages,
        current_page=page,
        subjects_by_grade=subjects_by_grade,
        teachers=all_teachers,
        classes_json=classes_json,
        msg=msg,
        error_msg=error_msg,
        assignment_map=assignment_map,
    )


# ======================
# API: LẤY DỮ LIỆU TKB CỦA 1 LỚP + 1 TUẦN
# ======================
@app.route('/admin/tkb/data')
def get_tkb_data():
    if 'user_id' not in session:
        return {'error': 'Unauthorized'}, 401

    class_id = request.args.get('class_id', type=int)
    week_id = request.args.get('week_id', type=int)

    if not class_id or not week_id:
        return {'entries': []}

    rows = supabase.table("timetable") \
        .select("*") \
        .eq("class_id", class_id) \
        .eq("week_id", week_id) \
        .execute().data

    return {'entries': rows}


# ======================
# API: LƯU TKB (XÓA CŨ + INSERT MỚI)
# ======================
@app.route('/admin/tkb/save', methods=['POST'])
def save_tkb():
    if 'user_id' not in session or session.get('role_name') != 'Admin':
        return {'success': False, 'error': 'Unauthorized'}, 401

    data = request.get_json()
    class_id = data.get('class_id')
    week_id = data.get('week_id')
    semester_id = data.get('semester_id')
    entries = data.get('entries', [])

    if not class_id or not week_id or not semester_id:
        return {'success': False, 'error': 'Thiếu thông tin lớp / tuần / học kỳ'}

    try:
        # Xóa TKB cũ của lớp + tuần này
        supabase.table("timetable") \
            .delete() \
            .eq("class_id", class_id) \
            .eq("week_id", week_id) \
            .execute()

        # Insert mới
        if entries:
            rows = []
            for e in entries:
                rows.append({
                    'class_id': class_id,
                    'week_id': week_id,
                    'semester_id': semester_id,
                    'day_of_week': e['day_of_week'],
                    'session': e['session'],
                    'period': e['period'],
                    'subject_id': e['subject_id'],
                    'teacher_id': e.get('teacher_id'),
                })
            supabase.table("timetable").insert(rows).execute()

        return {'success': True}
    except Exception as ex:
        return {'success': False, 'error': str(ex)}


# ======================
# API: XÓA TKB 1 LỚP + 1 TUẦN
# ======================
@app.route('/admin/tkb/delete', methods=['POST'])
def delete_tkb():
    if 'user_id' not in session or session.get('role_name') != 'Admin':
        return {'success': False, 'error': 'Unauthorized'}, 401

    data = request.get_json()
    class_id = data.get('class_id')
    week_id = data.get('week_id')

    if not class_id or not week_id:
        return {'success': False, 'error': 'Thiếu thông tin'}

    try:
        supabase.table("timetable") \
            .delete() \
            .eq("class_id", class_id) \
            .eq("week_id", week_id) \
            .execute()
        return {'success': True}
    except Exception as ex:
        return {'success': False, 'error': str(ex)}


# ======================
# XUẤT EXCEL TKB
# ======================
@app.route('/admin/tkb/export')
def export_tkb_excel():
    if 'user_id' not in session or session.get('role_name') != 'Admin':
        return redirect(url_for('login'))

    semester_id = request.args.get('semester_id', type=int)
    class_id = request.args.get('class_id', type=int)

    # Lấy dữ liệu TKB
    query = supabase.table("timetable") \
        .select("*, classes(class_name), school_weeks(week_number), subjects(subject_name), teachers(full_name)")
    if semester_id:
        query = query.eq("semester_id", semester_id)
    if class_id:
        query = query.eq("class_id", class_id)

    rows = query.order("class_id").order("week_id").order("day_of_week").order("period").execute().data

    # Tạo DataFrame
    records = []
    for r in rows:
        records.append({
            'Lớp': (r.get('classes') or {}).get('class_name', ''),
            'Tuần': (r.get('school_weeks') or {}).get('week_number', ''),
            'Thứ': r.get('day_of_week', ''),
            'Buổi': r.get('session', ''),
            'Tiết': r.get('period', ''),
            'Môn học': (r.get('subjects') or {}).get('subject_name', ''),
            'Giáo viên': (r.get('teachers') or {}).get('full_name', ''),
        })

    df = pd.DataFrame(records)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='ThoiKhoaBieu')
    output.seek(0)

    today = datetime.now().strftime("%d-%m-%Y")
    return send_file(output,
                     download_name=f"TKB_{today}.xlsx",
                     as_attachment=True,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ======================
# API: SAO CHÉP TKB TỪ TUẦN KHÁC (CÙNG LỚP)
# ======================
@app.route('/admin/tkb/copy', methods=['POST'])
def copy_tkb():
    if 'user_id' not in session or session.get('role_name') != 'Admin':
        return {'success': False, 'error': 'Unauthorized'}, 401

    data = request.get_json()
    class_id = data.get('class_id')
    src_week_id = data.get('src_week_id')
    dest_week_id = data.get('dest_week_id')
    semester_id = data.get('semester_id')

    if not all([class_id, src_week_id, dest_week_id, semester_id]):
        return {'success': False, 'error': 'Thiếu thông tin'}

    if src_week_id == dest_week_id:
        return {'success': False, 'error': 'Tuần nguồn và đích không được trùng nhau!'}

    try:
        # Lấy TKB của tuần nguồn
        src_rows = supabase.table("timetable") \
            .select("*") \
            .eq("class_id", class_id) \
            .eq("week_id", src_week_id) \
            .execute().data

        if not src_rows:
            return {'success': False, 'error': 'Tuần nguồn chưa có dữ liệu TKB!'}

        # Xóa TKB cũ của tuần đích
        supabase.table("timetable") \
            .delete() \
            .eq("class_id", class_id) \
            .eq("week_id", dest_week_id) \
            .execute()

        # Insert mới với week_id = dest_week_id
        new_rows = []
        for r in src_rows:
            new_rows.append({
                'class_id': class_id,
                'week_id': dest_week_id,
                'semester_id': semester_id,
                'day_of_week': r['day_of_week'],
                'session': r['session'],
                'period': r['period'],
                'subject_id': r['subject_id'],
                'teacher_id': r.get('teacher_id'),
            })
        supabase.table("timetable").insert(new_rows).execute()

        return {'success': True, 'count': len(new_rows)}
    except Exception as e:
        return {'success': False, 'error': str(e)}


# ======================
# API: LẤY DỮ LIỆU TKB TỪ HK1 (COPY TOÀN BỘ LỚP)
# ======================
@app.route('/admin/tkb/copy-from-hk1', methods=['POST'])
def copy_tkb_from_hk1():
    if 'user_id' not in session or session.get('role_name') != 'Admin':
        return {'success': False, 'error': 'Unauthorized'}, 401

    data = request.get_json()
    class_id = data.get('class_id')
    dest_semester_id = data.get('dest_semester_id')

    if not class_id or not dest_semester_id:
        return {'success': False, 'error': 'Thiếu thông tin'}

    try:
        # Tìm HK1 (semester có tên chứa "kỳ 1" hoặc semester_id nhỏ hơn)
        all_sems = supabase.table("semesters").select("*").order("semester_id").execute().data
        src_sem = next((s for s in all_sems if 'kỳ 1' in s['semester_name'].lower()), None)

        if not src_sem:
            return {'success': False, 'error': 'Không tìm thấy Học kỳ 1 trong hệ thống!'}

        if src_sem['semester_id'] == dest_semester_id:
            return {'success': False, 'error': 'Không thể copy từ HK1 sang chính HK1!'}

        # Lấy toàn bộ TKB của lớp ở HK1
        src_rows = supabase.table("timetable") \
            .select("*") \
            .eq("class_id", class_id) \
            .eq("semester_id", src_sem['semester_id']) \
            .execute().data

        if not src_rows:
            return {'success': False, 'error': f'Lớp này chưa có TKB ở {src_sem["semester_name"]}!'}

        # Lấy mapping tuần HK1 → tuần HK2 theo week_number
        hk1_weeks = supabase.table("school_weeks") \
            .select("*") \
            .eq("semester_id", src_sem['semester_id']) \
            .execute().data

        hk2_weeks = supabase.table("school_weeks") \
            .select("*") \
            .eq("semester_id", dest_semester_id) \
            .execute().data

        # Map theo week_number: hk1_week_id → hk2_week_id
        hk1_num_to_id = {w['week_number']: w['week_id'] for w in hk1_weeks}
        hk2_num_to_id = {w['week_number']: w['week_id'] for w in hk2_weeks}
        hk1_id_to_num = {w['week_id']: w['week_number'] for w in hk1_weeks}

        # Xóa toàn bộ TKB hiện tại của lớp ở HK đích
        supabase.table("timetable") \
            .delete() \
            .eq("class_id", class_id) \
            .eq("semester_id", dest_semester_id) \
            .execute()

        # Copy sang HK đích
        new_rows = []
        for r in src_rows:
            week_num = hk1_id_to_num.get(r['week_id'])
            dest_week_id = hk2_num_to_id.get(week_num) if week_num else None

            if not dest_week_id:
                continue  # Bỏ qua nếu HK2 không có tuần tương ứng

            new_rows.append({
                'class_id': class_id,
                'week_id': dest_week_id,
                'semester_id': dest_semester_id,
                'day_of_week': r['day_of_week'],
                'session': r['session'],
                'period': r['period'],
                'subject_id': r['subject_id'],
                'teacher_id': r.get('teacher_id'),
            })

        if new_rows:
            supabase.table("timetable").insert(new_rows).execute()

        return {'success': True, 'count': len(new_rows)}
    except Exception as e:
        return {'success': False, 'error': str(e)}


# ======================
# API: NHẬP TKB TỪ EXCEL
# ======================
@app.route('/admin/tkb/import', methods=['POST'])
def import_tkb_excel():
    if 'user_id' not in session or session.get('role_name') != 'Admin':
        return {'success': False, 'error': 'Unauthorized'}, 401

    file = request.files.get('file')
    class_id = request.form.get('class_id', type=int)
    week_id = request.form.get('week_id', type=int)
    semester_id = request.form.get('semester_id', type=int)

    if not file or not class_id or not week_id or not semester_id:
        return {'success': False, 'error': 'Thiếu thông tin'}

    try:
        df = pd.read_excel(file, header=0)
        df.columns = df.columns.astype(str).str.strip()

        # Lấy map tên môn → subject_id và tên GV → teacher_id
        subjects = supabase.table("subjects").select("subject_id, subject_name").execute().data
        teachers = supabase.table("teachers").select("teacher_id, full_name").execute().data
        sub_map = {s['subject_name'].strip().lower(): s['subject_id'] for s in subjects}
        tea_map = {t['full_name'].strip().lower(): t['teacher_id'] for t in teachers}

        # Xóa TKB cũ
        supabase.table("timetable").delete() \
            .eq("class_id", class_id).eq("week_id", week_id).execute()

        rows = []
        errors = []
        for idx, row in df.iterrows():
            day = row.get('day_of_week') or row.get('Thứ')
            session = str(row.get('session') or row.get('Buổi') or '').strip()
            period = row.get('period') or row.get('Tiết')
            sub_name = str(row.get('subject_name') or row.get('Môn học') or '').strip()
            tea_name = str(row.get('teacher_name') or row.get('Giáo viên') or '').strip()

            if not day or not session or not period:
                errors.append(f"Dòng {idx + 2}: thiếu Thứ/Buổi/Tiết")
                continue

            sub_id = sub_map.get(sub_name.lower())
            tea_id = tea_map.get(tea_name.lower()) if tea_name else None

            if not sub_id:
                errors.append(f"Dòng {idx + 2}: môn '{sub_name}' không tìm thấy")
                continue

            rows.append({
                'class_id': class_id,
                'week_id': week_id,
                'semester_id': semester_id,
                'day_of_week': int(day),
                'session': session,
                'period': int(period),
                'subject_id': sub_id,
                'teacher_id': tea_id,
            })

        if rows:
            supabase.table("timetable").insert(rows).execute()

        return {'success': True, 'count': len(rows), 'errors': errors}
    except Exception as e:
        return {'success': False, 'error': str(e)}


# ======================
# API: TẢI FILE EXCEL MẪU NHẬP TKB
# ======================
@app.route('/admin/tkb/import-template')
def tkb_import_template():
    data = {
        'day_of_week': [2, 2, 3],
        'session': ['Sáng', 'Sáng', 'Chiều'],
        'period': [1, 2, 1],
        'subject_name': ['Toán', 'Tiếng Việt', 'Toán'],
        'teacher_name': ['Nguyễn Văn A', 'Trần Thị B', 'Nguyễn Văn A'],
    }
    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='TKB_Mau')
    output.seek(0)
    return send_file(output, download_name="Mau_Nhap_TKB.xlsx", as_attachment=True)


# ======================
# API: LẤY GIÁO VIÊN THEO MÔN HỌC + LỚP + HỌC KỲ
# ======================
@app.route('/admin/tkb/teachers-by-subject')
def get_teachers_by_subject():
    if 'user_id' not in session:
        return {'teachers': []}, 401

    subject_id = request.args.get('subject_id', type=int)
    class_id = request.args.get('class_id', type=int)
    semester_id = request.args.get('semester_id', type=int)

    if not subject_id or not class_id or not semester_id:
        return {'teachers': []}

    # Lấy GV được phân công dạy môn này tại lớp này trong học kỳ này
    res = supabase.table("teaching_assignments") \
        .select("teacher_id, teachers(teacher_id, full_name)") \
        .eq("subject_id", subject_id) \
        .eq("class_id", class_id) \
        .eq("semester_id", semester_id) \
        .execute().data

    teachers = []
    for r in res:
        t = r.get('teachers')
        if t:
            teachers.append({
                'teacher_id': t['teacher_id'],
                'full_name': t['full_name']
            })

    return {'teachers': teachers}


# ==========================================
# PHÂN HỆ GIÁO VIÊN: NHẬP ĐIỂM MÔN HỌC
# ==========================================
@app.route('/teacher/input_scores')
def input_scores():
    if 'user_id' not in session: return redirect(url_for('login'))
    t_res = supabase.table("teachers").select("teacher_id").eq("user_id", session['user_id']).execute()
    teacher_id = t_res.data[0]['teacher_id']

    # 1. PHÂN CẤP: Lấy danh sách Học kỳ giáo viên có dạy
    assigned_sems = supabase.table("teaching_assignments").select("semester_id, semesters(semester_name)").eq(
        "teacher_id", teacher_id).execute().data
    sem_map = {a['semester_id']: a['semesters']['semester_name'] for a in assigned_sems}
    semesters = [{'semester_id': k, 'semester_name': v} for k, v in sem_map.items()]

    # Mặc định chọn học kỳ hiện tại hoặc học kỳ đầu tiên có trong danh sách phân công
    selected_semester_id = request.args.get('semester_id') or (str(semesters[0]['semester_id']) if semesters else '1')

    # 2. PHÂN CẤP: Lấy Khối/Lớp/Môn theo Học kỳ đã chọn
    assignments = supabase.table("teaching_assignments") \
        .select(
        "class_id, classes(class_name, grade_id, grades(grade_name)), subject_id, subjects(subject_name, evaluation_type)") \
        .eq("teacher_id", teacher_id).eq("semester_id", selected_semester_id).execute().data

    # ĐƯA BỘ LỌC LÊN TRƯỚC ĐỂ LỌC MÔN HỌC
    selected_grade = request.args.get('grade')
    selected_class = request.args.get('class_id')
    selected_subject = request.args.get('subject_id')

    if selected_class and not any(str(a['class_id']) == str(selected_class) for a in assignments):
        selected_class = None
        selected_grade = None
        selected_subject = None

    filter_grades, filter_classes, filter_subjects = {}, {}, {}
    for a in assignments:
        g_id = a['classes']['grade_id']
        filter_grades[g_id] = a['classes']['grades']['grade_name']
        filter_classes[a['class_id']] = {'name': a['classes']['class_name'], 'grade_id': g_id}

        # Thêm logic check để dropdown môn không bị lộn xộn lớp khác
        if selected_class:
            if str(a['class_id']) == str(selected_class):
                filter_subjects[a['subject_id']] = a['subjects']['subject_name']
        elif selected_grade:
            if str(g_id) == str(selected_grade):
                filter_subjects[a['subject_id']] = a['subjects']['subject_name']
        else:
            filter_subjects[a['subject_id']] = a['subjects']['subject_name']

    students_list, score_map, subject_info = [], {}, None

    if selected_class and selected_subject:
        # Lấy thông tin loại môn học (Chấm điểm hay Đánh giá)
        subject_info = supabase.table("subjects").select("*").eq("subject_id", selected_subject).single().execute().data
        is_eval = subject_info['evaluation_type'] == 'Đánh giá'

        # Lấy danh sách học sinh của lớp
        st_res = supabase.table("student_enrollments").select("student_id, students(ho_ten, ma_dinh_danh)").eq(
            "class_id", selected_class).execute()
        students_list = sorted(st_res.data, key=lambda x: get_vietnamese_sort_key(x['students']['ho_ten']))

        # 3. GỘP ĐIỂM CẢ NĂM: Truy vấn toàn bộ điểm của môn này tại lớp này
        all_scores = supabase.table("subject_scores_v2").select("*").eq("subject_id", selected_subject).execute().data

        # Khởi tạo khung dữ liệu trống cho từng học sinh để tránh lỗi Undefined trong Jinja2
        for s in students_list:
            sid = s['student_id']
            score_map[sid] = {'hk1': {}, 'hk2': {}, 'current': {}, 'dtb_cn': ''}

        # Hàm dọn dẹp dữ liệu và chuyển đổi điểm số sang Đ/C cho môn đánh giá
        def clean(v):
            if v is None or str(v) == 'nan': return ''
            if is_eval and isinstance(v, (int, float)): return 'Đ' if v >= 5 else 'C'
            return v

        # Đổ dữ liệu vào map theo từng học kỳ
        for s in all_scores:
            sid = s['student_id']
            if sid in score_map:
                data_cleaned = {k: clean(v) for k, v in s.items()}

                # Phân loại điểm theo kỳ để hiển thị cột tương ứng
                if str(s['semester_id']) == '1':
                    score_map[sid]['hk1'] = data_cleaned
                elif str(s['semester_id']) == '2':
                    score_map[sid]['hk2'] = data_cleaned

                # Lưu bản ghi của học kỳ đang chọn vào 'current' để giáo viên chỉnh sửa
                if str(s['semester_id']) == str(selected_semester_id):
                    score_map[sid]['current'] = data_cleaned

                # Cập nhật ĐTB Cả năm nếu có
                if s.get('dtb_cn'):
                    score_map[sid]['dtb_cn'] = clean(s['dtb_cn'])

    return render_template('teacher/input_scores.html',
                           semesters=semesters,
                           filter_grades=filter_grades,
                           filter_classes=filter_classes,
                           filter_subjects=filter_subjects,
                           students=students_list,
                           score_map=score_map,
                           subject_info=subject_info,
                           selected_grade=selected_grade,
                           selected_class=selected_class,
                           selected_subject=selected_subject,
                           selected_semester_id=selected_semester_id)

# ====================================
# LƯU ĐIỂM
# ====================================
@app.route('/teacher/save_subject_scores', methods=['POST'])
def save_subject_scores():
    if 'user_id' not in session: return redirect(url_for('login'))
    form_data = request.form.to_dict()
    sub_id = form_data.get('subject_id_hidden')
    sem_id = form_data.get('semester_id_hidden')

    # Kiểm tra loại môn học từ bảng subjects
    subj = supabase.table("subjects").select("evaluation_type").eq("subject_id", sub_id).single().execute().data
    is_grading = subj['evaluation_type'] == 'Chấm điểm'

    t_res = supabase.table("teachers").select("teacher_id").eq("user_id", session['user_id']).execute()
    teacher_id = t_res.data[0]['teacher_id']

    student_ids = set([k.split('_')[1] for k in form_data.keys() if k.startswith('tx1_')])

    # Lấy history để tính ĐTB cả năm cho môn tính điểm
    history = {}
    if is_grading:
        existing = supabase.table("subject_scores_v2").select("student_id, dtb_hk1, dtb_hk2").eq("subject_id",
                                                                                                 sub_id).execute().data
        for s in existing:
            sid = s['student_id']
            if sid not in history: history[sid] = {'dtb_hk1': None, 'dtb_hk2': None}
            if s.get('dtb_hk1'): history[sid]['dtb_hk1'] = s['dtb_hk1']
            if s.get('dtb_hk2'): history[sid]['dtb_hk2'] = s['dtb_hk2']

    upsert_data = []
    for sid in student_ids:
        row = {"student_id": int(sid), "subject_id": int(sub_id), "semester_id": int(sem_id),
               "updated_by_teacher_id": teacher_id,
               "feedback_highlight": form_data.get(f'feedback_{sid}'),
               "feedback_report_card": form_data.get(f'report_{sid}')}

        # A. NẾU LÀ MÔN CHẤM ĐIỂM (Toán, TV...)
        if is_grading:
            def get_f(k):
                v = form_data.get(f'{k}_{sid}')
                if not v or not v.strip(): return None
                num = float(v)
                return num / 10 if num > 10 else num

            tx1, tx2, tx3, tx4 = get_f('tx1'), get_f('tx2'), get_f('tx3'), get_f('tx4')
            gk, ck = get_f('gk'), get_f('ck')
            row.update({"tx_1": tx1, "tx_2": tx2, "tx_3": tx3, "tx_4": tx4, "midterm_score": gk, "final_score": ck})

            if gk is not None and ck is not None:
                valid_txs = [v for v in [tx1, tx2, tx3, tx4] if v is not None]
                dtb_current = round((sum(valid_txs) + gk * 2 + ck * 3) / (len(valid_txs) + 5), 1)

                h = history.get(int(sid), {})
                h1 = dtb_current if str(sem_id) == '1' else h.get('dtb_hk1')
                h2 = dtb_current if str(sem_id) == '2' else h.get('dtb_hk2')
                if h1 is not None and h2 is not None: row["dtb_cn"] = round((h1 + h2 * 2) / 3, 1)
                row["dtb_hk2" if str(sem_id) == '2' else "dtb_hk1"] = dtb_current

        # B. NẾU LÀ MÔN ĐÁNH GIÁ (GDTC, Đạo đức...)
        else:
            # 1. Chuyển Đ/C thành số để lưu vào các cột Real (Đ=10.0, C=0.0)
            def map_to_num(k):
                val = form_data.get(f'{k}_{sid}')
                return 10.0 if val == 'Đ' else (0.0 if val == 'C' else None)

            row.update({
                "tx_1": map_to_num('tx1'), "tx_2": map_to_num('tx2'),
                "tx_3": map_to_num('tx3'), "tx_4": map_to_num('tx4'),
                "dtb_hk1": map_to_num('avg1'), "dtb_hk2": map_to_num('avg2'), "dtb_cn": map_to_num('avgcn')
            })

            # 2. Lưu trực tiếp Đ/C vào các cột TEXT bạn đã tạo
            row["midterm_result_level"] = form_data.get(f'gk_{sid}')
            row["final_result_level"] = form_data.get(f'ck_{sid}')

        upsert_data.append(row)

    if upsert_data:
        supabase.table("subject_scores_v2").upsert(upsert_data,
                                                   on_conflict="student_id, subject_id, semester_id").execute()

    return redirect(
        url_for('input_scores', grade=form_data.get('grade_hidden'), class_id=form_data.get('class_id_hidden'),
                subject_id=sub_id, semester_id=sem_id, msg="Lưu bảng điểm thành công!"))


# ==============================
# XUẤT FILE ĐIỂM
# ==============================
def no_accent_vietnamese(s):
    if not s: return ""
    s = re.sub(r'[àáạảãâầấậẩẫăằắặẳẵ]', 'a', s)
    s = re.sub(r'[ÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴ]', 'A', s)
    s = re.sub(r'[èéẹẻẽêềếệểễ]', 'e', s)
    s = re.sub(r'[ÈÉẸẺẼÊỀẾỆỂỄ]', 'E', s)
    s = re.sub(r'[òóọỏõôồốộổỗơờớợởỡ]', 'o', s)
    s = re.sub(r'[ÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠ]', 'O', s)
    s = re.sub(r'[ìíịỉĩ]', 'i', s)
    s = re.sub(r'[ÌÍỊỈĨ]', 'I', s)
    s = re.sub(r'[ùúụủũưừứựửữ]', 'u', s)
    s = re.sub(r'[ÙÚỤỦŨƯỪỨỰỬỮ]', 'U', s)
    s = re.sub(r'[ỳýỵỷỹ]', 'y', s)
    s = re.sub(r'[ỲÝỴỶỸ]', 'Y', s)
    s = re.sub(r'[đ]', 'd', s)
    s = re.sub(r'[Đ]', 'D', s)
    return s.replace(" ", "_")


@app.route('/teacher/export_scores')
def export_scores():
    if 'user_id' not in session: return redirect(url_for('login'))

    class_id = request.args.get('class_id')
    subject_id = request.args.get('subject_id')
    semester_id = request.args.get('semester_id')

    if not all([class_id, subject_id, semester_id]):
        return "Vui lòng chọn đầy đủ bộ lọc trước khi xuất!"

    # 1. TRUY VẤN THÔNG TIN BỔ TRỢ
    cls_res = supabase.table("classes").select("class_name").eq("class_id", class_id).single().execute().data
    sub_res = supabase.table("subjects").select("subject_name, evaluation_type").eq("subject_id",
                                                                                    subject_id).single().execute().data
    sem_res = supabase.table("semesters").select("semester_name").eq("semester_id", semester_id).single().execute().data
    is_eval = sub_res['evaluation_type'] == 'Đánh giá'

    # 2. LẤY DỮ LIỆU HỌC SINH VÀ ĐIỂM
    st_res = supabase.table("student_enrollments").select("student_id, students(ho_ten, ma_dinh_danh)").eq("class_id",
                                                                                                           class_id).execute()
    students = sorted(st_res.data, key=lambda x: get_vietnamese_sort_key(x['students']['ho_ten']))

    all_scores = supabase.table("subject_scores_v2").select("*").eq("subject_id", subject_id).execute().data
    score_map = {}
    for s in all_scores:
        sid = s['student_id']
        if sid not in score_map: score_map[sid] = {'hk1': {}, 'hk2': {}, 'cn': ''}
        if str(s['semester_id']) == '1':
            score_map[sid]['hk1'] = s
        elif str(s['semester_id']) == '2':
            score_map[sid]['hk2'] = s
        if s.get('dtb_cn'): score_map[sid]['cn'] = s['dtb_cn']

    # 3. KHỞI TẠO EXCEL
    wb = Workbook()
    ws = wb.active
    ws.title = "BangDiem"

    bold_font = Font(bold=True)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'),
                         bottom=Side(style='thin'))

    def fmt(v):
        if v is None or str(v) == 'nan' or v == '': return ""
        if is_eval and isinstance(v, (int, float)): return "Đ" if v >= 5 else "C"
        return v

    # --- DÒNG 1: TIÊU ĐỀ CHÍNH ---
    ws.merge_cells('A1:N1')
    ws['A1'] = f"BẢNG ĐIỂM MÔN: {sub_res['subject_name'].upper()}"
    ws['A1'].font = Font(size=14, bold=True);
    ws['A1'].alignment = center_align

    # --- DÒNG 2: THÔNG TIN LỚP/KỲ ---
    ws.append(
        [f"Lớp: {cls_res['class_name']}", "", "", f"Học kỳ: {sem_res['semester_name']}", "", "", "", "", "", "", "", "",
         "", ""])
    ws.merge_cells('A2:C2');
    ws.merge_cells('D2:N2')
    ws['A2'].font = bold_font;
    ws['E2'].font = bold_font

    # --- DÒNG 3 & 4: HEADER 2 TẦNG (KHỚP ẢNH image_e258d5.png) ---
    # Merge các ô dọc
    ws.merge_cells('A3:A4');
    ws['A3'] = 'STT'
    ws.merge_cells('B3:B4');
    ws['B3'] = 'Mã định danh Bộ GD&ĐT'
    ws.merge_cells('C3:C4');
    ws['C3'] = 'Họ tên'

    # Header ĐĐGtx gộp 4 ô ngang
    ws.merge_cells('D3:G3');
    ws['D3'] = 'ĐĐGtx'
    ws['D4'] = '1';
    ws['E4'] = '2';
    ws['F4'] = '3';
    ws['G4'] = '4'

    # Các ô đơn gộp dọc
    ws.merge_cells('H3:H4');
    ws['H3'] = 'ĐĐGgk'
    ws.merge_cells('I3:I4');
    ws['I3'] = 'ĐĐGck'
    ws.merge_cells('J3:J4');
    ws['J3'] = 'ĐTB HK1'
    ws.merge_cells('K3:K4');
    ws['K3'] = 'ĐTB HK2'
    ws.merge_cells('L3:L4');
    ws['L3'] = 'ĐTB CN'
    ws.merge_cells('M3:M4');
    ws['M3'] = 'Nhận xét sự tiến bộ, ưu điểm nổi bật, hạn chế chủ yếu'
    ws.merge_cells('N3:N4');
    ws['N3'] = 'Nhận xét ghi học bạ'

    # 4. ĐỔ DỮ LIỆU
    for idx, s in enumerate(students, 1):
        sid = s['student_id']
        d = score_map.get(sid, {'hk1': {}, 'hk2': {}, 'cn': ''})
        curr = d['hk2'] if str(semester_id) == '2' else d['hk1']

        # Lấy GK/CK tùy theo môn
        gk_val = curr.get('midterm_result_level') if is_eval else curr.get('midterm_score')
        ck_val = curr.get('final_result_level') if is_eval else curr.get('final_score')

        row_data = [
            idx, s['students']['ma_dinh_danh'], s['students']['ho_ten'],
            fmt(curr.get('tx_1')), fmt(curr.get('tx_2')), fmt(curr.get('tx_3')), fmt(curr.get('tx_4')),
            fmt(gk_val), fmt(ck_val),
            fmt(d['hk1'].get('dtb_hk1')), fmt(d['hk2'].get('dtb_hk2')), fmt(d.get('cn')),
            curr.get('feedback_highlight', ''), curr.get('feedback_report_card', '')
        ]
        ws.append(row_data)

    # 5. ĐỊNH DẠNG (BOLD 4 HÀNG ĐẦU, BORDER, CĂN GIỮA)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=14):
        for cell in row:
            cell.border = thin_border
            cell.alignment = center_align
            if cell.row <= 4: cell.font = bold_font

    # Chỉnh độ rộng cột cho chuyên nghiệp
    ws.column_dimensions['B'].width = 20  # Mã định danh
    ws.column_dimensions['C'].width = 25  # Họ tên
    ws.column_dimensions['M'].width = 45  # Nhận xét 1
    ws.column_dimensions['N'].width = 45  # Nhận xét 2
    for col in ['A', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']:
        ws.column_dimensions[col].width = 10

    # 6. TÊN FILE KHÔNG DẤU
    safe_class = no_accent_vietnamese(cls_res['class_name'])
    safe_subject = no_accent_vietnamese(sub_res['subject_name'])
    filename = f"BangDiem_{safe_class}_{safe_subject}.xlsx"

    output = io.BytesIO();
    wb.save(output);
    output.seek(0)
    return send_file(output, download_name=filename, as_attachment=True)

# ===============================
# NẠP DỮ LỆU TỪ FILE EXCEL
# ===============================
@app.route('/teacher/import_scores', methods=['POST'])
def import_scores():
    if 'user_id' not in session: return redirect(url_for('login'))
    if 'file' not in request.files: return redirect(request.referrer)
    file = request.files['file']
    if file.filename == '': return redirect(request.referrer)

    # 1. LẤY LẠI CÁC BIẾN TỪ FORM ẨN TRƯỚC KHI REDIRECT
    sem_id = request.form.get('semester_id')
    sub_id = request.form.get('subject_id')
    class_id = request.form.get('class_id')
    grade = request.form.get('grade')

    t_res = supabase.table("teachers").select("teacher_id").eq("user_id", session['user_id']).execute()
    teacher_id = t_res.data[0]['teacher_id']

    try:
        # --- KHÓA BẢO MẬT: CHECK LỚP, MÔN & HỌC KỲ ---
        check_df = pd.read_excel(file, nrows=2, header=None)
        f_sub_name = str(check_df.iloc[0, 0]).replace("BẢNG ĐIỂM MÔN:", "").strip().upper()
        f_class_name = str(check_df.iloc[1, 0]).replace("Lớp:", "").strip().upper()
        f_sem_name = str(check_df.iloc[1, 3]).replace("Học kỳ:", "").strip().upper()

        db_cls = supabase.table("classes").select("class_name").eq("class_id", class_id).single().execute().data[
            'class_name'].upper()
        db_sub = supabase.table("subjects").select("subject_name").eq("subject_id", sub_id).single().execute().data[
            'subject_name'].upper()
        db_sem = supabase.table("semesters").select("semester_name").eq("semester_id", sem_id).single().execute().data[
            'semester_name'].upper()

        if f_class_name != db_cls or f_sub_name != db_sub or f_sem_name != db_sem:
            msg = f"Từ chối: File Excel ({f_class_name}-{f_sub_name}-{f_sem_name}) không khớp bộ lọc!"
            return redirect(
                url_for('input_scores', grade=grade, class_id=class_id, subject_id=sub_id, semester_id=sem_id, msg=msg))

        # -----------------------------------------------------------
        # Kiểm tra loại môn học để xử lý tính toán
        subj = supabase.table("subjects").select("evaluation_type").eq("subject_id", sub_id).single().execute().data
        is_grading = subj['evaluation_type'] == 'Chấm điểm'

        # 2. LẤY LỊCH SỬ ĐIỂM CỦA CẢ 2 KỲ ĐỂ TÍNH TOÁN CHÉO
        history = {}
        if is_grading:
            # Truy vấn cả HK1 và HK2 hiện có trong Database
            existing = supabase.table("subject_scores_v2").select("student_id, dtb_hk1, dtb_hk2").eq("subject_id",
                                                                                                     sub_id).execute().data
            history = {s['student_id']: {'h1': s.get('dtb_hk1'), 'h2': s.get('dtb_hk2')} for s in existing}

        # ĐỌC DỮ LIỆU TỪ DÒNG 5 (Bỏ qua 4 dòng đầu)
        df = pd.read_excel(file, skiprows=4, header=None)
        upsert_data = []

        for _, row in df.iterrows():
            ma_dinh_danh = str(row[1]).strip()
            if not ma_dinh_danh or ma_dinh_danh == 'nan': continue

            st_res = supabase.table("students").select("student_id").eq("ma_dinh_danh", ma_dinh_danh).execute()
            if not st_res.data: continue
            sid = st_res.data[0]['student_id']

            def parse_v(v):
                return float(v) if pd.notna(v) and str(v).strip() != '' else None

            row_record = {
                "student_id": sid, "subject_id": int(sub_id), "semester_id": int(sem_id),
                "updated_by_teacher_id": teacher_id,
                "feedback_highlight": str(row[12]) if pd.notna(row[12]) else None,
                "feedback_report_card": str(row[13]) if pd.notna(row[13]) else None
            }

            if is_grading:
                # A. MÔN CHẤM ĐIỂM: Tự tính ĐTB và Cả năm
                txs = [parse_v(row[3]), parse_v(row[4]), parse_v(row[5]), parse_v(row[6])]
                gk, ck = parse_v(row[7]), parse_v(row[8])
                valid_txs = [v for v in txs if v is not None]

                # Tính ĐTB của học kỳ hiện tại đang nạp
                dtb_current = round((sum(valid_txs) + gk * 2 + ck * 3) / (len(valid_txs) + 5),
                                    1) if gk is not None and ck is not None else None

                # Lấy dữ liệu kỳ cũ từ DB
                h_hist = history.get(sid, {'h1': None, 'h2': None})

                # Xác định h1, h2: Nếu nạp kỳ 1 thì lấy điểm mới làm h1 và lấy h2 từ DB. Ngược lại.
                h1 = dtb_current if str(sem_id) == '1' else h_hist['h1']
                h2 = dtb_current if str(sem_id) == '2' else h_hist['h2']

                row_record.update({
                    "tx_1": txs[0], "tx_2": txs[1], "tx_3": txs[2], "tx_4": txs[3],
                    "midterm_score": gk, "final_score": ck,
                    "dtb_hk1": h1, "dtb_hk2": h2
                })

                # TỰ ĐỘNG TÍNH ĐTB CẢ NĂM: (HK1 + HK2*2) / 3
                if h1 is not None and h2 is not None:
                    row_record["dtb_cn"] = round((h1 + h2 * 2) / 3, 1)
            else:
                # B. MÔN ĐÁNH GIÁ: Nạp hết Đ/C
                def map_eval(v):
                    return 10.0 if str(v).strip() == 'Đ' else (0.0 if str(v).strip() == 'C' else None)

                row_record.update({
                    "tx_1": map_eval(row[3]), "tx_2": map_eval(row[4]), "tx_3": map_eval(row[5]),
                    "tx_4": map_eval(row[6]),
                    "midterm_result_level": str(row[7]) if row[7] in ['Đ', 'C'] else None,
                    "final_result_level": str(row[8]) if row[8] in ['Đ', 'C'] else None,
                    "dtb_hk1": map_eval(row[9]), "dtb_hk2": map_eval(row[10]), "dtb_cn": map_eval(row[11])
                })
            upsert_data.append(row_record)

        if upsert_data:
            supabase.table("subject_scores_v2").upsert(upsert_data,
                                                       on_conflict="student_id, subject_id, semester_id").execute()
            msg = f"Thành công: Đã nạp điểm cho lớp {db_cls}!"
        else:
            msg = "Thất bại: Không có dữ liệu hợp lệ."

    except Exception as e:
        msg = f"Lỗi hệ thống: {str(e)}"

    return redirect(
        url_for('input_scores', grade=grade, class_id=class_id, subject_id=sub_id, semester_id=sem_id, msg=msg))


# ==========================================
# PHÂN HỆ GIÁO VIÊN: ĐÁNH GIÁ NĂNG LỰC/PHẨM CHẤT
# ==========================================
def get_vietnamese_sort_key(full_name):
    if not full_name: return ("", "")
    parts = full_name.strip().split()
    if not parts: return ("", "")
    return (parts[-1].lower(), " ".join(parts[:-1]).lower())


@app.route('/teacher/evaluation')
def teacher_evaluation():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    t_res = supabase.table("teachers").select("teacher_id").eq("user_id", session['user_id']).execute()
    teacher_id = t_res.data[0]['teacher_id'] if t_res.data else None

    # =================================================================
    # 🔥 BƯỚC 1: LẤY CÁC HỌC KỲ THUỘC NĂM HỌC GV ĐÓ LÀM CHỦ NHIỆM
    # =================================================================
    semesters = []

    # CHẶN BẮT BUỘC: Chỉ truy vấn khi teacher_id tồn tại và hợp lệ
    if teacher_id:
        # Đi xuyên qua bảng classes để lấy academic_year_id gốc nhằm đảm bảo tính chính xác bản ghi
        hr_assignments_all = supabase.table("homeroom_assignments") \
            .select("class_id, classes(academic_year_id)") \
            .eq("teacher_id", teacher_id) \
            .execute().data

        # Gom danh sách các academic_year_id duy nhất từ thông tin lớp thực tế
        year_ids = list(set([
            item['classes']['academic_year_id']
            for item in hr_assignments_all
            if item.get('classes') and item['classes'].get('academic_year_id')
        ]))

        if year_ids:
            # Lấy ra các học kỳ thuộc đúng (các) năm học mà GV này đang có lớp chủ nhiệm
            semesters = supabase.table("semesters") \
                .select("semester_id, semester_name, year_id, is_current") \
                .in_("year_id", year_ids) \
                .order("semester_id") \
                .execute().data

    # Xác định Học kỳ đang chọn (mặc định lấy kỳ hiện tại hoặc kỳ đầu tiên)
    selected_semester_id = request.args.get('semester_id')
    if selected_semester_id and not any(str(s['semester_id']) == str(selected_semester_id) for s in semesters):
        selected_semester_id = None

    if not selected_semester_id:
        current_sem = next((s for s in semesters if s.get('is_current')), None)
        selected_semester_id = str(current_sem['semester_id']) if current_sem else (
            str(semesters[0]['semester_id']) if semesters else '1'
        )

    # Tìm year_id của Học kỳ đang chọn để ràng buộc lớp chủ nhiệm của năm đó
    selected_sem_obj = next((s for s in semesters if str(s['semester_id']) == str(selected_semester_id)), None)
    selected_year_id = selected_sem_obj['year_id'] if selected_sem_obj else None

    # =================================================================
    # 2. LẤY DANH SÁCH LỚP CHỦ NHIỆM THEO NĂM HỌC ĐANG CHỌN
    # =================================================================
    query_assignments = supabase.table("homeroom_assignments") \
        .select("class_id, classes(class_name, grade_id, grades(grade_name))") \
        .eq("teacher_id", teacher_id)

    if selected_year_id:
        query_assignments = query_assignments.eq("year_id", selected_year_id)

    raw_assigned_classes = query_assignments.execute().data

    # 🔥 SỬA LỖI TẠI ĐÂY: Duyệt loại bỏ trùng lặp để loại bỏ lỗi hiện 2 chữ 1A
    seen_classes = set()
    assigned_classes = []
    for c in raw_assigned_classes:
        if c.get('classes') and c['class_id'] not in seen_classes:
            seen_classes.add(c['class_id'])
            assigned_classes.append(c)

    # Đọc tham số bộ lọc lớp và khối từ URL xuống
    selected_grade = request.args.get('grade')
    selected_class = request.args.get('class_id')

    # Khóa bảo vệ bộ lọc chéo năm học: Nếu lớp từ URL không thuộc năm học mới -> Xóa trắng lớp/khối
    if selected_class and not any(str(c['class_id']) == str(selected_class) for c in assigned_classes):
        selected_class = None
        selected_grade = None

    # 🔥 BƯỚC 3: LỌC DANH MỤC KHỐI DỰA TRÊN LỚP CHỦ NHIỆM THỰC TẾ ĐỂ ĐỔ LÊN DROPDOWN
    filter_grades = {}
    for a in assigned_classes:
        if a.get('classes') and a['classes'].get('grades'):
            g_id = a['classes']['grade_id']
            g_name = a['classes']['grades']['grade_name']
            filter_grades[g_id] = g_name

    selected_type = request.args.get('type') or "Nhận xét"
    selected_time = request.args.get('time') or "Giữa học kỳ"
    eval_times = ["Giữa học kỳ", "Cuối học kỳ"]

    # Bảo vệ lỗi cache của bảng rules đánh giá định kỳ
    try:
        rules_res = supabase.table("evaluation_rules").select("evaluation_type") \
            .in_("category", ["Năng lực cốt lõi", "Phẩm chất chủ yếu"]).execute()
        eval_rules = list(set([r['evaluation_type'] for r in rules_res.data]))
        if not eval_rules:
            eval_rules = ["Nhận xét"]
    except Exception:
        eval_rules = ["Nhận xét"]

    students_list, eval_map, class_name = [], {}, ""

    # =================================================================
    # 4. LOAD DỮ LIỆU HỌC SINH VÀ ĐIỂM ĐÁNH GIÁ THEO LỚP ĐÃ CHỌN
    # =================================================================
    if selected_class:
        target_class = next((c for c in assigned_classes if str(c['class_id']) == str(selected_class)), None)
        if target_class:
            class_name = target_class['classes']['class_name']

        st_res = supabase.table("student_enrollments").select(
            "student_id, students(ho_ten, ngay_sinh, ma_dinh_danh)").eq("class_id", selected_class).execute()

        sorted_data = sorted(st_res.data, key=lambda x: get_vietnamese_sort_key(x['students']['ho_ten']))

        student_ids = []
        for item in sorted_data:
            student = item['students']
            student['student_id'] = item['student_id']
            students_list.append(student)
            student_ids.append(item['student_id'])

        # Chỉ lấy bảng điểm của học sinh trong lớp chọn để tối ưu tốc độ hệ thống
        if student_ids:
            ev_res = supabase.table("periodic_evaluations").select("*") \
                .eq("semester_id", selected_semester_id) \
                .eq("evaluation_time", selected_time) \
                .eq("evaluation_type", selected_type) \
                .in_("student_id", student_ids).execute()
            ev_data = ev_res.data
        else:
            ev_data = []

        nl_chung = ['Tự chủ và tự học', 'Giao tiếp hợp tác', 'Giải quyết vấn đề và sáng tạo']
        nl_dacthu = ['Ngôn ngữ', 'Tính toán', 'Khoa học', 'Thẩm mỹ', 'Thể chất']

        for ev in ev_data:
            s_id = ev['student_id']
            crit = ev['criteria_name']
            raw_comment = ev['comment'] if ev['comment'] else ""

            if s_id not in eval_map:
                eval_map[s_id] = {}
            eval_map[s_id][crit] = ev['result_level']

            # Logic tách Mã [Mã] và Nội dung nhận xét đổ lên giao diện UI
            code = ""
            content = raw_comment
            if raw_comment.startswith("[") and "]" in raw_comment:
                code = raw_comment[1:raw_comment.find("]")]
                content = raw_comment[raw_comment.find("]") + 1:].strip()

            if crit in nl_chung:
                eval_map[s_id]['chung_code'] = code
                eval_map[s_id]['chung_text'] = content
            elif crit in nl_dacthu:
                eval_map[s_id]['dacthu_code'] = code
                eval_map[s_id]['dacthu_text'] = content

    return render_template('teacher/evaluation.html',
                           semesters=semesters,
                           selected_semester_id=selected_semester_id,
                           filter_grades=filter_grades,
                           assigned_classes=assigned_classes,
                           eval_rules=eval_rules,
                           eval_times=eval_times,
                           students=students_list,
                           eval_map=eval_map,
                           class_name=class_name,
                           selected_grade=selected_grade,
                           selected_class=selected_class,
                           selected_time=selected_time,
                           selected_type=selected_type)
# ==================================
# LƯU DỮ LIỆU
# ==================================
@app.route('/teacher/save_evaluation', methods=['POST'])
def save_evaluation():
    if 'user_id' not in session: return redirect(url_for('login'))

    # 1. Lấy thông tin giáo viên
    t_res = supabase.table("teachers").select("teacher_id").eq("user_id", session['user_id']).execute()
    if not t_res.data:
        return "Lỗi: Không tìm thấy thông tin giáo viên trong hệ thống."
    teacher_id = t_res.data[0]['teacher_id']

    form_data = request.form.to_dict()

    # 2. LẤY CÁC BIẾN QUAN TRỌNG TỪ Ô HIDDEN CỦA HTML
    # Cần lấy chính xác để khi redirect trang web không bị mất bộ lọc (Khối, Lớp, Học kỳ...)
    semester_id = form_data.get('semester_id_hidden')
    eval_time = form_data.get('evaluation_time_hidden')
    eval_type = form_data.get('evaluation_type_hidden')
    selected_grade = form_data.get('grade_hidden')
    selected_class = form_data.get('class_id')

    upsert_data = []

    # Danh sách tiêu chí để phân loại nhóm (Lưu ý: Phải khớp 100% với tên trong Database)
    nl_chung = ['Tự chủ và tự học', 'Giao tiếp hợp tác', 'Giải quyết vấn đề và sáng tạo']
    nl_dacthu = ['Ngôn ngữ', 'Tính toán', 'Khoa học', 'Thẩm mỹ', 'Thể chất']
    nang_luc_all = nl_chung + nl_dacthu

    # Tìm danh sách ID học sinh có trong form gửi lên
    student_ids = set([k.split('_')[1] for k in form_data.keys() if k.startswith('eval_')])

    for s_id in student_ids:
        # Lấy và làm sạch (strip) Mã và Nội dung nhận xét
        code_c = form_data.get(f'code_chung_{s_id}', '').strip()
        text_c = form_data.get(f'comment_chung_{s_id}', '').strip()
        code_dt = form_data.get(f'code_dacthu_{s_id}', '').strip()
        text_dt = form_data.get(f'comment_dacthu_{s_id}', '').strip()

        # Gộp mã và nội dung thành chuỗi lưu vào cột 'comment'
        full_comment_chung = f"[{code_c}] {text_c}" if code_c else text_c
        full_comment_dacthu = f"[{code_dt}] {text_dt}" if code_dt else text_dt

        # Duyệt qua từng tiêu chí của học sinh đó
        for key, value in form_data.items():
            if key.startswith(f'eval_{s_id}_'):
                crit_name = key.replace(f'eval_{s_id}_', '')

                # Chỉ xử lý nếu giáo viên có chọn mức độ (Tốt, Đạt, Cần cố gắng)
                if value and value.strip():
                    cat = 'Năng lực' if crit_name in nang_luc_all else 'Phẩm chất'

                    # Gán nhận xét tương ứng cho từng nhóm tiêu chí
                    final_comment = None
                    if crit_name in nl_chung:
                        final_comment = full_comment_chung
                    elif crit_name in nl_dacthu:
                        final_comment = full_comment_dacthu

                    upsert_data.append({
                        "student_id": int(s_id),
                        "semester_id": int(semester_id),
                        "teacher_id": teacher_id,
                        "category": cat,
                        "criteria_name": crit_name,
                        "result_level": value.strip(),  # Xóa dấu cách thừa nếu có
                        "comment": final_comment,
                        "evaluation_time": eval_time,
                        "evaluation_type": eval_type
                    })

    # 3. THỰC HIỆN LƯU VÀO DATABASE
    if upsert_data:
        try:
            # Sử dụng on_conflict để tự động cập nhật (Update) nếu đã có dữ liệu cũ
            supabase.table("periodic_evaluations").upsert(
                upsert_data,
                on_conflict="student_id, semester_id, criteria_name, evaluation_time, evaluation_type"
            ).execute()
            msg = "Đã lưu dữ liệu thành công!"
        except Exception as e:
            msg = f"Lỗi Database: {str(e)}"
    else:
        msg = "Chưa có dữ liệu để lưu (Vui lòng chọn mức độ đánh giá)."

    # 4. QUAY LẠI TRANG VỚI ĐẦY ĐỦ THÔNG SỐ ĐỂ GIỮ NGUYÊN GIAO DIỆN HIỆN TẠI
    return redirect(url_for('teacher_evaluation',
                            grade=selected_grade,
                            class_id=selected_class,
                            semester_id=semester_id,
                            time=eval_time,
                            type=eval_type,
                            msg=msg))


# =====================
# XUẤT EXCEL
# =====================
@app.route('/teacher/export_evaluation')
def export_evaluation():
    if 'user_id' not in session: return redirect(url_for('login'))

    class_id = request.args.get('class_id')
    semester_id = request.args.get('semester_id')
    eval_time = request.args.get('time')
    eval_type = request.args.get('type') or "Nhận xét"

    if not all([class_id, semester_id]):
        return "Vui lòng chọn đầy đủ Lớp và Học kỳ trước khi xuất!"

    # --- TRUY VẤN THÔNG TIN BỔ TRỢ ---
    cls_res = supabase.table("classes").select("class_name").eq("class_id", class_id).single().execute().data
    sem_res = supabase.table("semesters").select("semester_name").eq("semester_id", semester_id).single().execute().data

    # --- BƯỚC 1: ĐỊNH NGHĨA TIÊU CHÍ ---
    nl_chung = ['Tự chủ và tự học', 'Giao tiếp hợp tác', 'Giải quyết vấn đề và sáng tạo']
    nl_dacthu = ['Ngôn ngữ', 'Tính toán', 'Khoa học', 'Thẩm mỹ', 'Thể chất']
    pham_chat = ['Yêu nước', 'Nhân ái', 'Chăm chỉ', 'Trung thực', 'Trách nhiệm']
    all_ordered_criteria = nl_chung + nl_dacthu + pham_chat

    # --- BƯỚC 2: LẤY DANH SÁCH HỌC SINH ---
    # Thêm ma_dinh_danh vào lệnh select bên dưới
    st_res = supabase.table("student_enrollments") \
        .select("student_id, students(ho_ten, ngay_sinh, ma_dinh_danh)") \
        .eq("class_id", class_id).execute()

    if not st_res.data: return "Lớp này không có học sinh!"

    students_data = [
        {"student_id": i['student_id'], "ho_ten": i['students']['ho_ten'],
         "ma_dinh_danh": i['students']['ma_dinh_danh'], "ngay_sinh": i['students']['ngay_sinh']}
        for i in st_res.data
    ]
    students_data = sorted(students_data, key=lambda x: get_vietnamese_sort_key(x['ho_ten']))
    df_final = pd.DataFrame(students_data)

    # --- BƯỚC 3: LẤY VÀ XỬ LÝ DỮ LIỆU ĐÁNH GIÁ ---
    ev_res = supabase.table("periodic_evaluations").select("student_id, criteria_name, result_level, comment").eq(
        "semester_id", semester_id).eq("evaluation_time", eval_time).eq("evaluation_type", eval_type).execute()

    if ev_res.data:
        raw_ev = pd.json_normalize(ev_res.data)
        df_levels = raw_ev.pivot_table(index='student_id', columns='criteria_name', values='result_level',
                                       aggfunc='first').reset_index()
        comment_records = []
        for sid in raw_ev['student_id'].unique():
            s_comments = raw_ev[raw_ev['student_id'] == sid]

            def get_split_comment(group_list):
                row = s_comments[s_comments['criteria_name'].isin(group_list)]
                full_c = row['comment'].iloc[0] if not row.empty and row['comment'].iloc[0] else ""
                if full_c.startswith("[") and "]" in full_c:
                    return full_c[1:full_c.find("]")], full_c[full_c.find("]") + 1:].strip()
                return "", full_c

            code_c, text_c = get_split_comment(nl_chung);
            code_dt, text_dt = get_split_comment(nl_dacthu)
            comment_records.append(
                {'student_id': sid, 'Mã nhận xét C': code_c, 'Nội dung C': text_c, 'Mã nhận xét DT': code_dt,
                 'Nội dung DT': text_dt})
        df_final = pd.merge(df_final, df_levels, on='student_id', how='left')
        df_final = pd.merge(df_final, pd.DataFrame(comment_records), on='student_id', how='left')

    df_final.insert(0, 'STT', range(1, len(df_final) + 1))

    # Thêm 'ma_dinh_danh' vào giữa ho_ten và ngay_sinh
    final_columns = ['STT', 'ho_ten', 'ma_dinh_danh', 'ngay_sinh'] + all_ordered_criteria + \
                    ['Mã nhận xét C', 'Nội dung C', 'Mã nhận xét DT', 'Nội dung DT']

    df_export = df_final.reindex(columns=final_columns)

    # --- BƯỚC 4: XUẤT EXCEL TỐI GIẢN ---
    def no_accent_vietnamese(s):

        if not s: return ""

        s = re.sub(r'[àáạảãâầấậẩẫăằắặẳẵ]', 'a', s)

        s = re.sub(r'[ÀÁẠẢÃÂẦẤẬẨẪĂẰẮẶẲẴ]', 'A', s)

        s = re.sub(r'[èéẹẻẽêềếệểễ]', 'e', s)

        s = re.sub(r'[ÈÉẸẺẼÊỀẾỆỂỄ]', 'E', s)

        s = re.sub(r'[òóọỏõôồốộổỗơờớợởỡ]', 'o', s)

        s = re.sub(r'[ÒÓỌỎÕÔỒỐỘỔỖƠỜỚỢỞỠ]', 'O', s)

        s = re.sub(r'[ìíịỉĩ]', 'i', s)

        s = re.sub(r'[ÌÍỊỈĨ]', 'I', s)

        s = re.sub(r'[ùúụủũưừứựửữ]', 'u', s)

        s = re.sub(r'[ÙÚỤỦŨƯỪỨỰỬỮ]', 'U', s)

        s = re.sub(r'[ỳýỵỷỹ]', 'y', s)

        s = re.sub(r'[ỲÝỴỶỸ]', 'Y', s)

        s = re.sub(r'[đ]', 'd', s)

        s = re.sub(r'[Đ]', 'D', s)

        return s.replace(" ", "_")

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. Xuất dữ liệu từ dòng 6 (A6 -> U...)
        df_export.to_excel(writer, index=False, sheet_name='Export', startrow=5, header=False)
        ws = writer.sheets['Export']

        # Định nghĩa các Style cơ bản (Chỉ Font, Căn lề và Khung)
        bold_font = Font(bold=True, name='Arial', size=11)
        title_font = Font(bold=True, name='Arial', size=14)
        center_wrap = Alignment(horizontal='center', vertical='center', wrap_text=True)
        left_wrap = Alignment(horizontal='left', vertical='top', wrap_text=True)
        border = Border(left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))

        # --- DÒNG 1: TIÊU ĐỀ CHÍNH (Gộp 21 cột A -> U) ---
        ws.merge_cells('A1:U1')
        ws['A1'] = f"BẢNG ĐÁNH GIÁ ĐỊNH KỲ NĂNG LỰC VÀ PHẨM CHẤT - {eval_type.upper()}"
        ws['A1'].font = title_font
        ws['A1'].alignment = center_wrap
        ws.row_dimensions[1].height = 30

        # --- DÒNG 2: THÔNG TIN LỚP/KỲ ---
        ws.merge_cells('A2:D2')
        ws['A2'] = f"Lớp: {cls_res['class_name']}"
        ws.merge_cells('E2:U2')
        ws['E2'] = f"Học kỳ: {sem_res['semester_name']} - Thời điểm: {eval_time}"
        ws['A2'].font = bold_font
        ws['E2'].font = bold_font

        # --- DÒNG 3-5: HEADER 3 TẦNG (CĂN CHỈNH TỪ CỘT A -> U) ---
        # 4 cột cố định đầu tiên
        ws.merge_cells('A3:A5');
        ws['A3'] = 'STT'
        ws.merge_cells('B3:B5');
        ws['B3'] = 'Họ tên'
        ws.merge_cells('C3:C5');
        ws['C3'] = 'Mã định danh Bộ GD&ĐT'
        ws.merge_cells('D3:D5');
        ws['D3'] = 'Ngày sinh'

        # Nhóm Năng lực cốt lõi (E -> L: 8 cột)
        ws.merge_cells('E3:L3');
        ws['E3'] = 'Năng lực cốt lõi'
        ws.merge_cells('E4:G4');
        ws['E4'] = 'Năng lực chung'
        ws.merge_cells('H4:L4');
        ws['H4'] = 'Năng lực đặc thù'

        # Nhóm Phẩm chất (M -> Q: 5 cột)
        ws.merge_cells('M3:Q4');
        ws['M3'] = 'Phẩm chất chủ yếu'

        # Nhóm Nhận xét (R -> U: 4 cột)
        ws.merge_cells('R3:S4');
        ws['R3'] = 'Nhận xét năng lực chung'
        ws.merge_cells('T3:U4');
        ws['T3'] = 'Nhận xét năng lực đặc thù'

        # DÒNG 5: Tên các tiêu chí (Bắt đầu từ cột E - index 5)
        detail_headers = nl_chung + nl_dacthu + pham_chat + ['Mã nhận xét', 'Nội dung', 'Mã nhận xét', 'Nội dung']
        for i, h in enumerate(detail_headers):
            cell = ws.cell(row=5, column=5 + i)
            cell.value = h

        # --- THIẾT LẬP KÍCH THƯỚC Ô ---
        ws.column_dimensions['A'].width = 5  # STT
        ws.column_dimensions['B'].width = 25  # Họ tên
        ws.column_dimensions['C'].width = 18  # Mã định danh (Rộng hơn chút để hiện đủ mã)
        ws.column_dimensions['D'].width = 12  # Ngày sinh
        ws.column_dimensions['S'].width = 45  # Nội dung Chung
        ws.column_dimensions['U'].width = 45  # Nội dung Đặc thù
        for col_l in ['E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'T']:
            ws.column_dimensions[col_l].width = 10

        # --- ĐỊNH DẠNG TOÀN BẢNG (KHÔNG MÀU MÈ) ---
        for r in range(1, ws.max_row + 1):
            if r >= 6:
                ws.row_dimensions[r].height = 45  # Cho hàng cao lên để hiện đủ nội dung nhận xét

            for c in range(1, 22):
                cell = ws.cell(row=r, column=c)
                cell.border = border
                if r <= 5:  # Header
                    cell.alignment = center_wrap
                    cell.font = bold_font
                else:  # Dữ liệu
                    if c in [19, 21]:  # Nội dung căn lề trái, trên cùng
                        cell.alignment = left_wrap
                    else:
                        cell.alignment = center_wrap

    output.seek(0)
    safe_class = no_accent_vietnamese(cls_res['class_name'])
    return send_file(output, download_name=f"DanhGiaNLPC_{safe_class}.xlsx", as_attachment=True)
# =================================
# NHẬP ĐGĐK NLPC TỪ EXCEL
# =================================
@app.route('/teacher/import_evaluation', methods=['POST'])
def import_evaluation():
    if 'user_id' not in session: return redirect(url_for('login'))
    if 'file' not in request.files: return redirect(request.referrer)

    file = request.files['file']
    if file.filename == '': return redirect(request.referrer)

    semester_id = request.form.get('semester_id')
    eval_time = request.form.get('time')
    eval_type = request.form.get('type')
    class_id = request.form.get('class_id')
    grade = request.form.get('grade')

    t_res = supabase.table("teachers").select("teacher_id").eq("user_id", session['user_id']).execute()
    teacher_id = t_res.data[0]['teacher_id']

    try:
        # 1. ĐỌC THÔNG TIN NGỮ CẢNH ĐỂ KIỂM TRA BẢO MẬT
        check_df = pd.read_excel(file, nrows=2, header=None)
        f_class_name = str(check_df.iloc[1, 0]).replace("Lớp:", "").strip().upper()
        # Đọc ô E2 (index 4) vì file 21 cột của bạn gộp A2:D2
        f_header_info = str(check_df.iloc[1, 4]).upper()

        db_cls = supabase.table("classes").select("class_name").eq("class_id", class_id).single().execute().data[
            'class_name'].upper()
        db_sem = \
        supabase.table("semesters").select("semester_name").eq("semester_id", semester_id).single().execute().data[
            'semester_name'].upper()
        db_time = eval_time.upper()

        if f_class_name != db_cls or db_sem not in f_header_info or db_time not in f_header_info:
            msg = f"Từ chối: File này không khớp với bộ lọc bạn đang chọn (Lớp, Kỳ hoặc Thời điểm)!"
            return redirect(url_for('teacher_evaluation', grade=grade, class_id=class_id,
                                    semester_id=semester_id, time=eval_time, type=eval_type, msg=msg))

        # 2. ĐỌC DỮ LIỆU TỪ DÒNG 6 (skip 5 dòng header)
        df = pd.read_excel(file, skiprows=5, header=None)

        nl_chung = ['Tự chủ và tự học', 'Giao tiếp hợp tác', 'Giải quyết vấn đề và sáng tạo']
        nl_dacthu = ['Ngôn ngữ', 'Tính toán', 'Khoa học', 'Thẩm mỹ', 'Thể chất']
        pham_chat = ['Yêu nước', 'Nhân ái', 'Chăm chỉ', 'Trung thực', 'Trách nhiệm']
        all_criteria = nl_chung + nl_dacthu + pham_chat

        upsert_data = []

        for _, row in df.iterrows():
            # Cột C (index 2) là Mã định danh
            mdd = str(row[2]).strip()
            if not mdd or mdd == 'nan': continue

            # Tìm student_id theo Mã định danh
            st_res = supabase.table("students").select("student_id").eq("ma_dinh_danh", mdd).execute()
            if not st_res.data: continue
            s_id = st_res.data[0]['student_id']

            # --- SỬA TÊN BIẾN ĐỒNG NHẤT ĐỂ TRÁNH LỖI DEFINED ---
            # Nhận xét Chung: Cột R(17) và S(18)
            code_c = str(row[17]).strip() if pd.notna(row[17]) else ""
            text_c = str(row[18]).strip() if pd.notna(row[18]) else ""
            full_comment_chung = f"[{code_c}] {text_c}" if code_c and code_c != 'nan' else text_c

            # Nhận xét Đặc thù: Cột T(19) và U(20)
            code_dt = str(row[19]).strip() if pd.notna(row[19]) else ""
            text_dt = str(row[20]).strip() if pd.notna(row[20]) else ""
            full_comment_dacthu = f"[{code_dt}] {text_dt}" if code_dt and code_dt != 'nan' else text_dt

            # Duyệt các cột điểm (Bắt đầu từ cột E - index 4)
            for i, crit in enumerate(all_criteria):
                val = str(row[4 + i]).strip()
                if val in ['Tốt', 'Đạt', 'Cần cố gắng']:
                    cat = 'Năng lực' if i < 8 else 'Phẩm chất'

                    # Gán đúng tên biến đã khai báo ở trên
                    f_comment = None
                    if crit in nl_chung:
                        f_comment = full_comment_chung
                    elif crit in nl_dacthu:
                        f_comment = full_comment_dacthu

                    upsert_data.append({
                        "student_id": s_id, "semester_id": int(semester_id), "teacher_id": teacher_id,
                        "category": cat, "criteria_name": crit, "result_level": val,
                        "comment": f_comment, "evaluation_time": eval_time, "evaluation_type": eval_type
                    })

        if upsert_data:
            supabase.table("periodic_evaluations").upsert(
                upsert_data, on_conflict="student_id, semester_id, criteria_name, evaluation_time, evaluation_type"
            ).execute()
            msg = f"Thành công: Đã nạp dữ liệu đánh giá cho lớp {db_cls}!"
        else:
            msg = "Thất bại: Không tìm thấy dữ liệu đánh giá hợp lệ trong file."

    except Exception as e:
        msg = f"Lỗi hệ thống: {str(e)}"

    return redirect(url_for('teacher_evaluation', grade=grade, class_id=class_id,
                            semester_id=semester_id, time=eval_time, type=eval_type, msg=msg))


# ======================
# HELPER: LẤY THÔNG TIN GIÁO VIÊN & LỚP CHỦ NHIỆM
# ======================
def get_teacher_homeroom(user_id):
    t_res = supabase.table("teachers").select("teacher_id").eq("user_id", user_id).execute()
    teacher_id = t_res.data[0]['teacher_id'] if t_res.data else None

    # Lấy năm học hiện tại
    cur_year = supabase.table("academic_years").select("*").eq("is_current", True).execute().data
    year_id = cur_year[0]['year_id'] if cur_year else None

    # Lấy học kỳ hiện tại (để dùng cho điểm danh, kế hoạch GD)
    cur_sem = supabase.table("semesters").select("*").eq("is_current", True).execute().data
    semester_id = cur_sem[0]['semester_id'] if cur_sem else None

    assignments = []
    if teacher_id and year_id:
        assignments = supabase.table("homeroom_assignments") \
            .select("class_id, classes(class_id, class_name, grade_id, grades(grade_id, grade_name))") \
            .eq("teacher_id", teacher_id) \
            .eq("year_id", year_id) \
            .execute().data

    grades, classes = {}, []
    for a in assignments:
        cls = a.get('classes', {})
        grd = cls.get('grades', {})
        if grd:
            grades[grd['grade_id']] = grd['grade_name']
        classes.append({
            'class_id':   cls['class_id'],
            'class_name': cls['class_name'],
            'grade_id':   cls['grade_id']
        })

    grades_list = [{'grade_id': k, 'grade_name': v} for k, v in grades.items()]
    return teacher_id, semester_id, grades_list, classes


# ======================
# ĐIỂM DANH (CHUYÊN CẦN)
# ======================
@app.route('/teacher/diem-danh', methods=['GET', 'POST'])
def diem_danh():
    if 'user_id' not in session or session.get('role_name') != 'Teacher':
        return redirect(url_for('login'))

    teacher_id, semester_id, grades, classes = get_teacher_homeroom(session['user_id'])

    msg = error_msg = None

    if request.method == 'POST':
        action = request.form.get('action')

        if action == 'add':
            try:
                supabase.table("attendance").insert({
                    "student_id":   int(request.form['student_id']),
                    "class_id":     int(request.form['class_id']),
                    "semester_id":  semester_id,
                    "absence_date": request.form['absence_date'],
                    "session":      request.form.get('session', 'Sáng'),
                    "is_permitted": request.form.get('is_permitted') == 'true',
                    "reason":       request.form.get('reason', '').strip() or None,
                }).execute()
                msg = 'Thêm vắng học thành công!'
            except Exception as e:
                error_msg = f'Lỗi: {str(e)}'

        elif action == 'edit':
            try:
                att_id = int(request.form['attendance_id'])
                supabase.table("attendance").update({
                    "student_id":   int(request.form['student_id']),
                    "class_id":     int(request.form['class_id']),
                    "absence_date": request.form['absence_date'],
                    "session":      request.form.get('session', 'Sáng'),
                    "is_permitted": request.form.get('is_permitted') == 'true',
                    "reason":       request.form.get('reason', '').strip() or None,
                }).eq('attendance_id', att_id).execute()
                msg = 'Cập nhật thành công!'
            except Exception as e:
                error_msg = f'Lỗi: {str(e)}'

        elif action == 'delete':
            ids_str = request.form.get('ids', '')
            ids = [int(x) for x in ids_str.split(',') if x.strip().isdigit()]
            if ids:
                supabase.table("attendance").delete().in_('attendance_id', ids).execute()
                msg = f'Đã xóa {len(ids)} bản ghi!'

    # GET params
    selected_grade = request.args.get('grade_id', type=int)
    selected_class = request.args.get('class_id', type=int)
    from_date      = request.args.get('from_date', '')
    to_date        = request.args.get('to_date', '')

    # Lọc classes theo grade
    filtered_classes = [c for c in classes if not selected_grade or c['grade_id'] == selected_grade]

    # Lấy học sinh của lớp chọn
    students = []
    if selected_class:
        enr = supabase.table("student_enrollments") \
            .select("student_id, students(student_id, ho_ten, ma_dinh_danh)") \
            .eq("class_id", selected_class).execute().data
        students = [e['students'] for e in enr if e.get('students')]

    # Lấy điểm danh
    attendances = []
    raw = []
    
    # CHỈ LẤY DỮ LIỆU KHI GIÁO VIÊN ĐÃ BẤM CHỌN LỚP CỤ THỂ
    if selected_class:
        q = supabase.table("attendance") \
            .select("*, students(ho_ten, ma_dinh_danh, ngay_sinh, gioi_tinh), classes(class_name)") \
            .eq("semester_id", semester_id) \
            .eq("class_id", selected_class)
            
        if from_date:
            q = q.gte("absence_date", from_date)
        if to_date:
            q = q.lte("absence_date", to_date)
            
        raw = q.order("absence_date", desc=True).execute().data

    for r in raw:
        st = r.pop('students', {}) or {}
        cls = r.pop('classes', {}) or {}
        attendances.append({**r,
            'ho_ten': st.get('ho_ten',''), 'ma_dinh_danh': st.get('ma_dinh_danh',''),
            'ngay_sinh': st.get('ngay_sinh',''), 'gioi_tinh': st.get('gioi_tinh',''),
            'class_name': cls.get('class_name',''),
        })

    return render_template('teacher/diem_danh.html',
        grades=grades, classes=filtered_classes, students=students,
        attendances=attendances,
        selected_grade=selected_grade, selected_class=selected_class,
        from_date=from_date, to_date=to_date,
        msg=msg, error_msg=error_msg)

# ======================
# XUẤT EXCEL ĐIỂM DANH
# ======================
@app.route('/teacher/diem-danh/export')
def export_diem_danh():
    if 'user_id' not in session or session.get('role_name') != 'Teacher':
        return redirect(url_for('login'))
 
    teacher_id, semester_id, grades, classes = get_teacher_homeroom(session['user_id'])
    class_id = request.args.get('class_id', type=int)
 
    q = supabase.table("attendance") \
        .select("*, students(ho_ten, ma_dinh_danh, ngay_sinh, gioi_tinh), classes(class_name)") \
        .eq("semester_id", semester_id)
    if class_id:
        q = q.eq("class_id", class_id)
    elif classes:
        q = q.in_("class_id", [c['class_id'] for c in classes])
    rows = q.order("absence_date", desc=True).execute().data
 
    wb = Workbook()
    ws = wb.active
    ws.title = "Điểm danh"
 
    # Style
    header_font  = Font(bold=True, color="FFFFFF")
    header_fill  = __import__('openpyxl').styles.PatternFill("solid", fgColor="1a3a6b")
    center       = Alignment(horizontal="center", vertical="center")
    thin         = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
 
    headers = ["STT", "Mã định danh", "Họ tên", "Ngày sinh", "Giới tính",
               "Lớp", "Ngày nghỉ", "Buổi nghỉ", "Có phép", "Lý do"]
    col_widths = [6, 18, 22, 14, 10, 8, 14, 10, 10, 30]
 
    ws.row_dimensions[1].height = 20
    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font   = header_font
        cell.fill   = header_fill
        cell.alignment = center
        cell.border = thin
        ws.column_dimensions[get_column_letter(ci)].width = w
 
    for ri, r in enumerate(rows, 2):
        st  = r.get('students') or {}
        cls = r.get('classes')  or {}
        ngay_sinh = st.get('ngay_sinh', '')
        if ngay_sinh:
            try: ngay_sinh = datetime.strptime(str(ngay_sinh)[:10], '%Y-%m-%d').strftime('%d/%m/%Y')
            except: pass
        absence = r.get('absence_date', '')
        if absence:
            try: absence = datetime.strptime(str(absence)[:10], '%Y-%m-%d').strftime('%d/%m/%Y')
            except: pass
 
        row_data = [
            ri - 1,
            st.get('ma_dinh_danh', ''),
            st.get('ho_ten', ''),
            ngay_sinh,
            st.get('gioi_tinh', ''),
            cls.get('class_name', ''),
            absence,
            r.get('session', ''),
            'Có phép' if r.get('is_permitted') else 'Không phép',
            r.get('reason', '') or '',
        ]
        for ci, val in enumerate(row_data, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.alignment = Alignment(horizontal="center" if ci != 3 and ci != 10 else "left", vertical="center")
            cell.border = thin
 
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    today = datetime.now().strftime('%d-%m-%Y')
    return send_file(output,
                     download_name=f"DiemDanh_{today}.xlsx",
                     as_attachment=True,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
 
 
# ======================
# TẢI FILE MẪU ĐIỂM DANH
# ======================
@app.route('/teacher/diem-danh/template')
def download_diem_danh_template():
    if 'user_id' not in session:
        return redirect(url_for('login'))
 
    wb = Workbook()
    ws = wb.active
    ws.title = "Mẫu điểm danh"
 
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = __import__('openpyxl').styles.PatternFill("solid", fgColor="1a3a6b")
    red_font    = Font(bold=True, color="FF0000")
    center      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin        = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
 
    headers   = ["STT", "Mã định danh (*)", "Họ tên", "Ngày nghỉ (*)\ndd/mm/yyyy",
                 "Buổi nghỉ (*)\nSáng/Chiều/Cả ngày", "Có phép\nCó/Không", "Lý do nghỉ"]
    col_widths = [6, 20, 22, 18, 22, 12, 30]
    required   = {1, 3}  # index của cột bắt buộc (0-based)
 
    ws.row_dimensions[1].height = 36
    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font      = Font(bold=True, color="FFFFFF") if (ci - 1) not in required else Font(bold=True, color="FFFF00")
        cell.fill      = header_fill
        cell.alignment = center
        cell.border    = thin
        ws.column_dimensions[get_column_letter(ci)].width = w
 
    # 3 dòng mẫu
    samples = [
        [1, "HS001", "Nguyễn Văn A", "15/05/2026", "Sáng", "Có", "Ốm"],
        [2, "HS002", "Trần Thị B",   "16/05/2026", "Cả ngày", "Không", ""],
    ]
    for ri, row in enumerate(samples, 2):
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.alignment = Alignment(horizontal="center" if ci != 3 and ci != 7 else "left")
            cell.border = thin
 
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output,
                     download_name="Mau_DiemDanh.xlsx",
                     as_attachment=True,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
 
 
# ======================
# NHẬP EXCEL ĐIỂM DANH
# ======================
@app.route('/teacher/diem-danh/import', methods=['POST'])
def import_diem_danh():
    if 'user_id' not in session or session.get('role_name') != 'Teacher':
        return redirect(url_for('login'))
 
    teacher_id, semester_id, grades, classes = get_teacher_homeroom(session['user_id'])
    class_id = request.form.get('class_id', type=int)
 
    file = request.files.get('file')
    if not file or not file.filename.endswith(('.xlsx', '.xls')):
        return redirect(url_for('diem_danh',
                                class_id=class_id,
                                error_msg='Vui lòng chọn file Excel (.xlsx)!'))
 
    try:
        df = pd.read_excel(file, header=0)
        df.columns = [str(c).strip() for c in df.columns]
 
        inserted = 0
        errors   = []
 
        for idx, row in df.iterrows():
            ma_dd = str(row.get('Mã định danh (*)', '') or '').strip()
            if not ma_dd or ma_dd == 'nan':
                continue
 
            # Tìm student_id
            st_res = supabase.table("students").select("student_id") \
                .eq("ma_dinh_danh", ma_dd).execute().data
            if not st_res:
                errors.append(f"Dòng {idx+2}: Không tìm thấy HS mã '{ma_dd}'")
                continue
 
            student_id = st_res[0]['student_id']
 
            # Parse ngày
            ngay_nghi = str(row.get('Ngày nghỉ (*)\ndd/mm/yyyy', '') or '').strip()
            try:
                absence_date = datetime.strptime(ngay_nghi, '%d/%m/%Y').strftime('%Y-%m-%d')
            except:
                errors.append(f"Dòng {idx+2}: Ngày nghỉ sai định dạng '{ngay_nghi}'")
                continue
 
            buoi = str(row.get('Buổi nghỉ (*)\nSáng/Chiều/Cả ngày', 'Sáng') or 'Sáng').strip()
            co_phep_raw = str(row.get('Có phép\nCó/Không', 'Không') or 'Không').strip().lower()
            co_phep = co_phep_raw in ('có', 'co', 'true', '1', 'yes')
            ly_do   = str(row.get('Lý do nghỉ', '') or '').strip() or None
 
            # Lấy class_id từ file hoặc filter
            target_class = class_id
            if not target_class and classes:
                target_class = classes[0]['class_id']
 
            supabase.table("attendance").insert({
                "student_id":   student_id,
                "class_id":     target_class,
                "semester_id":  semester_id,
                "absence_date": absence_date,
                "session":      buoi,
                "is_permitted": co_phep,
                "reason":       ly_do,
            }).execute()
            inserted += 1
 
        msg = f'Nhập thành công {inserted} bản ghi!'
        if errors:
            msg += f' ({len(errors)} lỗi: ' + '; '.join(errors[:3]) + ('...' if len(errors) > 3 else '') + ')'
 
        return redirect(url_for('diem_danh', class_id=class_id, msg=msg))
 
    except Exception as e:
        return redirect(url_for('diem_danh', class_id=class_id,
                                error_msg=f'Lỗi đọc file: {str(e)}'))
    
# ======================
# KẾ HOẠCH GIÁO DỤC
# ======================
@app.route('/teacher/ke-hoach-gd', methods=['GET', 'POST'])
def ke_hoach_gd():
    if 'user_id' not in session or session.get('role_name') != 'Teacher':
        return redirect(url_for('login'))
 
    teacher_id, semester_id, grades, classes = get_teacher_homeroom(session['user_id'])
    msg = error_msg = None
 
    selected_grade = request.args.get('grade_id', type=int)
    selected_class = request.args.get('class_id', type=int)
    filtered_classes = [c for c in classes if not selected_grade or c['grade_id'] == selected_grade]
 
    if request.method == 'POST':
        class_id = request.form.get('class_id', type=int)
        grade_id = request.form.get('grade_id', type=int)
        # Lấy tất cả subject_id từ form
        subject_ids = set()
        for key in request.form:
            for prefix in ['hk1_total_', 'hk1_9_10_', 'hk1_7_8_', 'hk1_5_6_', 'hk1_u5_',
                           'cn_total_', 'cn_9_10_', 'cn_7_8_', 'cn_5_6_', 'cn_u5_']:
                if key.startswith(prefix):
                    subject_ids.add(int(key[len(prefix):]))
        try:
            for sid in subject_ids:
                def g(name): return int(request.form.get(f'{name}_{sid}') or 0)
                payload = {
                    'class_id': class_id, 'subject_id': sid, 'semester_id': semester_id,
                    'hk1_total':   g('hk1_total'), 'hk1_9_10': g('hk1_9_10'),
                    'hk1_7_8':    g('hk1_7_8'),    'hk1_5_6':  g('hk1_5_6'),
                    'hk1_under5': g('hk1_u5'),
                    'ca_nam_total':   g('cn_total'), 'ca_nam_9_10': g('cn_9_10'),
                    'ca_nam_7_8':    g('cn_7_8'),    'ca_nam_5_6':  g('cn_5_6'),
                    'ca_nam_under5': g('cn_u5'),
                }
                # Upsert
                existing = supabase.table("score_summary") \
                    .select("summary_id") \
                    .eq("class_id", class_id).eq("subject_id", sid).eq("semester_id", semester_id) \
                    .execute().data
                if existing:
                    supabase.table("score_summary").update(payload) \
                        .eq("summary_id", existing[0]['summary_id']).execute()
                else:
                    supabase.table("score_summary").insert(payload).execute()
            msg = 'Lưu kế hoạch thành công!'
            selected_class = class_id
            selected_grade = grade_id
        except Exception as e:
            error_msg = f'Lỗi: {str(e)}'
 
    # Tự động tính kế hoạch từ subject_scores_v2
    subject_rows = []
    score_map = {}
    if selected_class:
        cls_info = next((c for c in classes if c['class_id'] == selected_class), None)
        if cls_info:
            subject_rows = supabase.table("subjects") \
                .select("subject_id, subject_name, evaluation_type") \
                .eq("grade_id", cls_info['grade_id']) \
                .eq("status", "Đang áp dụng") \
                .order("subject_id").execute().data
 
            # Lấy danh sách học sinh trong lớp
            enr = supabase.table("student_enrollments") \
                .select("student_id") \
                .eq("class_id", selected_class).execute().data
            student_ids = [e['student_id'] for e in enr if e.get('student_id')]
 
            if student_ids:
                subject_ids = [r['subject_id'] for r in subject_rows]
 
                # Lấy toàn bộ điểm của học sinh trong lớp
                raw_scores = supabase.table("subject_scores_v2") \
                    .select("student_id, subject_id, semester_id, dtb_hk1, dtb_hk2, dtb_cn") \
                    .in_("student_id", student_ids) \
                    .in_("subject_id", subject_ids) \
                    .execute().data
 
                def count_range(scores, lo, hi):
                    return sum(1 for x in scores if lo <= x <= hi)
 
                for subj in subject_rows:
                    sid = subj['subject_id']
                    is_grading = subj.get('evaluation_type') == 'Chấm điểm'
                    subj_scores = [s for s in raw_scores if s['subject_id'] == sid]
 
                    hk1_scores = [s['dtb_hk1'] for s in subj_scores if s.get('dtb_hk1') is not None]
                    cn_scores  = [s['dtb_cn']  for s in subj_scores if s.get('dtb_cn')  is not None]
                    hk1_total  = len(hk1_scores)
                    cn_total   = len(cn_scores)
 
                    if is_grading:
                        score_map[sid] = {
                            'hk1_total':     hk1_total,
                            'hk1_9_10':      count_range(hk1_scores, 9, 10),
                            'hk1_7_8':       count_range(hk1_scores, 7, 8.99),
                            'hk1_5_6':       count_range(hk1_scores, 5, 6.99),
                            'hk1_under5':    count_range(hk1_scores, 0, 4.99),
                            'ca_nam_total':  cn_total,
                            'ca_nam_9_10':   count_range(cn_scores, 9, 10),
                            'ca_nam_7_8':    count_range(cn_scores, 7, 8.99),
                            'ca_nam_5_6':    count_range(cn_scores, 5, 6.99),
                            'ca_nam_under5': count_range(cn_scores, 0, 4.99),
                        }
                    else:
                        # Môn đánh giá: 10=Đạt, 0=Chưa đạt
                        score_map[sid] = {
                            'hk1_total':     hk1_total,
                            'hk1_9_10':      sum(1 for x in hk1_scores if x == 10.0),
                            'hk1_7_8':       0,
                            'hk1_5_6':       0,
                            'hk1_under5':    sum(1 for x in hk1_scores if x == 0.0),
                            'ca_nam_total':  cn_total,
                            'ca_nam_9_10':   sum(1 for x in cn_scores if x == 10.0),
                            'ca_nam_7_8':    0,
                            'ca_nam_5_6':    0,
                            'ca_nam_under5': sum(1 for x in cn_scores if x == 0.0),
                        }
 
    return render_template('teacher/ke_hoach_gd.html',
        grades=grades, classes=filtered_classes,
        selected_grade=selected_grade, selected_class=selected_class,
        subject_rows=subject_rows, score_map=score_map,
        msg=msg, error_msg=error_msg)


# ======================
# LẤY DANH SÁCH NĂM TRƯỚC (danh_sach_hs) - DÙNG CHO LẤY KẾT QUẢ HỌC TẬP NĂM TRƯỚC
# ======================
@app.route('/teacher/danh-sach-hs/lay-ket-qua')
def danh_sach_lay_ket_qua():
    if 'user_id' not in session or session.get('role_name') != 'Teacher':
        return redirect(url_for('login'))

    class_id = request.args.get('class_id', type=int)
    if not class_id:
        return redirect(url_for('danh_sach_hs', error_msg='Chưa chọn lớp!'))

    try:
        # Lấy năm học hiện tại
        cur_year = supabase.table("academic_years").select("*").eq("is_current", True).execute().data
        if not cur_year:
            return redirect(url_for('danh_sach_hs', class_id=class_id,
                                    error_msg='Không tìm thấy năm học hiện tại!'))
        prev_year_id = cur_year[0]['year_id'] - 1

        # Lấy enrollments lớp hiện tại
        cur_enr = supabase.table("student_enrollments").select("student_id") \
            .eq("class_id", class_id).execute().data
        student_ids = [e['student_id'] for e in cur_enr]

        if not student_ids:
            return redirect(url_for('danh_sach_hs', class_id=class_id,
                                    error_msg='Lớp chưa có học sinh!'))

        # Lấy kết quả năm trước từ student_enrollments
        prev_enr = supabase.table("student_enrollments").select("student_id, final_result") \
            .in_("student_id", student_ids) \
            .eq("academic_year_id", prev_year_id).execute().data

        updated = 0
        for pe in prev_enr:
            if pe.get('final_result') and pe['final_result'] != 'Chưa có kết quả':
                supabase.table("student_enrollments").update({
                    "final_result": pe['final_result']
                }).eq("student_id", pe['student_id']) \
                  .eq("class_id", class_id).execute()
                updated += 1

        return redirect(url_for('danh_sach_hs', class_id=class_id,
                                msg=f'Đã lấy kết quả năm trước cho {updated} học sinh!'))
    except Exception as e:
        return redirect(url_for('danh_sach_hs', class_id=class_id,
                                error_msg=f'Lỗi: {str(e)}'))
# ======================
# DANH SÁCH HỌC SINH
# ======================
@app.route('/teacher/danh-sach-hs', methods=['GET', 'POST'])
def danh_sach_hs():
    if 'user_id' not in session or session.get('role_name') != 'Teacher':
        return redirect(url_for('login'))

    teacher_id, semester_id, grades, classes = get_teacher_homeroom(session['user_id'])
    msg = error_msg = None

    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'update':
            try:
                student_id = int(request.form['student_id'])
                supabase.table("students").update({
                    "cho_o_hien_tai":      request.form.get('cho_o_hien_tai', '').strip() or None,
                    "sdt_ph":              request.form.get('sdt_ph', '').strip() or None,
                    "chuc_vu_doan_doi":    request.form.get('chuc_vu_doan_doi', '').strip() or None,
                    "chuc_vu_to_chuc_lop": request.form.get('chuc_vu_to_chuc_lop', '').strip() or None,
                }).eq("student_id", student_id).execute()
                msg = 'Cập nhật học sinh thành công!'
            except Exception as e:
                error_msg = f'Lỗi: {str(e)}'

    selected_grade = request.args.get('grade_id', type=int)
    selected_class = request.args.get('class_id', type=int)
    filtered_classes = [c for c in classes if not selected_grade or c['grade_id'] == selected_grade]

    students = []
    if selected_class:
        enr = supabase.table("student_enrollments") \
            .select("students(*)") \
            .eq("class_id", selected_class).execute().data
        students = [e['students'] for e in enr if e.get('students')]
        students.sort(key=lambda s: s.get('ho_ten', ''))

    return render_template('teacher/danh_sach_hs.html',
        grades=grades, classes=filtered_classes,
        selected_grade=selected_grade, selected_class=selected_class,
        students=students, msg=msg, error_msg=error_msg)
# ======================
# RUN APP
# ======================
if __name__ == '__main__':
    app.run(debug=True)