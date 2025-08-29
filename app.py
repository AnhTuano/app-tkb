
import os
import re
from flask import Flask, render_template, request, jsonify, session, redirect, url_for

from ictu_service import ICTUService

# Biến toàn cục cho ICTUService
ictu_service = None

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  # Thay đổi thành secret key của bạn




@app.route('/scores')
def scores():
    print('[DEBUG] /scores route called')
    return jsonify({'error': True, 'message': 'Not implemented'}), 501


@app.errorhandler(404)
def page_not_found(e):
    return jsonify({'error': 'Not found'}), 404

@app.route('/')
def index():
    """Trang chủ"""
    global ictu_service
    print('[DEBUG] / route called')
    # Nếu chưa có service, khởi tạo và thử restore session
    if not ictu_service:
        print('[DEBUG] Khởi tạo ICTUService')
        ictu_service = ICTUService()
        # Nếu service đã restore session thành công, cập nhật Flask session
        if ictu_service.is_logged_in and ictu_service.last_username:
            print('[DEBUG] Đã restore session thành công')
            session['logged_in'] = True
            session['user_info'] = {
                'name': 'Auto-restored',  # Tên sẽ được cập nhật khi gọi API
                'studentId': 'Đang tải...',
                'studentDuration': 'Đang tải...'
            }
    if 'logged_in' not in session:
        print('[DEBUG] Chưa đăng nhập, chuyển hướng login')
        return redirect(url_for('login'))
    print(f"[DEBUG] user_info: {session.get('user_info')}")
    return jsonify({'user_info': session.get('user_info')})

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Trang đăng nhập"""
    print(f"[DEBUG] /login route called, method={request.method}")
    if request.method == 'GET':
        return jsonify({'message': 'Login page'})
    try:
        data = request.get_json() if request.is_json else request.form
        print(f"[DEBUG] Login data: {data}")
        username = data.get('username')
        password = data.get('password')
        if not username or not password:
            print('[DEBUG] Thiếu username hoặc password')
            return jsonify({"error": True, "message": "Vui lòng nhập đầy đủ thông tin"})
        # Đăng nhập thật qua ICTUService
        global ictu_service
        ictu_service = ICTUService()
        result = ictu_service.login(username, password)
        print(f"[DEBUG] Kết quả login: {result}")
        if not result["error"]:
            session['logged_in'] = True
            session['user_info'] = {
                'name': result.get('name'),
                'studentId': result.get('studentId'),
                'studentDuration': result.get('studentDuration'),
                'email': result.get('email', username + '@ictu.edu.vn'),
                'major': result.get('major', 'Chưa cập nhật') # Lấy thông tin ngành
            }
            # Luôn trả về avatar mặc định
            result['avatar_url'] = 'https://cdn-icons-png.flaticon.com/512/3135/3135715.png'
            print(f"[DEBUG] Đăng nhập thành công, session user_info: {session['user_info']}")
            return jsonify(result)
        else:
            print('[DEBUG] Đăng nhập thất bại')
            return jsonify(result)
    except Exception as e:
        print(f"[ERROR] Lỗi server khi login: {str(e)}")
        return jsonify({"error": True, "message": f"Lỗi server: {str(e)}"})

@app.route('/logout')
def logout():
    """Đăng xuất"""
    print('[DEBUG] /logout route called')
    global ictu_service
    # Gọi logout trên service để xóa session file
    if ictu_service:
        print('[DEBUG] Gọi ictu_service.logout()')
        ictu_service.logout()
    session.clear()
    ictu_service = None
    print('[DEBUG] Đã logout, chuyển hướng login')
    return redirect(url_for('login'))

@app.route('/api/lichthi')
def api_lichthi():
    """API lấy lịch thi từ ICTUService"""
    print('[DEBUG] /api/lichthi route called')
    global ictu_service
    if not ictu_service:
        print('[DEBUG] Khởi tạo ICTUService')
        ictu_service = ICTUService()
    try:
        result = ictu_service.get_exam_schedule()
        print(f"[DEBUG] Kết quả get_exam_schedule: {result}")
        # Chuẩn hóa dữ liệu trả về cho frontend
        lichthiData = []
        for row in result.get('lichthiData', []):
            lichthiData.append({
                'maHP': row.get('maHP', ''),
                'tenHP': row.get('tenHP', ''),
                'soTC': row.get('soTC', ''),
                'ngayThi': row.get('ngayThi', ''),
                'caThi': row.get('caThi', ''),
                'hinhThucThi': row.get('hinhThucThi', ''),
                'soBaoDanh': row.get('soBaoDanh', ''),
                'phongThi': row.get('phongThi', ''),
                'ghiChu': row.get('ghiChu', '')
            })
        return jsonify({
            'error': False,
            'lichthiData': lichthiData
        })
    except Exception as e:
        print(f"[ERROR] Lỗi server khi lấy lịch thi: {str(e)}")
        return jsonify({"error": True, "message": f"Lỗi server: {str(e)}"})

@app.route('/api/dangkihoc')
def api_dangkihoc():
    print('[DEBUG] /api/dangkihoc route called')
    return jsonify({'error': True, 'message': 'Not implemented'}), 501

@app.route('/api/scores')
def api_scores():
    """API lấy điểm số từ ICTUService, chuẩn hóa dữ liệu trả về cho frontend"""
    print('[DEBUG] /api/scores route called')
    global ictu_service
    if not ictu_service:
        print('[DEBUG] Khởi tạo ICTUService')
        ictu_service = ICTUService()
    try:
        result = ictu_service.get_scores()
        diemSoData = result.get('diemSoData', [])
        tongKetData = result.get('tongKetData', [])
        print(f"[DEBUG] Số môn học: {len(diemSoData)}")
        for i, row in enumerate(diemSoData):
            print(f"[DEBUG] {i+1}. {row.get('maHP','')} - {row.get('tenHP','')} | TC: {row.get('soTC','')} | TK: {row.get('tongKet','')} | ĐG: {row.get('danhGia','')}")
        print(f"[DEBUG] Tổng kết học kỳ: {tongKetData}")
        # Chuẩn hóa diemSoData
        diemSoData_out = []
        for i, row in enumerate(diemSoData):
            diemSoData_out.append({
                'maHP': row.get('maHP', ''),
                'tenHP': row.get('tenHP', ''),
                'soTC': row.get('soTC', ''),
                'CC': row.get('chuyenCan', ''),
                'THI': row.get('thi', ''),
                'KTHP': row.get('tongKet', ''),
                'diemChu': row.get('diemChu', ''),
                'danhGia': row.get('danhGia', '')
            })
        return jsonify({
            'error': False,
            'message': 'Success',
            'diemSoData': diemSoData_out,
            'tongKetData': tongKetData
        })
    except Exception as e:
        print(f"[ERROR] Lỗi server khi lấy điểm số: {str(e)}")
        return jsonify({"error": True, "message": f"Lỗi server: {str(e)}"})

@app.route('/api/search')
def api_search():
    print('[DEBUG] /api/search route called')
    return jsonify({'error': True, 'message': 'Not implemented'}), 501

@app.route('/api/timetable')
def api_timetable():
    """API lấy thời khóa biểu từ ICTUService, chuẩn hóa dữ liệu trả về cho frontend"""
    print('[DEBUG] /api/timetable route called')
    global ictu_service
    if not ictu_service:
        print('[DEBUG] Khởi tạo ICTUService')
        ictu_service = ICTUService()
    try:
        semester = request.args.get('semester')
        academic_year = request.args.get('academic_year')
        week = request.args.get('week')
        print(f"[DEBUG] Params: semester={semester}, academic_year={academic_year}, week={week}")
        # Ưu tiên lấy từ Excel, nếu lỗi thì fallback sang HTML
        result = ictu_service.get_student_timetable_excel(semester=semester, academic_year=academic_year, week=week)
        if result.get('error'):
            print(f"[WARNING] Failed to get timetable from Excel: {result.get('message')}. Trying HTML...")
            result = ictu_service.get_student_timetable(semester=semester, academic_year=academic_year, week=week)
        if result.get('error'):
            print(f"[ERROR] Lỗi khi lấy thời khóa biểu: {result}")
            return jsonify(result)
        timetable = result.get('timetableData', [])
        print(f"[DEBUG] Số dòng thời khoá biểu: {len(timetable)}")
        for i, row in enumerate(timetable):
            print(f"[DEBUG] {i+1}. {row.get('tenHP','')} | {row.get('phong','')} | {row.get('ca','')} | {row.get('thu','')} {row.get('tuan','')}")
        # Cập nhật thông tin ngành vào session nếu có
        if 'major' in result and session.get('user_info'):
            session['user_info']['major'] = result['major']
            session.modified = True # Đánh dấu session đã thay đổi để lưu lại
            print(f"[DEBUG] Session major updated from timetable API: {session['user_info']['major']}")
        return jsonify({
            'error': False,
            'timetable': timetable,
            'source': result.get('source', 'unknown'),
            'totalRows': result.get('totalRows', 0),
            'major': result.get('major', 'Chưa cập nhật') # Thêm major vào response cho frontend
        })
    except Exception as e:
        print(f"[ERROR] API Timetable exception: {str(e)}")
        return jsonify({"error": True, "message": f"Lỗi server khi lấy thời khóa biểu: {str(e)}"})

@app.route('/api/timetable_options')
def api_timetable_options():
    """API lấy các tùy chọn lọc thời khóa biểu (học kỳ, năm học, tuần)"""
    print('[DEBUG] /api/timetable_options route called')
    global ictu_service
    if 'logged_in' not in session or not ictu_service:
        print('[DEBUG] Chưa đăng nhập hoặc chưa có ictu_service')
        return jsonify({"error": True, "message": "Chưa đăng nhập"})
    try:
        result = ictu_service.get_timetable_options()
        print(f"[DEBUG] Kết quả get_timetable_options: {result}")
        return jsonify(result)
    except Exception as e:
        print(f"[ERROR] API Timetable Options exception: {str(e)}")
        return jsonify({"error": True, "message": f"Lỗi server khi lấy tùy chọn thời khóa biểu: {str(e)}"})

# Routes cho các trang
@app.route('/api/session_status')
def api_session_status():
    print('[DEBUG] /api/session_status route called')
    return jsonify({'error': True, 'message': 'Not implemented'}), 501

@app.route('/thoikhoabieu')
def thoikhoabieu():
    print('[DEBUG] /thoikhoabieu route called')
    return jsonify({'error': True, 'message': 'Not implemented'}), 501

@app.route('/debug-help')
def debug_help():
    print('[DEBUG] /debug-help route called')
    return jsonify({'error': True, 'message': 'Not implemented'}), 501

# Chạy bằng waitress nếu chạy trực tiếp
if __name__ == '__main__':
    from waitress import serve
    print('Serving with waitress on http://0.0.0.0:5000 ...')
    serve(app, host='0.0.0.0', port=5000)
