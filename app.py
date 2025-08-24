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
    if 'logged_in' not in session:
        return redirect(url_for('login'))
    # Dữ liệu mẫu, thực tế lấy từ service
    scores = [
        {'name': 'Toán rời rạc', 'mid': 8, 'final': 9, 'total': 8.5},
        {'name': 'Lập trình Python', 'mid': 9, 'final': 10, 'total': 9.5},
    ]
    return render_template('scores.html', scores=scores)


@app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html'), 404

@app.route('/')
def index():
    """Trang chủ"""
    global ictu_service
    
    # Nếu chưa có service, khởi tạo và thử restore session
    if not ictu_service:
        ictu_service = ICTUService()
        
        # Nếu service đã restore session thành công, cập nhật Flask session
        if ictu_service.is_logged_in and ictu_service.last_username:
            session['logged_in'] = True
            session['user_info'] = {
                'name': 'Auto-restored',  # Tên sẽ được cập nhật khi gọi API
                'studentId': 'Đang tải...',
                'studentDuration': 'Đang tải...'
            }
    
    if 'logged_in' not in session:
        return redirect(url_for('login'))
    return render_template('index.html', user_info=session.get('user_info'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Trang đăng nhập"""
    if request.method == 'GET':
        return render_template('login.html')
    try:
        data = request.get_json() if request.is_json else request.form
        username = data.get('username')
        password = data.get('password')
        if not username or not password:
            return jsonify({"error": True, "message": "Vui lòng nhập đầy đủ thông tin"})
        # Đăng nhập thật qua ICTUService
        global ictu_service
        ictu_service = ICTUService()
        result = ictu_service.login(username, password)
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
            return jsonify(result)
        else:
            return jsonify(result)
    except Exception as e:
        return jsonify({"error": True, "message": f"Lỗi server: {str(e)}"})

@app.route('/logout')
def logout():
    """Đăng xuất"""
    global ictu_service
    
    # Gọi logout trên service để xóa session file
    if ictu_service:
        ictu_service.logout()
        
    session.clear()
    ictu_service = None
    return redirect(url_for('login'))

@app.route('/api/lichthi')
def api_lichthi():
    """API lấy lịch thi từ ICTUService"""
    global ictu_service
    if not ictu_service:
        ictu_service = ICTUService()
    try:
        result = ictu_service.get_exam_schedule()
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
        return jsonify({"error": True, "message": f"Lỗi server: {str(e)}"})

@app.route('/api/dangkihoc')
def api_dangkihoc():
    """API lấy thông tin đăng ký học (mẫu)"""
    if 'logged_in' not in session:
        return jsonify({"error": True, "message": "Chưa đăng nhập"})
    data = {
        "error": False,
        "registered_courses": [
            {"subject": "Lập trình Python", "credits": 3, "status": "Đã đăng ký"},
            {"subject": "Toán rời rạc", "credits": 2, "status": "Đã đăng ký"}
        ]
    }
    return jsonify(data)

@app.route('/api/scores')
def api_scores():
    """API lấy điểm số từ ICTUService, chuẩn hóa dữ liệu trả về cho frontend"""
    global ictu_service
    if not ictu_service:
        ictu_service = ICTUService()
    try:
        result = ictu_service.get_scores()
        # Chuẩn hóa diemSoData
        diemSoData = []
        for i, row in enumerate(result.get('diemSoData', [])):
            diemSoData.append({
                'maHP': row.get('maHP', ''),
                'tenHP': row.get('tenHP', ''),
                'soTC': row.get('soTC', ''),
                'CC': row.get('chuyenCan', ''),
                'THI': row.get('thi', ''),
                'KTHP': row.get('tongKet', ''),
                'diemChu': row.get('diemChu', ''),
                'danhGia': row.get('danhGia', '')
            })
        # Chuẩn hóa tongKetData
        tongKetData = result.get('tongKetData', [])
        return jsonify({
            'error': False,
            'message': 'Success',
            'diemSoData': diemSoData,
            'tongKetData': tongKetData
        })
    except Exception as e:
        return jsonify({"error": True, "message": f"Lỗi server: {str(e)}"})

@app.route('/api/search')
def api_search():
    """API tìm kiếm lịch học (mẫu)"""
    if 'logged_in' not in session:
        return jsonify({"error": True, "message": "Chưa đăng nhập"})
    keyword = request.args.get('keyword', '').lower()
    # Dữ liệu mẫu
    all_schedules = [
        {"subject": "Lập trình Python", "date": "15/08/2025", "room": "A101", "time": "7:00-9:30"},
        {"subject": "Toán rời rạc", "date": "16/08/2025", "room": "B202", "time": "9:45-12:00"},
        {"subject": "Cơ sở dữ liệu", "date": "17/08/2025", "room": "C303", "time": "13:00-15:30"}
    ]
    filtered = [item for item in all_schedules if keyword in item["subject"].lower()]
    return jsonify({"error": False, "results": filtered})

@app.route('/api/timetable')
def api_timetable():
    """API lấy thời khóa biểu từ ICTUService, chuẩn hóa dữ liệu trả về cho frontend"""
    global ictu_service
    if not ictu_service:
        ictu_service = ICTUService()
    try:
        semester = request.args.get('semester')
        academic_year = request.args.get('academic_year')
        week = request.args.get('week')

        # Ưu tiên lấy từ Excel, nếu lỗi thì fallback sang HTML
        result = ictu_service.get_student_timetable_excel(semester=semester, academic_year=academic_year, week=week)
        if result.get('error'):
            print(f"[WARNING] Failed to get timetable from Excel: {result.get('message')}. Trying HTML...")
            result = ictu_service.get_student_timetable(semester=semester, academic_year=academic_year, week=week)

        if result.get('error'):
            return jsonify(result)

        # Cập nhật thông tin ngành vào session nếu có
        if 'major' in result and session.get('user_info'):
            session['user_info']['major'] = result['major']
            session.modified = True # Đánh dấu session đã thay đổi để lưu lại
            print(f"[DEBUG] Session major updated from timetable API: {session['user_info']['major']}")

        # Dữ liệu đã được chuẩn hóa trong ICTUService, chỉ cần trả về
        return jsonify({
            'error': False,
            'timetable': result.get('timetableData', []),
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
    global ictu_service
    if 'logged_in' not in session or not ictu_service:
        return jsonify({"error": True, "message": "Chưa đăng nhập"})
    try:
        result = ictu_service.get_timetable_options()
        return jsonify(result)
    except Exception as e:
        print(f"[ERROR] API Timetable Options exception: {str(e)}")
        return jsonify({"error": True, "message": f"Lỗi server khi lấy tùy chọn thời khóa biểu: {str(e)}"})

# Routes cho các trang
@app.route('/api/session_status')
def api_session_status():
    """API kiểm tra trạng thái session (mẫu)"""
    if 'logged_in' in session:
        return jsonify({
            "logged_in": True,
            "session_saved": True,
            "last_username": session.get('user_info', {}).get('name', None),
            "session_url_base": "http://localhost:5000/",
            "message": "Session active"
        })
    else:
        return jsonify({
            "logged_in": False,
            "session_saved": False,
            "message": "Not logged in"
        })

@app.route('/thoikhoabieu')
def thoikhoabieu():
    """Trang thời khóa biểu"""
    if 'logged_in' not in session:
        return redirect(url_for('login'))
    return render_template('thoikhoabieu.html', user_info=session.get('user_info'))

@app.route('/debug-help')
def debug_help():
    """Trang hướng dẫn debug"""
    return render_template('debug_help.html')

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
