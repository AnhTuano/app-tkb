import requests
import hashlib
from bs4 import BeautifulSoup
import json
import traceback
import re
import pickle
import os
from datetime import datetime, timedelta # Import datetime and timedelta

class ICTUService:
    def get_student_timetable(self, semester=None, academic_year=None, week=None):
        """Lấy thời khóa biểu sinh viên từ trang HTML (fallback khi Excel lỗi)"""
        try:
            auth_check = self._ensure_logged_in()
            if auth_check:
                return auth_check

            print(f"[DEBUG] Lấy thời khóa biểu từ HTML...")
            timetable_url = f"{self.base_url}/Reports/Form/StudentTimeTable.aspx"
            response = self.session.get(timetable_url, timeout=30, allow_redirects=True)
            if response.status_code != 200:
                return self._handle_error(f"Không thể truy cập trang thời khóa biểu (HTTP {response.status_code})", response.status_code)

            soup = BeautifulSoup(response.text, 'html.parser')
            table = soup.find('table', {'id': 'grdStudentTimeTable'})
            if not table:
                return self._handle_error("Không tìm thấy bảng thời khóa biểu trên trang HTML", 404)

            headers = []
            timetable_data = []
            for i, row in enumerate(table.find_all('tr')):
                cols = row.find_all(['td', 'th'])
                col_text = [col.get_text(strip=True) for col in cols]
                if i == 0:
                    headers = col_text
                else:
                    if len(col_text) == len(headers):
                        timetable_data.append(dict(zip(headers, col_text)))

            # Chuẩn hóa dữ liệu trả về giống Excel
            mapped_data = []
            for item in timetable_data:
                mapped_item = {
                    "stt": item.get('STT', ''),
                    "lopHocPhan": item.get('Lớp học phần', ''),
                    "maHP": item.get('Mã HP', ''),
                    "tenHP": item.get('Tên HP', ''),
                    "soTC": item.get('Số TC', ''),
                    "thu": item.get('Thứ', ''),
                    "tiet": item.get('Tiết học', ''),
                    "phong": item.get('Phòng', ''),
                    "giangVien": item.get('Giảng viên', ''),
                    "meetLink": '',
                    "siSo": item.get('Sĩ số', ''),
                    "soDK": item.get('Số ĐK', ''),
                    "hocPhi": item.get('Học phí', ''),
                    "ghiChu": item.get('Ghi chú', ''),
                    "from_date": '',
                    "to_date": '',
                    "week_number": '',
                    "lesson_type": ''
                }
                mapped_data.append(mapped_item)

            return {
                "error": False,
                "timetableData": mapped_data,
                "originalColumns": headers,
                "source": "html",
                "totalRows": len(mapped_data),
                "major": ""
            }
        except Exception as e:
            print(f"[ERROR] Timetable HTML exception: {str(e)}")
            print(f"[ERROR] Traceback: {traceback.format_exc()}")
            return self._handle_error("Error fetching timetable HTML fallback", 500)
    def __init__(self):
        self.base_url = "http://220.231.119.171/kcntt"
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        })
        self.is_logged_in = False
        self.session_url_base = None
        self.last_username = None
        self.last_password = None

    # Bỏ hoàn toàn lưu session ra file để tránh xung đột nhiều người dùng
    def _save_session(self, username, password):
        pass
    
    # Bỏ hoàn toàn đọc session từ file để tránh xung đột nhiều người dùng
    def _load_session(self):
        return False
    
    def _validate_session(self):
        """Kiểm tra session có còn hợp lệ không"""
        try:
            if not self.is_logged_in:
                return False
                
            # Thử truy cập trang cần đăng nhập
            test_url = f"{self.base_url}/StudyRegister/StudyRegister.aspx"
            response = self.session.get(test_url, timeout=10)
            
            # Nếu bị redirect về login thì session hết hạn
            if 'login.aspx' in response.url.lower():
                self.is_logged_in = False
                return False
                
            return response.status_code == 200
            
        except Exception as e:
            print(f"[ERROR] Session validation failed: {e}")
            return False
    
    def _auto_relogin(self):
        """Tự động đăng nhập lại với thông tin đã lưu"""
        try:
            if not self.last_username or not self.last_password:
                print(f"[DEBUG] No saved credentials for auto-relogin")
                return False
                
            print(f"[DEBUG] Auto-relogin for user: {self.last_username}")
            
            # Reset session
            self.session = requests.Session()
            self.session.headers.update({
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
            })
            
            # Đăng nhập lại
            result = self.login(self.last_username, self.last_password)
            
            return not result.get('error', True)
            
        except Exception as e:
            print(f"[ERROR] Auto-relogin failed: {e}")
            return False
    
    def logout(self):
        """Đăng xuất (không xóa file session nữa)"""
        self.is_logged_in = False
        self.session_url_base = None
        self.last_username = None
        self.last_password = None
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        })
        print(f"[DEBUG] Logged out (no file session)")
    
    def _ensure_logged_in(self):
        """Đảm bảo đã đăng nhập, tự động relogin nếu cần"""
        if not self.is_logged_in:
            return self._handle_error("Chưa đăng nhập vào hệ thống", 401)
            
        # Kiểm tra session có còn hợp lệ không
        if not self._validate_session():
            print(f"[DEBUG] Session expired, attempting auto-relogin...")
            
            if self._auto_relogin():
                print(f"[DEBUG] Auto-relogin successful")
                return None  # Success
            else:
                print(f"[DEBUG] Auto-relogin failed")
                return self._handle_error("Phiên đăng nhập đã hết hạn, vui lòng đăng nhập lại", 401)
        
        return None  # Success
        
    def _handle_error(self, message: str, status_code: int = 500):
        """Helper function to create error response"""
        return {
            "error": True,
            "message": message,
            "status": status_code
        }
        
    def login(self, username, password):
        """Đăng nhập vào hệ thống"""
        try:
            print(f"[DEBUG] Bắt đầu đăng nhập cho user: {username}")
            
            # Lấy trang login để có session và viewstate
            login_url = f"{self.base_url}/login.aspx"
            print(f"[DEBUG] Truy cập trang login: {login_url}")
            
            session_response = self.session.get(login_url, timeout=30)
            print(f"[DEBUG] Session response status: {session_response.status_code}")
            
            if session_response.status_code != 200:
                return self._handle_error(f"Không thể truy cập trang đăng nhập (HTTP {session_response.status_code})", 500)
            
            soup = BeautifulSoup(session_response.text, 'html.parser')
            
            # Lấy tất cả form elements giống như Next.js
            form = soup.find('form', {'id': 'Form1'})
            if not form:
                return self._handle_error("Không tìm thấy form đăng nhập", 404)
            
            # Tạo body data từ tất cả form elements
            post_data = {}
            
            # Lấy tất cả input, select, textarea elements
            form_elements = form.find_all(['input', 'select', 'textarea'])
            
            for element in form_elements:
                name = element.get('name')
                value = element.get('value', '')
                
                if name:
                    if name == 'txtUserName':
                        value = username
                    elif name == 'txtPassword':
                        value = hashlib.md5(password.encode()).hexdigest()
                    
                    if value:
                        post_data[name] = value
            
            print(f"[DEBUG] Form data keys: {list(post_data.keys())}")
            print(f"[DEBUG] Password MD5: {post_data.get('txtPassword', '')[:10]}...")
            
            # Thực hiện đăng nhập
            login_response = self.session.post(session_response.url, data=post_data, timeout=30)
            print(f"[DEBUG] Login response status: {login_response.status_code}")
            print(f"[DEBUG] Login response URL: {login_response.url}")
            
            # Kiểm tra lỗi đăng nhập
            login_soup = BeautifulSoup(login_response.text, 'html.parser')
            error_info = login_soup.find('span', {'id': 'lblErrorInfo'})
            
            if error_info and error_info.get_text(strip=True):
                return self._handle_error(error_info.get_text(strip=True), 401)
            
            # Lấy thông tin từ trang Home
            home_url = f"{self.base_url}/Home.aspx"
            home_response = self.session.get(home_url, timeout=30)
            print(f"[DEBUG] Home response status: {home_response.status_code}")
            
            if home_response.status_code != 200:
                return self._handle_error(f"Không thể truy cập trang chủ (HTTP {home_response.status_code})", 500)
            
            home_soup = BeautifulSoup(home_response.text, 'html.parser')
            
            # Lấy thông tin sinh viên giống như Next.js
            student_info = home_soup.find('span', {'id': 'PageHeader1_lblUserFullName'})
            
            if not student_info or not student_info.get_text(strip=True):
                return self._handle_error("Không thể lấy thông tin sinh viên", 404)
            
            # Parse thông tin sinh viên từ format "Tên (MSSV)"
            student_text = student_info.get_text(strip=True)
            start_index = student_text.find("(")
            end_index = student_text.find(")")
            
            if start_index != -1 and end_index != -1:
                name = student_text[:start_index].strip()
                student_id = student_text[start_index + 1:end_index]
            else:
                name = student_text
                student_id = "N/A"
            
            # Lấy thông tin ngành (Major) - thử nhiều cách để tăng độ chính xác
            major = "Chưa cập nhật"
            # 1. Theo id phổ biến
            major_element = home_soup.find('span', {'id': 'lblNganh'})
            if major_element and major_element.get_text(strip=True):
                major = major_element.get_text(strip=True)
            else:
                # 2. Theo text chứa "Ngành:"
                major_element = home_soup.find(lambda tag: tag.name in ['span','div','td'] and tag.get_text(strip=True).startswith('Ngành:'))
                if major_element:
                    major = major_element.get_text(strip=True).replace('Ngành:', '').strip()
                else:
                    # 3. Theo class phổ biến
                    major_element = home_soup.find(class_=re.compile(r'nganh|major', re.I))
                    if major_element and major_element.get_text(strip=True):
                        major = major_element.get_text(strip=True)
                    else:
                        # 4. Theo thẻ td/th có text "Ngành" bên trái
                        label_cell = home_soup.find(lambda tag: tag.name in ['td','th'] and 'ngành' in tag.get_text(strip=True).lower())
                        if label_cell and label_cell.find_next_sibling(['td','th']):
                            major = label_cell.find_next_sibling(['td','th']).get_text(strip=True)
                        else:
                            # 5. Theo bất kỳ span/div nào có text chứa "ngành"
                            generic = home_soup.find(lambda tag: tag.name in ['span','div'] and 'ngành' in tag.get_text(strip=True).lower())
                            if generic:
                                # Lấy phần sau dấu : nếu có
                                txt = generic.get_text(strip=True)
                                if ':' in txt:
                                    major = txt.split(':',1)[-1].strip()
                                else:
                                    major = txt.strip()

            # Nếu vẫn chưa lấy được ngành, thử lấy từ trang TKB (StudentTimeTable.aspx)
            if (not major or major == "Chưa cập nhật"):
                try:
                    timetable_url = f"{self.base_url}/Reports/Form/StudentTimeTable.aspx"
                    timetable_response = self.session.get(timetable_url, timeout=15)
                    if timetable_response.status_code == 200:
                        timetable_soup = BeautifulSoup(timetable_response.text, 'html.parser')
                        tkb_major = ""
                        major_tag = timetable_soup.find(lambda tag: tag.name in ['span','div'] and 'ngành' in tag.get_text(strip=True).lower())
                        if major_tag:
                            txt = major_tag.get_text(strip=True)
                            if ':' in txt:
                                tkb_major = txt.split(':',1)[-1].strip()
                            else:
                                tkb_major = txt.strip()
                        if not tkb_major:
                            label_cell = timetable_soup.find(lambda tag: tag.name in ['td','th'] and 'ngành' in tag.get_text(strip=True).lower())
                            if label_cell and label_cell.find_next_sibling(['td','th']):
                                tkb_major = label_cell.find_next_sibling(['td','th']).get_text(strip=True)
                        if tkb_major:
                            # Loại bỏ mã sinh viên, tên sinh viên nếu có, chỉ lấy tên ngành
                            # Ví dụ: "DTC245200672 - Nguyễn Anh Tuấn - Chuyên ngành Công nghệ thông tin" => "Công nghệ thông tin"
                            if '-' in tkb_major:
                                major = tkb_major.split('-')[-1].strip()
                            else:
                                major = tkb_major.strip()
                            # Bỏ tiền tố "Chuyên ngành" nếu có
                            if major.lower().startswith('chuyên ngành'):
                                major = major[len('chuyên ngành'):].strip()
                            print(f"[DEBUG] Extracted Major from TKB: {major}")
                except Exception as e:
                    print(f"[ERROR] Try get major from timetable failed: {e}")
            print(f"[DEBUG] Extracted Major: {major}")

            # Lấy thông tin thời gian học
            study_register_url = f"{self.base_url}/StudyRegister/StudyRegister.aspx"
            study_response = self.session.get(study_register_url, timeout=30)
            
            student_duration = "N/A"
            if study_response.status_code == 200:
                study_soup = BeautifulSoup(study_response.text, 'html.parser')
                duration_element = study_soup.find('span', {'id': 'lblDuration'})
                if duration_element:
                    student_duration = duration_element.get_text(strip=True)
            
            print(f"[DEBUG] Đăng nhập thành công - Name: {name}, ID: {student_id}")
            
            self.is_logged_in = True
            
            # Lưu thông tin đăng nhập để duy trì session
            self.last_username = username
            self.last_password = password
            self._save_session(username, password)  # Sẽ không làm gì cả
            
            return {
                "error": False,
                "message": "Đăng nhập thành công!",
                "name": name,
                "studentId": student_id,
                "studentDuration": student_duration,
                "major": major, # Thêm thông tin ngành vào kết quả trả về
                "email": f"{username}@ictu.edu.vn",
                "token": "session_maintained"  # Python session tự động maintain
            }
                
        except requests.exceptions.Timeout:
            return self._handle_error("Kết nối timeout - Vui lòng thử lại", 408)
        except requests.exceptions.ConnectionError:
            return self._handle_error("Lỗi kết nối - Kiểm tra mạng internet", 503)
        except Exception as e:
            print(f"[ERROR] Login exception: {str(e)}")
            print(f"[ERROR] Traceback: {traceback.format_exc()}")
            return self._handle_error(f"Lỗi đăng nhập: {str(e)}", 500)
    
    def get_exam_schedule(self):
        """Lấy lịch thi"""
        try:
            # Kiểm tra và tự động relogin nếu cần
            auth_check = self._ensure_logged_in()
            if auth_check:  # Có lỗi
                return auth_check
                
            print(f"[DEBUG] Lấy lịch thi...")
            
            exam_url = f"{self.base_url}/StudentViewExamList.aspx"
            response = self.session.get(exam_url, timeout=30)
            
            print(f"[DEBUG] Exam schedule response status: {response.status_code}")
            
            if response.status_code != 200:
                return self._handle_error(f"Không thể truy cập trang lịch thi (HTTP {response.status_code})", response.status_code)
            
            # Kiểm tra có bị redirect về login không
            if 'login.aspx' in response.url.lower():
                return self._handle_error("Phiên đăng nhập đã hết hạn", 401)
            
            soup = BeautifulSoup(response.text, 'html.parser')
            print(f"[DEBUG] Page title: {soup.title.text if soup.title else 'No title'}")
            
            # Tìm bảng lịch thi với ID chính xác như Next.js
            table = soup.find('table', {'id': 'tblCourseList'})
            
            if not table:
                return self._handle_error("Table 'tblCourseList' not found", 404)
            
            lichthiData = []
            rows = table.find_all('tr')
            print(f"[DEBUG] Found {len(rows)} rows in table")
            
            # Bỏ header row (index 0) và xử lý từng dòng từ index 1
            for i in range(1, len(rows)):
                row = rows[i]
                cells = row.find_all('td')
                
                if len(cells) >= 10:  # Đảm bảo có đủ cột
                    # Lấy text content từ mỗi cell và trim whitespace
                    row_data = [cell.get_text(strip=True) for cell in cells]
                    
                    # Destructure array giống như Next.js
                    stt, maHP, tenHP, soTC, ngayThi, caThi, hinhThucThi, soBaoDanh, phongThi, ghiChu = row_data[:10]
                    
                    # Kiểm tra nếu stt không rỗng (giống logic Next.js)
                    if stt != "":
                        exam_item = {
                            "stt": stt,
                            "maHP": maHP,
                            "tenHP": tenHP,
                            "soTC": soTC,
                            "ngayThi": ngayThi,
                            "caThi": caThi,
                            "hinhThucThi": hinhThucThi,
                            "soBaoDanh": soBaoDanh,
                            "phongThi": phongThi,
                            "ghiChu": ghiChu
                        }
                        lichthiData.append(exam_item)
                        print(f"[DEBUG] Added exam: {exam_item['tenHP']}")
            
            print(f"[DEBUG] Total exams found: {len(lichthiData)}")
            
            return {
                "error": False,
                "lichthiData": lichthiData
            }
            
        except requests.exceptions.Timeout:
            return self._handle_error("Kết nối timeout khi lấy lịch thi", 408)
        except requests.exceptions.ConnectionError:
            return self._handle_error("Lỗi kết nối khi lấy lịch thi", 503)
        except Exception as e:
            print(f"[ERROR] Exam schedule exception: {str(e)}")
            print(f"[ERROR] Traceback: {traceback.format_exc()}")
            return self._handle_error(f"Error fetching exam schedule data after login", 500)
    
    def get_study_registration(self):
        """Lấy thông tin đăng ký học"""
        try:
            # Kiểm tra và tự động relogin nếu cần
            auth_check = self._ensure_logged_in()
            if auth_check:  # Có lỗi
                return auth_check
                
            print(f"[DEBUG] Lấy thông tin đăng ký học...")
            
            register_url = f"{self.base_url}/StudyRegister/StudyRegister.aspx"
            response = self.session.get(register_url, timeout=30)
            
            print(f"[DEBUG] Study registration response status: {response.status_code}")
            
            if response.status_code != 200:
                return self._handle_error(f"Không thể truy cập trang đăng ký học (HTTP {response.status_code})", response.status_code)
            
            # Kiểm tra có bị redirect về login không
            if 'login.aspx' in response.url.lower():
                return self._handle_error("Phiên đăng nhập đã hết hạn", 401)
            
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Lấy thông tin thời gian học
            duration_element = soup.find('span', {'id': 'lblDuration'})
            student_duration = duration_element.get_text(strip=True) if duration_element else None
            print(f"[DEBUG] Student duration: {student_duration}")
            
            # Lấy dropdown khóa học
            course_select = soup.find('select', {'id': 'drpCourse'})
            
            if not course_select:
                return self._handle_error("No courses found", 404)
            
            # Lấy tất cả options và filter như Next.js
            options = course_select.find_all('option')
            print(f"[DEBUG] Found {len(options)} course options")
            
            courses = []
            for option in options:
                value = option.get('value')
                text = option.get_text(strip=True)
                
                # Filter option có value (giống Next.js logic)
                if value:
                    courses.append({
                        "value": value,
                        "text": text
                    })
                    print(f"[DEBUG] Added course: {text}")
            
            return {
                "error": False,
                "studentDuration": student_duration,
                "courses": courses
            }
            
        except requests.exceptions.Timeout:
            return self._handle_error("Kết nối timeout khi lấy thông tin đăng ký học", 408)
        except requests.exceptions.ConnectionError:
            return self._handle_error("Lỗi kết nối khi lấy thông tin đăng ký học", 503)
        except Exception as e:
            print(f"[ERROR] Study registration exception: {str(e)}")
            print(f"[ERROR] Traceback: {traceback.format_exc()}")
            return self._handle_error("Error fetching study registration data", 500)

    def get_scores(self):
        """Lấy điểm số"""
        try:
            # Kiểm tra và tự động relogin nếu cần
            auth_check = self._ensure_logged_in()
            if auth_check:  # Có lỗi
                return auth_check
                
            print(f"[DEBUG] Lấy thông tin điểm số...")
            
            scores_url = f"{self.base_url}/StudentMark.aspx"
            response = self.session.get(scores_url, timeout=30)
            
            print(f"[DEBUG] Scores response status: {response.status_code}")
            
            if response.status_code != 200:
                return self._handle_error(f"Không thể truy cập trang điểm (HTTP {response.status_code})", response.status_code)
            
            # Kiểm tra có bị redirect về login không
            if 'login.aspx' in response.url.lower():
                return self._handle_error("Phiên đăng nhập đã hết hạn", 401)
            
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Tìm tables như Next.js
            table_score_detail = (soup.find('table', {'id': 'tblMarkDetail'}) or 
                                 soup.find('table', {'id': 'tblStudentMark'}))
            table_score_sum = soup.find('table', {'id': 'tblSumMark'})
            
            print(f"[DEBUG] Score detail table found: {table_score_detail is not None}")
            print(f"[DEBUG] Score sum table found: {table_score_sum is not None}")
            
            if not table_score_detail or not table_score_sum:
                return self._handle_error("Table not found", 404)
            
            # Process detail table (giống Next.js)
            data_score_detail = []
            rows = table_score_detail.find_all('tr')
            print(f"[DEBUG] Detail table has {len(rows)} rows")
            
            # Bắt đầu từ row 2 như Next.js (skip 2 header rows)
            for i in range(2, len(rows)):
                row = rows[i]
                cells = row.find_all('td')
                
                if len(cells) < 14:
                    continue
                
                # Map theo exact columns như Next.js
                score_item = {
                    "stt": cells[0].get_text(strip=True),
                    "maHP": cells[1].get_text(strip=True),
                    "tenHP": cells[2].get_text(strip=True),
                    "soTC": cells[3].get_text(strip=True),
                    "danhGia": cells[8].get_text(strip=True),
                    "chuyenCan": cells[10].get_text(strip=True),
                    "thi": cells[11].get_text(strip=True),
                    "tongKet": cells[12].get_text(strip=True),
                    "diemChu": cells[13].get_text(strip=True)
                }
                data_score_detail.append(score_item)
                print(f"[DEBUG] Added score detail: {score_item['tenHP']}")
            
            # Process sum table (giống Next.js)
            data_score_sum = []
            rows_sum = table_score_sum.find_all('tr')
            print(f"[DEBUG] Sum table has {len(rows_sum)} rows")
            
            # Bắt đầu từ row 2 như Next.js
            for i in range(2, len(rows_sum)):
                row = rows_sum[i]
                cells = row.find_all('td')
                
                if len(cells) < 14:
                    continue
                
                # Map theo exact columns như Next.js
                sum_item = {
                    "namHoc": cells[0].get_text(strip=True),
                    "hocKy": cells[1].get_text(strip=True),
                    "TBTL10": cells[2].get_text(strip=True),
                    "TBTL4": cells[4].get_text(strip=True),
                    "TC": cells[6].get_text(strip=True),
                    "TBC10": cells[8].get_text(strip=True),
                    "TBC4": cells[10].get_text(strip=True)
                }
                data_score_sum.append(sum_item)
                print(f"[DEBUG] Added score sum: {sum_item['namHoc']} - {sum_item['hocKy']}")
            
            return {
                "error": False,
                "message": "Success",
                "diemSoData": data_score_detail,
                "tongKetData": data_score_sum
            }
            
        except requests.exceptions.Timeout:
            return self._handle_error("Kết nối timeout khi lấy điểm", 408)
        except requests.exceptions.ConnectionError:
            return self._handle_error("Lỗi kết nối khi lấy điểm", 503)
        except Exception as e:
            print(f"[ERROR] Scores exception: {str(e)}")
            print(f"[ERROR] Traceback: {traceback.format_exc()}")
            return self._handle_error("Error fetching scores data", 500)
    
    def search_schedule(self, keyword=""):
        """Tìm kiếm lịch học"""
        try:
            # Lấy lịch thi trước
            exam_result = self.get_exam_schedule()
            
            if exam_result["error"]:
                return exam_result
            
            lichthiData = exam_result["lichthiData"]
            
            # Nếu có keyword, lọc kết quả
            if keyword:
                filtered_data = []
                keyword_lower = keyword.lower()
                
                for item in lichthiData:
                    if (keyword_lower in item["tenHP"].lower() or 
                        keyword_lower in item["maHP"].lower() or
                        keyword_lower in item["ngayThi"].lower() or
                        keyword_lower in item["phongThi"].lower()):
                        filtered_data.append(item)
                
                return {
                    "error": False,
                    "message": f"Tìm thấy {len(filtered_data)} kết quả",
                    "lichthiData": filtered_data,
                    "keyword": keyword
                }
            else:
                return {
                    "error": False,
                    "message": f"Tất cả lịch thi ({len(lichthiData)} kết quả)",
                    "lichthiData": lichthiData
                }
                
        except Exception as e:
            return {"error": True, "message": f"Lỗi tìm kiếm lịch học: {str(e)}"}
    
    def get_student_timetable_excel(self, semester=None, academic_year=None, week=None):
        """Lấy thời khóa biểu sinh viên từ file Excel với tùy chọn học kỳ, năm học, tuần"""
        try:
            # Kiểm tra và tự động relogin nếu cần
            auth_check = self._ensure_logged_in()
            if auth_check:  # Có lỗi
                return auth_check
                
            print(f"[DEBUG] Lấy thời khóa biểu từ Excel...")
            
            # Truy cập trang thời khóa biểu
            timetable_url_basic = f"{self.base_url}/Reports/Form/StudentTimeTable.aspx"
            response = self.session.get(timetable_url_basic, timeout=30, allow_redirects=True)
            
            print(f"[DEBUG] Timetable page status: {response.status_code}")
            
            if response.status_code != 200:
                return self._handle_error(f"Không thể truy cập trang thời khóa biểu (HTTP {response.status_code})", response.status_code)
            
            # Parse HTML để lấy form data
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # Tìm form và nút Excel
            form = soup.find('form', {'id': 'Form1'})
            if not form:
                return self._handle_error("Không tìm thấy form trên trang", 404)
            
            excel_button = soup.find('input', {'id': 'btnView'})
            if not excel_button:
                # Fallback: tìm theo value chứa Excel
                excel_button = soup.find('input', lambda tag: tag.get('value') and 'excel' in tag.get('value', '').lower())
            
            if not excel_button:
                return self._handle_error("Không tìm thấy nút Xuất file Excel", 404)
            
            print(f"[DEBUG] Found Excel button: {excel_button.get('name')}")
            
            # Tạo form data từ tất cả input, select, textarea
            form_data = {}
            form_elements = form.find_all(['input', 'select', 'textarea'])
            
            for element in form_elements:
                name = element.get('name')
                if name:
                    value = element.get('value', '')
                    
                    # Set specific values if provided
                    if name == 'drpHocKy' and semester is not None: # Assuming drpHocKy for semester
                        value = semester
                    elif name == 'drpNamHoc' and academic_year is not None: # Assuming drpNamHoc for academic year
                        value = academic_year
                    elif name == 'drpTuan' and week is not None: # Assuming drpTuan for week
                        value = week

                    if element.name == 'input':
                        input_type = element.get('type', '').lower()
                        if input_type in ['checkbox', 'radio']:
                            if element.get('checked'):
                                form_data[name] = value
                        elif input_type == 'submit':
                            # Chỉ add submit button được click
                            if element.get('id') == 'btnView':
                                form_data[name] = value
                        else:
                            form_data[name] = value
                    elif element.name == 'select':
                        # For select, ensure the value is one of the options
                        options = element.find_all('option')
                        selected_option_value = value # Use the provided value or default
                        
                        # Check if the provided value exists in options
                        if selected_option_value not in [opt.get('value') for opt in options]:
                            # If not, try to find a default selected option or the first one
                            selected_option = element.find('option', selected=True)
                            if selected_option:
                                selected_option_value = selected_option.get('value', '')
                            elif options:
                                selected_option_value = options[0].get('value', '')
                        
                        form_data[name] = selected_option_value

                    elif element.name == 'textarea':
                        form_data[name] = element.get_text()
            
            print(f"[DEBUG] Form data prepared with {len(form_data)} fields")
            print(f"[DEBUG] Selected semester: {form_data.get('drpHocKy')}, year: {form_data.get('drpNamHoc')}, week: {form_data.get('drpTuan')}")

            # Submit form để xuất Excel
            form_action = form.get('action') or response.url
            if not form_action.startswith('http'):
                from urllib.parse import urljoin
                form_action = urljoin(response.url, form_action)
            
            print(f"[DEBUG] Submitting form to: {form_action}")
            
            # Thêm headers cần thiết cho ASP.NET
            headers = {
                'Content-Type': 'application/x-www-form-urlencoded',
                'Referer': response.url,
                'Origin': f"{self.base_url}",
            }
            
            download_response = self.session.post(form_action, data=form_data, headers=headers, timeout=30)
            
            print(f"[DEBUG] Excel download response status: {download_response.status_code}")
            print(f"[DEBUG] Content-Type: {download_response.headers.get('content-type', 'N/A')}")
            
            # Kiểm tra response có phải file Excel không
            content_type = download_response.headers.get('content-type', '').lower()
            content_disposition = download_response.headers.get('content-disposition', '').lower()
            
            is_excel = (
                any(excel_type in content_type for excel_type in ['excel', 'spreadsheet', 'application/vnd.ms-excel', 'application/vnd.openxmlformats']) or
                'excel' in content_disposition or
                '.xls' in content_disposition
            )
            
            if download_response.status_code == 200 and is_excel:
                print(f"[DEBUG] ✅ Received Excel file successfully!")
                
                # Parse Excel với pandas
                import pandas as pd
                import io
                
                try:
                    # Đọc Excel từ bytes
                    excel_data = io.BytesIO(download_response.content)

                    # --- BƯỚC 1: Tìm header row chính xác ---
                    # Đọc trước một vài hàng để tìm header
                    df_temp = pd.read_excel(excel_data, engine='xlrd', header=None, nrows=15) # Đọc 15 hàng đầu tiên

                    header_row_idx = None
                    timetable_keywords = ['STT', 'Lớp học phần', 'Học phần', 'Thời gian', 'Địa điểm', 'Giảng viên', 'Thứ', 'Tiết học']

                    for i, row in df_temp.iterrows():
                        # Chuyển đổi tất cả các giá trị trong hàng thành chuỗi và nối lại
                        row_text = ' '.join(str(cell) for cell in row.values if pd.notna(cell)).upper()
                        keyword_count = sum(1 for keyword in timetable_keywords if keyword.upper() in row_text)

                        if keyword_count >= 3:  # Tìm thấy hàng có ít nhất 3 keywords phù hợp
                            header_row_idx = i
                            print(f"[DEBUG] Excel: Found header row at index {i}: {keyword_count} keywords. Row content: {row_text[:100]}...")
                            break

                    if header_row_idx is None:
                        print(f"[ERROR] Excel: Could not find a suitable header row. Falling back to default.")
                        header_row_idx = 0 # Fallback to first row if no clear header found

                    # --- BƯỚC 2: Đọc lại toàn bộ file với header đã tìm được ---
                    excel_data.seek(0) # Reset stream position
                    df = pd.read_excel(excel_data, engine='xlrd', header=header_row_idx)

                    print(f"[DEBUG] Excel parsed: {len(df)} rows, {len(df.columns)} columns")
                    print(f"[DEBUG] Final columns after header detection: {list(df.columns)}")

                    # --- BƯỚC 3: Chuẩn hóa tên cột ---
                    # Find the actual column name for 'Thứ' dynamically
                    thu_col_name = None
                    for col in df.columns:
                        if 'thứ' in str(col).lower():
                            thu_col_name = col
                            break

                    # Create a dictionary for renaming common typos and the dynamically found 'Thứ' column
                    rename_dict = {
                        'Giảng viên/ link meet': 'Giảng viên/ link meet',
                        'Địa điểm': 'Phòng',
                        'Tiết học': 'Tiết học'
                    }
                    if thu_col_name:
                        print(f"[DEBUG] Excel: Found 'Thứ' column as '{thu_col_name}'")
                        # Add it to the rename dictionary to standardize it to 'Thứ'
                        if thu_col_name != 'Thứ':
                            rename_dict[thu_col_name] = 'Thứ'
                    else:
                        print(f"[WARNING] Excel: Could not find a column for 'Thứ'. Date calculation might be incorrect.")

                    df.rename(columns=rename_dict, inplace=True)
                    print(f"[DEBUG] Columns after rename: {list(df.columns)}")

                    # --- BỔ SUNG: Trích xuất ngành từ file Excel nếu có ---
                    major_excel = ""
                    # 1. Kiểm tra các dòng đầu của df_temp (dùng để tìm header) để tìm ngành
                    for i, row in df_temp.iterrows():
                        row_text = ' '.join(str(cell) for cell in row.values if pd.notna(cell)).strip()
                        match = re.search(r'Ngành\s*:?\s*(.+)', row_text, re.IGNORECASE)
                        if match:
                            major_excel = match.group(1).strip()
                            print(f"[DEBUG] Excel: Found major in header: {major_excel}")
                            break
                    # 2. Nếu chưa có, thử tìm trong các dòng đầu của DataFrame chính (df)
                    if not major_excel:
                        for i in range(min(5, len(df))):
                            row_text = ' '.join(str(cell) for cell in df.iloc[i].values if pd.notna(cell)).strip()
                            match = re.search(r'Ngành\s*:?\s*(.+)', row_text, re.IGNORECASE)
                            if match:
                                major_excel = match.group(1).strip()
                                print(f"[DEBUG] Excel: Found major in main df: {major_excel}")
                                break
                    # 3. Nếu vẫn chưa có, thử lấy từ cột "Lớp học phần" nếu có định dạng đặc biệt
                    if not major_excel and 'Lớp học phần' in df.columns:
                        for val in df['Lớp học phần'].head(5):
                            if isinstance(val, str) and 'ngành' in val.lower():
                                match = re.search(r'Ngành\s*:?\s*(.+)', val, re.IGNORECASE)
                                if match:
                                    major_excel = match.group(1).strip()
                                    print(f"[DEBUG] Excel: Found major in 'Lớp học phần': {major_excel}")
                                    break
                    if not major_excel:
                        major_excel = "Chưa cập nhật"

                    # Convert DataFrame to list of dictionaries
                    mapped_data = []
                    current_week_info = {
                        "from_date": "",
                        "to_date": "",
                        "week_number": ""
                    }
                    last_thu = "" # Biến để lưu giá trị "Thứ" cuối cùng

                    for _, row in df.iterrows():
                        row_dict = row.to_dict()

                        # Clean up values (remove NaN, strip whitespace)
                        for key, value in row_dict.items():
                            if pd.isna(value):
                                row_dict[key] = ""
                            else:
                                row_dict[key] = str(value).strip()

                        # --- Xử lý carry-forward cho cột "Thứ" để xử lý merged cells ---
                        current_thu = row_dict.get('Thứ', '').strip()
                        if current_thu:
                            last_thu = current_thu # Cập nhật giá trị "Thứ" mới nhất
                        else:
                            # Nếu "Thứ" trống, sử dụng giá trị cuối cùng đã thấy
                            row_dict['Thứ'] = last_thu
                        # --- Kết thúc xử lý carry-forward ---

                        # Skip rows that are clearly not data (e.g., empty STT or week rows)
                        stt = row_dict.get('STT', '').strip()
                        lop_hoc_phan = row_dict.get('Lớp học phần', '').strip()
                        thoi_gian = row_dict.get('Thời gian', '').strip() # Get 'Thời gian' column

                        is_week_row = 'tuần' in lop_hoc_phan.lower() and 'đến' in lop_hoc_phan.lower()

                        if is_week_row:
                            # Extract week and date range from week row
                            week_match = re.search(r'Tuần (\d+) \(((\d{2}/\d{2}/\d{4}) đến (\d{2}/\d{2}/\d{4}))\)', lop_hoc_phan)
                            if week_match:
                                current_week_info['week_number'] = week_match.group(1)
                                current_week_info['from_date'] = week_match.group(3)
                                current_week_info['to_date'] = week_match.group(4)
                                try:
                                    # Parse the from_date into a datetime object for calculation
                                    current_week_info['start_datetime'] = datetime.strptime(current_week_info['from_date'], '%d/%m/%Y')
                                except ValueError:
                                    current_week_info['start_datetime'] = None
                                print(f"[DEBUG] Excel: Extracted week info: {current_week_info}")
                            else:
                                print(f"[DEBUG] Excel: Could not parse week info from: {lop_hoc_phan}")
                            continue # Skip to next row after processing week info

                        # --- Mapping và chuẩn hóa dữ liệu ---
                        # Check if it's a valid subject row (STT is a number and lopHocPhan is not empty)
                        is_valid_stt = False
                        if stt:
                            try:
                                # Allow both integer and float STT values
                                float(stt)
                                is_valid_stt = True
                            except ValueError:
                                pass # Not a valid number

                        # Skip rows that are clearly not subject data (e.g., empty STT or lopHocPhan, or non-numeric STT)
                        if not is_valid_stt or not lop_hoc_phan:
                            print(f"[DEBUG] Excel: Skipping non-subject row (STT: '{stt}', lopHocPhan: '{lop_hoc_phan}')")
                            continue

                        mapped_item = {
                            "stt": row_dict.get('STT', ''),
                            "lopHocPhan": lop_hoc_phan,
                            "maHP": '',
                            "tenHP": '',
                            "soTC": row_dict.get('Số TC', ''),
                            "thu": row_dict.get('Thứ', ''),
                            "tiet": row_dict.get('Tiết học', ''),
                            "phong": row_dict.get('Phòng', ''),
                            "giangVien": '',
                            "meetLink": '',
                            "siSo": row_dict.get('Sĩ số', ''),
                            "soDK": row_dict.get('Số ĐK', ''),
                            "hocPhi": row_dict.get('Học phí', ''),
                            "ghiChu": row_dict.get('Ghi chú', ''),
                            "from_date": '', # Khởi tạo rỗng, sẽ được tính toán sau
                            "to_date": '',     # Khởi tạo rỗng, sẽ được tính toán sau
                            "week_number": current_week_info['week_number'], # Inherit from current_week_info
                            "lesson_type": ''
                        }

                        # Trích xuất mã học phần và tên học phần
                        print(f"[DEBUG] Excel: lopHocPhan (repr): {repr(lop_hoc_phan)}")
                        ma_hp_match = None
                        # Find all matches and take the last one
                        # Sửa regex: dùng \( thay vì \\( để thoát dấu ngoặc đơn
                        all_matches = list(re.finditer(r'\(([^)]+)\)|\[([^\]]+)\]', lop_hoc_phan))
                        if all_matches:
                            ma_hp_match = all_matches[-1] # Take the last match

                        if ma_hp_match:
                            mapped_item['maHP'] = (ma_hp_match.group(1) or ma_hp_match.group(2)).strip()
                            # Get the part before the course code (e.g., "Cơ sở dữ liệu-1-25 ")
                            potential_ten_hp = lop_hoc_phan[:ma_hp_match.start()].strip()
                            print(f"[DEBUG] Excel: potential_ten_hp: '{potential_ten_hp}'")
                            # Remove the trailing "-X-YY" pattern if it exists
                            cleaned_ten_hp = re.sub(r'-\d+-\d+.*$', '', potential_ten_hp).strip()
                            print(f"[DEBUG] Excel: cleaned_ten_hp: '{cleaned_ten_hp}'")
                            mapped_item['tenHP'] = cleaned_ten_hp
                        else:
                            # If no course code found, use the entire string as the name
                            mapped_item['tenHP'] = lop_hoc_phan.strip()
                        print(f"[DEBUG] Excel: lopHocPhan: '{lop_hoc_phan}' -> maHP: '{mapped_item['maHP']}', tenHP: '{mapped_item['tenHP']}'")

                        # Trích xuất Giảng viên và Link Meet
                        giang_vien_raw = row_dict.get('Giảng viên/ link meet', '')
                        print(f"[DEBUG] Excel: Raw Giảng viên/ link meet (repr): {repr(giang_vien_raw)}")

                        meet_link_found = False
                        # Try to split by newline first, as it's a common pattern
                        if '\n' in giang_vien_raw:
                            parts = giang_vien_raw.split('\n', 1)
                            potential_link_part = parts[1].strip() if len(parts) > 1 else ''
                            url_match = re.search(r'(https?://)?(?:www\.)?meet\.google\.com/[^\s]+', potential_link_part)
                            if url_match:
                                full_link = url_match.group(0)
                                if not full_link.startswith('http'):
                                    full_link = 'https://' + full_link
                                mapped_item['meetLink'] = full_link
                                mapped_item['giangVien'] = parts[0].strip()
                                meet_link_found = True
                                print(f"[DEBUG] Excel: Extracted meetLink (split by newline): {mapped_item['meetLink']}, giangVien: {mapped_item['giangVien']}")
                        
                        if not meet_link_found:
                            # Fallback to searching in the original string if newline split didn't work or no newline
                            url_match = re.search(r'(https?://)?(?:www\.)?meet\.google\.com/[^\s]+', giang_vien_raw)
                            if url_match:
                                full_link = url_match.group(0)
                                if not full_link.startswith('http'):
                                    full_link = 'https://' + full_link
                                mapped_item['meetLink'] = full_link
                                mapped_item['giangVien'] = giang_vien_raw.replace(url_match.group(0), '').strip()
                                print(f"[DEBUG] Excel: Extracted meetLink (regex on original): {mapped_item['meetLink']}, giangVien: {mapped_item['giangVien']}")
                            else:
                                mapped_item['giangVien'] = giang_vien_raw.strip()
                                print(f"[DEBUG] Excel: No direct URL in Giảng viên/ link meet. giangVien: {mapped_item['giangVien']}")

                        # Chuẩn hóa Thứ và Tiết
                        thu_value = mapped_item['thu'].strip()
                        # Cải thiện logic để xử lý số thập phân (ví dụ: "2.0") thành số nguyên
                        if thu_value and thu_value.replace('.', '', 1).isdigit():
                            try:
                                thu_int = int(float(thu_value))
                                mapped_item['thu'] = f"Thứ {thu_int}"
                            except ValueError:
                                # Fallback if conversion fails
                                mapped_item['thu'] = thu_value
                        else:
                            mapped_item['thu'] = thu_value

                        mapped_item['tiet'] = mapped_item['tiet'].replace(' --> ', '-').replace('-->', '-').strip()
                        # Cập nhật thoiGian chỉ chứa thông tin Thứ
                        mapped_item['thoiGian'] = mapped_item['thu'].strip()

                        # Định nghĩa ánh xạ tiết học sang thời gian cụ thể
                        TIET_TIMES = {
                            1: {"start": "6:45", "end": "7:35"},
                            2: {"start": "7:40", "end": "8:30"},
                            3: {"start": "8:40", "end": "9:30"},
                            4: {"start": "9:40", "end": "10:30"},
                            5: {"start": "10:35", "end": "11:25"},
                            6: {"start": "13:00", "end": "13:50"},
                            7: {"start": "13:55", "end": "14:45"},
                            8: {"start": "14:55", "end": "15:45"},
                            9: {"start": "15:55", "end": "16:45"},
                            10: {"start": "16:50", "end": "17:40"},
                            11: {"start": "18:15", "end": "19:05"},
                            12: {"start": "19:10", "end": "20:00"},
                            13: {"start": "20:05", "end": "20:55"},
                            14: {"start": "20:20", "end": "21:10"},
                            15: {"start": "21:20", "end": "22:10"},
                        }

                        # Bổ sung logic xác định buổi học dựa trên tiết học và thời gian cụ thể
                        buoi_hoc = "Không xác định"
                        tiet_str = mapped_item['tiet']
                        if tiet_str:
                            # Extract all numbers from the 'tiet' string
                            tiet_numbers = [int(s) for s in re.findall(r'\d+', tiet_str) if s.isdigit()]
                            if tiet_numbers:
                                min_tiet = min(tiet_numbers)
                                max_tiet = max(tiet_numbers)

                                start_time = TIET_TIMES.get(min_tiet, {}).get("start", "")
                                end_time = TIET_TIMES.get(max_tiet, {}).get("end", "")

                                if start_time and end_time:
                                    buoi_hoc = f"{start_time} - {end_time}"
                            
                        mapped_item['buoiHoc'] = buoi_hoc
                        print(f"[DEBUG] Excel: Tiết: {tiet_str}, Buổi học: {buoi_hoc}")

                        # Tính toán ngày cụ thể cho môn học dựa trên 'Thứ' và ngày bắt đầu của tuần
                        if current_week_info.get('start_datetime') and mapped_item['thu']:
                            try:
                                # Extract the digit from "Thứ X"
                                thu_digit_match = re.search(r'\d+', mapped_item['thu'])
                                if thu_digit_match:
                                    thu_int = int(thu_digit_match.group(0))
                                    
                                    # Map 'Thứ' to a 0-indexed day of the week (0=Monday, 6=Sunday)
                                    day_of_week_index = thu_int - 2
                                    if 'chủ nhật' in mapped_item['thu'].lower():
                                        day_of_week_index = 6
                                    
                                    if 0 <= day_of_week_index <= 6:
                                        week_start_weekday = current_week_info['start_datetime'].weekday()
                                        days_to_add = (day_of_week_index - week_start_weekday + 7) % 7
                                        subject_date = current_week_info['start_datetime'] + timedelta(days=days_to_add)
                                        mapped_item['from_date'] = subject_date.strftime('%d/%m/%Y')
                                        mapped_item['to_date'] = subject_date.strftime('%d/%m/%Y')
                                        print(f"[DEBUG] Excel: Calculated subject date: {mapped_item['from_date']} for {mapped_item['thu']}")
                                    else:
                                        print(f"[DEBUG] Excel: Invalid 'Thứ' value after digit extraction: {mapped_item['thu']}")
                                else:
                                    print(f"[DEBUG] Excel: Could not extract digit from 'Thứ': {mapped_item['thu']}")
                            except (ValueError, TypeError) as e:
                                print(f"[ERROR] Excel: Error parsing 'Thứ' or calculating date: {e}")
                        else:
                            print(f"[DEBUG] Excel: Skipping date calculation due to missing week info or 'Thứ'.")

                        # Trích xuất thông tin loại buổi học từ cột 'Thời gian' hoặc 'Ghi chú'
                        # Try to extract lesson type from 'Thời gian' column first
                        lesson_type_match = re.search(r'\(([A-Z]{2,3})\)', thoi_gian)
                        if lesson_type_match:
                            mapped_item['lesson_type'] = lesson_type_match.group(1)
                            print(f"[DEBUG] Excel: Extracted lesson type: {mapped_item['lesson_type']}")
                        else:
                            # Fallback: try to find in 'Ghi chú' column
                            note_match = re.search(r'\(([A-Z]{2,3})\)', row_dict.get('Ghi chú', ''))
                            if note_match:
                                mapped_item['lesson_type'] = note_match.group(1)
                                print(f"[DEBUG] Excel: Extracted lesson type from note: {mapped_item['lesson_type']}")
            
                        mapped_data.append(mapped_item)
                        print(f"[DEBUG] Excel: Added final mapped item: {mapped_item.get('tenHP', 'N/A')} (Week: {mapped_item['week_number']}, Dates: {mapped_item['from_date']} - {mapped_item['to_date']})")
                    
                    return {
                        "error": False,
                        "timetableData": mapped_data,
                        "originalColumns": list(df.columns), # Return final columns after rename
                        "source": "excel",
                        "totalRows": len(mapped_data),
                        "major": major_excel
                    }
                except Exception as e:
                    print(f"[ERROR] Excel parsing exception: {str(e)}")
                    return self._handle_error(f"Lỗi khi phân tích file Excel: {str(e)}", 500)
            
            return {
                "error": True,
                "message": "Không thể tải file Excel. Vui lòng thử lại sau."
            }
        except requests.exceptions.Timeout:
            return self._handle_error("Kết nối timeout khi lấy thời khóa biểu", 408)
        except requests.exceptions.ConnectionError:
            return self._handle_error("Lỗi kết nối khi lấy thời khóa biểu", 503)
        except Exception as e:
            print(f"[ERROR] Timetable Excel exception: {str(e)}")
            print(f"[ERROR] Traceback: {traceback.format_exc()}")
            return self._handle_error("Error fetching timetable Excel file", 500)
