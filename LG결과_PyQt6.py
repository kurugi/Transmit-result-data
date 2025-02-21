import os
import sys
import openpyxl
from PyQt6.QtCore import QObject, QThread, pyqtSignal, Qt
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QPushButton, QLabel, QProgressBar, QFileDialog, QMessageBox
)

# =============================================================================
# 전역 상수 (파일명)
# =============================================================================
TARGET_FILE = "LG결과.xlsx"      # 결과 저장 파일 (고정)
OPINION_FILE = "병원소견.xlsx"      # 병원소견 파일 (고정)

# =============================================================================
# Excel 매핑 정보
# =============================================================================
# 병원소견.xlsx의 A1~A22 컬럼을 LG결과.xlsx의 컬럼명으로 매핑
opinion_columns = {
    "A1": "MDC_DECI", "A2": "STATE", "A3": "RECIPE1", "A4": "RECIPE2", "A5": "RECIPE3",
    "A6": "RECIPE4", "A7": "RECIPE5", "A8": "MDC_GRADE1", "A9": "OPIN_CODE1", "A10": "OPIN_DESCRIPT1",
    "A11": "MDC_GRADE2", "A12": "OPIN_CODE2", "A13": "OPIN_DESCRIPT2", "A14": "MDC_GRADE3",
    "A15": "OPIN_CODE3", "A16": "OPIN_DESCRIPT3", "A17": "MDC_GRADE4", "A18": "OPIN_CODE4",
    "A19": "OPIN_DESCRIPT4", "A20": "MDC_GRADE5", "A21": "OPIN_CODE5", "A22": "OPIN_DESCRIPT5"
}

# 병원결과.xlsx의 컬럼을 LG결과.xlsx의 컬럼명으로 매핑 (필요한 열만 예시로 포함)
column_map = {
    "사원번호": "EMP_NO",
    "주민등록번호": "SSN",
    "진료일자": "MDC_DATE",    
    "HA001": "BM01",
    "HA002": "BM02",
    "HA004": "BM04",
    "FAT": "BM05",
    "HA007": "BM07",
    "L90001": "BT01",
    "L70013": "BT02",
    "HB001": "CV02",
    "HB002": "CV01",
    "A8001": "CV03",
    "E6541": "CV04",
    "HO001": "OE101",
    "HO002": "OE102",
    "HO003": "OE103",
    "HO004": "OE104",
    "HO011": "OE205",
    "HO005": "OE301",
    "HO006": "OE302",
    "OHA01": "AM103",
    "OHA09": "AM104",
    "OHA02": "AM105",
    "OHA10": "AM106",
    "OHA03": "AM107",
    "OHA11": "AM108",
    "OHA04": "AM121",
    "OHA12": "AM122",
    "OHA05": "AM111",
    "OHA13": "AM112",
    "OHA06": "AM115",
    "OHA14": "AM116",
    "L2012": "CB101",
    "L2014": "CB110",
    "L07003501": "CB112",
    "L20220": "CB113",
    "L20141": "CB104",
    "L20142": "CB105",
    "L20143": "CB106",
    "L2015": "CB107",
    "L2016": "CB108",
    "L2011": "CB109",
    "L20191": "CB203",
    "L20193": "CB204",
    "L20192": "CB205",
    "L20194": "CB206",
    "L3012": "DM04",
    "L3231": "DM07",
    "L3015": "LP01",
    "L3081": "LP02",
    "L3082": "LP03",
    "L3083": "LP04",
    "L3016": "LF13",
    "LAC10401": "LF14",
    "A009": "LF15",
    "L3018": "LF03",
    "L3019": "LF04",
    "L3033": "LF05",    
    "L3020": "LF06",
    "L3062": "LF11",
    "LAC10402": "LF10",
    "L3051": "LF16",
    "LAC114": "LF17",
    "LAC158": "LF19",
    "N20501": "VE105",
    "N20502": "VE102",
    "N20511": "VE201",
    "L170360": "VE106",
    "L170380": "VE107",
    "L3032": "RF03",
    "L3013": "RF02",
    "LAC104031": "RF04",
    "LAC13701": "EL01",
    "L3031": "EL02",
    "L3041": "EL03",
    "L3042": "EL04",
    "RU103": "TF01",
    "N20006": "TF06",
    "N20003": "TF03",
    "LIS11001": "VE301",
    "L3092": "SY03",
    "L51821": "CE01",
    "L51811": "CE02",
    "N206101": "CE03",
    "N20609": "CE04",
    "N206151": "CE05",
    "L150026": "CE06",
    "L3014": "GT01",
    "L5114": "RA03",
    "L6103": "UA101",
    "L6102": "UA102",
    "L6110": "UA103",
    "L6104": "UA106",
    "L6107": "UA108",
    "L6108": "UA301",
    "L6109": "UA302",
    "L6105": "UA401",
    "L61132": "UA201",
    "L61131": "UA202",
    "L61133": "UA203",
    "TK531": "SE02",
    "TK02": "SE01",
    "TK03": "SE03",
    "RZ901A2": "BD01",
    "RP201": "RE101",
    "L9742HPC": "RE200",
    "RC4011": "GI303",
    "S1005": "GI201",
    "L5237": "GI203",
    "C5602": "GI205",
    "RU401": "US02",
    "S2322": "US03",
    "RU902A": "US04",
    "RU505": "US05",
    "S1008B": "US07",
    "CTN710": "US08",
    "N456232": "US09",
    "RC101": "US10",
    "C5601": "US13",
    # "MC090": "GY03",
    "RP20BBA": "GY03",
    "RU201": "GY04",
    "P30001": "GY05",
    "RC94HL": "RE402",
    "L2013": "CB102",
    "L3021": "LF12",
    "L1820": "LF08",
    "LAC162": "RA01",
    "TH01": "RA02",
    "L20221": "CB111",
    "L20190": "CB201",
    "RC94HC": "RE403",
    "L6106": "UA402",
    "L5237":"G1203",
    "RN801":"US15",
    "N456000":"US16",
    "N455004":"US17",
    "CTN711H":"US20",
    "SE60":"US11",
    "TX0B300":"GY07"
}

# =============================================================================
# 숫자형 변환이 필요한 열 목록
# =============================================================================
numeric_columns = {
    "MDC_DATE", "BM01", "BM02", "BM04", "BM05", "BM07", "CV02", "CV01", "CV03",
    "OE101", "OE102", "OE103", "OE104", "OE301", "OE302", "AM103", "AM104", "AM105",
    "AM106", "AM107", "AM108", "AM121", "AM122", "AM111", "AM112", "AM115", "AM116",
    "CB101", "CB110", "CB112", "CB113", "CB104", "CB105", "CB106", "CB107", "CB108",
    "CB109", "CB203", "CB204", "CB205", "CB206", "DM04", "DM07", "LP01", "LP02",
    "LP03", "LP04", "LF13", "LF14", "LF15", "LF04", "LF05", "LF06", "LF11", "LF10",
    "LF16", "LF19", "RF03", "RF02", "RF04", "EL01", "EL02", "EL03", "EL04", "TF06",
    "TF03", "VE301", "SY03", "CE01", "CE02", "CE03", "CE04", "CE05", "CE06", "GT01",
    "RA01", "RA02", "RA03", "UA101", "UA102", "CB102", "CB111", "CB201", "LF03",
    "LF12", "LF08", "LF17", "BM06"
}

# =============================================================================
# Excel 처리 관련 함수들
# =============================================================================
def load_hospital_opinions():
    """
    병원소견.xlsx 파일에서 (사원번호, 주민등록번호)를 키로 하여
    A1 ~ A22 (인덱스 4 ~ 25) 데이터를 읽어 딕셔너리로 반환합니다.
    """
    wb = openpyxl.load_workbook(OPINION_FILE, data_only=True)
    ws = wb.active
    opinion_data = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        ho_no = str(row[0]).strip()
        jumin = str(row[1]).strip()
        # 인덱스 4~25 : A1 ~ A22 데이터
        opinion_values = list(row[4:26])
        opinion_data[(ho_no, jumin)] = opinion_values
    wb.close()
    return opinion_data

def convert_to_number_safely(value):
    """
    숫자형 변환을 시도합니다.
    변환에 실패하면 원래 값을 반환합니다.
    """
    try:
        return float(value) if '.' in str(value) else int(value)
    except (ValueError, TypeError):
        return value

def process_emp_no(emp_no_value):
    """
    EMP_NO 값이 전적으로 숫자이면 int 타입으로,
    그렇지 않으면 문자열로 반환합니다.
    """
    if emp_no_value is None:
        return None
    emp_no_str = str(emp_no_value).strip()
    return int(emp_no_str) if emp_no_str.isdigit() else emp_no_str

def load_workbooks(source_file):
    """
    선택한 병원결과 파일과 고정된 LG결과, 병원소견 파일을 로드하고
    각 워크북과 워크시트를 반환합니다.
    """
    source_wb = openpyxl.load_workbook(source_file, data_only=True)
    source_ws = source_wb.active
    target_wb = openpyxl.load_workbook(TARGET_FILE)
    target_ws = target_wb.active
    return source_wb, source_ws, target_wb, target_ws

def clear_existing_data(target_ws):
    """
    LG결과.xlsx의 3행 이후의 모든 데이터를 삭제합니다.
    """
    for row in target_ws.iter_rows(min_row=3, max_row=target_ws.max_row):
        for cell in row:
            cell.value = None

def auto_adjust_column_width(ws):
    """
    워크시트의 각 열에 대해 셀 내용의 최대 길이에 따라 열 너비를 자동 조절합니다.
    """
    # 각 열을 순회합니다.
    for col in ws.columns:
        max_length = 0
        # 첫번째 셀의 column_letter를 가져옵니다.
        column = col[0].column_letter
        for cell in col:
            if cell.value:
                try:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
                except:
                    pass
        # 약간의 여유 공간을 더해 열 너비 설정
        adjusted_width = max_length + 2
        ws.column_dimensions[column].width = adjusted_width


def process_data_with_progress(source_ws, target_ws, opinion_data, column_map, progress_callback):
    """
    병원결과.xlsx의 데이터를 읽어 LG결과.xlsx에 기록하는 동안,
    progress_callback(progress: int)를 통해 진행률(0~100)을 업데이트합니다.
    """
    # 1) 헤더 읽기: 병원결과는 3행, LG결과는 2행
    header = [str(cell.value).strip() if cell.value else "" for cell in source_ws[3]]
    target_header = [str(cell.value).strip() if cell.value else "" for cell in target_ws[2]]
    
    # 2) LG결과의 헤더에 opinion_columns 항목이 없으면 추가
    for opin_target in opinion_columns.values():
        if opin_target not in target_header:
            target_header.append(opin_target)
    
    # 3) 변경된 헤더를 2행에 다시 기록
    for i, header_val in enumerate(target_header, start=1):
        target_ws.cell(row=2, column=i, value=header_val)
    
    rows = []
    data_rows = list(source_ws.iter_rows(min_row=5, values_only=True))
    total_rows = len(data_rows)
    total_steps = total_rows * 2  # 데이터 처리 + 쓰기 단계
    current_step = 0

    # 4) 데이터 처리 및 병원소견 데이터 병합
    for row in data_rows:
        row_data = {}
        emp_no_processed = None
        ssn_prefix = None
        sex_no = None  # 주민등록번호 기반 성별 코드 변수
        bm01_value = None  # BM01 값 저장 변수
        bm06_value = None  # BM06 계산 결과 저장 변수

        # 열 매핑에 따라 데이터 처리
        for source_col, target_col in column_map.items():
            if source_col in header and target_col in target_header:
                source_idx = header.index(source_col)
                value = row[source_idx] if row[source_idx] is not None else ""
                
                if target_col == "SSN":
                    ssn_prefix = str(value)[:8] if value else None
                    row_data[target_col] = ssn_prefix

                    # 주민등록번호 8번째 자리 확인 후 sex_no 값 결정
                    if ssn_prefix and len(ssn_prefix) >= 8:
                        eighth_digit = ssn_prefix[7]  # 8번째 자리 숫자
                        if eighth_digit in {'1', '3', '5'}:
                            sex_no = 22
                        elif eighth_digit in {'2', '4', '6'}:
                            sex_no = 21

                elif target_col == "MDC_DATE" and value:
                    if isinstance(value, str):
                        value = value.replace("-", "")
                    try:
                        value = int(value)
                    except ValueError:
                        value = None
                    row_data[target_col] = value
                elif target_col in numeric_columns and value != "":
                    row_data[target_col] = convert_to_number_safely(value)
                elif target_col == "EMP_NO":
                    emp_no_processed = process_emp_no(value)
                    row_data[target_col] = emp_no_processed
                else:
                    row_data[target_col] = value

                # BM01 값 저장
                if target_col == "BM01":
                    bm01_value = row_data[target_col]

        # 병원소견 데이터 병합 (키: (EMP_NO 문자열, 주민등록번호 앞 8자리))
        if emp_no_processed is not None and ssn_prefix:
            opinion_key = (str(emp_no_processed).strip(), ssn_prefix.strip())
            if opinion_key in opinion_data:
                opinion_values = opinion_data[opinion_key]
                for col_idx, opin_target in enumerate(opinion_columns.values()):
                    if col_idx < len(opinion_values):
                        row_data[opin_target] = opinion_values[col_idx]

        # BM06 값 계산: (BM01 / 100) * (BM01 / 100) * sex_no, 소수점 2자리 반올림
        if bm01_value is not None and sex_no is not None:
            try:
                bm06_value = round((float(bm01_value) / 100) * (float(bm01_value) / 100) * sex_no, 1)
            except (ValueError, TypeError):
                bm06_value = None  # 값이 유효하지 않으면 None 설정

        row_data["BM06"] = bm06_value  # BM06 값 추가
        rows.append(row_data)
        current_step += 1
        progress_callback(int(current_step / total_steps * 100))
    
    # 5) 정렬 후 LG결과.xlsx에 데이터 기록 (3행부터)
    rows.sort(key=lambda x: (x.get("MDC_DATE", float('inf')), str(x.get("EMP_NO", ""))))
    clear_existing_data(target_ws)
    
    for i, row_data in enumerate(rows, start=3):
        for target_col in target_header:
            target_idx = target_header.index(target_col)
            value = row_data.get(target_col, "")
            target_ws.cell(row=i, column=target_idx + 1, value=value)
        current_step += 1
        progress_callback(int(current_step / total_steps * 100))
    
    # 6) 데이터 기록이 완료되면 열 너비 자동 조절
    auto_adjust_column_width(target_ws)
    
    return rows


def save_and_open(target_wb):
    """
    LG결과.xlsx를 저장하고, 엑셀을 열 때 B3 셀을 선택하도록 설정합니다.
    """
    target_ws = target_wb.active  # 현재 활성화된 시트 가져오기
    target_ws.sheet_view.selection[0].activeCell = "A3"  # 활성 셀을 B3으로 설정
    target_ws.sheet_view.selection[0].sqref = "A3"  # 선택된 영역을 B3으로 설정

    target_wb.save(TARGET_FILE)  # 저장
    target_wb.close()

    try:
        if os.name == 'nt':  # Windows
            os.startfile(TARGET_FILE)
        elif sys.platform == "darwin":  # macOS
            os.system(f"open {TARGET_FILE}")
        else:  # Linux 등
            os.system(f"xdg-open {TARGET_FILE}")
    except Exception as e:
        QMessageBox.critical(None, "실행 오류", f"LG결과.xlsx 실행 중 오류 발생:\n{str(e)}")
    



# =============================================================================
# Worker 클래스 (백그라운드에서 Excel 변환 작업 수행)
# =============================================================================
class ConversionWorker(QObject):
    progress = pyqtSignal(int)  # 진행률 (0 ~ 100)
    finished = pyqtSignal()     # 작업 완료 시 시그널
    error = pyqtSignal(str)     # 오류 발생 시 시그널

    def __init__(self, source_file):
        super().__init__()
        self.source_file = source_file

    def run(self):
        """
        선택한 병원결과 파일을 로드하여 병원소견 데이터를 병합하고,
        LG결과.xlsx에 기록한 후 실행합니다.
        """
        try:
            source_wb, source_ws, target_wb, target_ws = load_workbooks(self.source_file)
            opinion_data = load_hospital_opinions()
            clear_existing_data(target_ws)
            process_data_with_progress(source_ws, target_ws, opinion_data, column_map, self.progress.emit)
            save_and_open(target_wb)
            source_wb.close()
            self.finished.emit()
        except Exception as e:
            self.error.emit(str(e))

# =============================================================================
# MainWindow 클래스 (UI)
# =============================================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("(병원결과&병원소견) ==> LG결과")
        self.source_file = None  # 선택된 병원결과 파일 경로
        self.worker = None
        self.thread = None
        self.setup_ui()

    def setup_ui(self):
        """UI 위젯들을 생성하고 배치합니다."""
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()
        layout.setSpacing(30)    # 위젯 간 간격 30px
        layout.addSpacing(30)    # 상단 여백 30px

        # 최소 너비 설정
        min_width = 380

        # 파일 선택 버튼
        self.select_button = QPushButton("1. 병원결과 파일 선택")
        self.select_button.setStyleSheet("background-color: #90EE90; font-size: 15px;")
        self.select_button.clicked.connect(self.select_file)
        layout.addWidget(self.select_button)

        # 변환 시작 버튼
        self.convert_button = QPushButton("2. 변환 시작")
        self.convert_button.setStyleSheet("background-color: skyblue; font-size: 15px;")
        self.convert_button.clicked.connect(self.convert_data)
        layout.addWidget(self.convert_button)

        # 진행률 표시 Progress Bar (텍스트: 오른쪽 하단 정렬)
        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignBottom)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        # 상태 레이블
        self.status_label = QLabel("상태: 대기 중")
        self.status_label.setStyleSheet("font-size: 15px;")
        layout.addWidget(self.status_label)

        # 프로그램 종료 버튼
        self.quit_button = QPushButton("프로그램 종료")
        self.quit_button.setStyleSheet("background-color: #FF6F61; font-size: 15px;")
        self.quit_button.clicked.connect(QApplication.quit)
        layout.addSpacing(30)    # 하단 여백 30px
        layout.addWidget(self.quit_button)

        central_widget.setLayout(layout)

        # 창 최소 크기 설정 (너비 400px 고정, 높이는 가변)
        self.setMinimumWidth(min_width)
        self.setMinimumHeight(350)

    def select_file(self):
        """
        파일 선택 다이얼로그를 실행하여 사용자가 원하는 Excel 파일을 선택할 수 있도록 함.
        """
        file_name, _ = QFileDialog.getOpenFileName(
            self,
            "병원결과 파일 선택",
            "",
            "Excel Files (*.xlsx);;All Files (*)"
        )
        if not file_name:
            self.status_label.setText("상태: 파일 선택이 취소되었습니다.")
            return

        # 선택한 파일 확장자 확인 (xlsx 파일만 허용)
        if not file_name.lower().endswith('.xlsx'):
            QMessageBox.warning(self, "파일 선택 오류", "올바른 Excel 파일을 선택하세요. (xlsx 파일만 허용)")
            self.status_label.setText("상태: 대기 중")
            return

        # 파일 경로 저장
        self.source_file = file_name
        self.status_label.setText(f"상태: {os.path.basename(file_name)} 파일 선택됨")


    def convert_data(self):
        """
        Excel 변환 작업을 백그라운드에서 수행합니다.
        진행률은 ProgressBar를 통해 업데이트됩니다.
        """
        if not self.source_file:
            QMessageBox.warning(self, "파일 미선택", "먼저 병원결과 파일을 선택해 주세요.")
            return

        self.status_label.setText("상태: 변환 진행 중...")
        self.progress_bar.setValue(0)
        self.convert_button.setEnabled(False)

        # 백그라운드 작업을 위한 QThread와 Worker 생성 및 연결
        self.thread = QThread()
        self.worker = ConversionWorker(self.source_file)
        self.worker.moveToThread(self.thread)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.conversion_finished)
        self.worker.error.connect(self.conversion_error)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.start()

    def update_progress(self, value):
        """진행률을 업데이트합니다."""
        self.progress_bar.setValue(value)

    def conversion_finished(self):
        """변환 작업 완료 후 상태를 업데이트합니다."""
        self.convert_button.setEnabled(True)
        self.status_label.setText("상태: 변환 완료")
        # 변환 완료 후 LG결과.xlsx 파일이 자동 실행됩니다.

    def conversion_error(self, error_msg):
        """오류 발생 시 메시지 박스로 알립니다."""
        self.convert_button.setEnabled(True)
        self.status_label.setText("상태: 오류 발생")
        QMessageBox.critical(self, "오류", f"오류 발생:\n{error_msg}")

# =============================================================================
# 메인 실행 코드
# =============================================================================
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


# pyinstaller --onefile --noconsole LG결과_PyQt6.py 