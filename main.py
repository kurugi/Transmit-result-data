import openpyxl
from openpyxl.styles import Alignment

def convert_emp_no(value):
    """EMP_NO가 숫자형이면 숫자로, 문자형이면 그대로 반환"""
    try:
        return int(value) if isinstance(value, (int, float)) and value == int(value) else str(value)
    except ValueError:
        return str(value)

def convert_to_numeric(value):
    """숫자형 데이터 변환 (변환 가능하면 float, 불가능하면 그대로)"""
    try:
        return float(value) if value not in (None, "", " ") else None
    except ValueError:
        return value

def truncate_ssn(value):
    """주민등록번호(SSN) 왼쪽 8자리만 출력"""
    return str(value)[:8] if isinstance(value, str) else value

def convert_mdc_date(value):
    """MDC_DATE에서 '-' 제거 후 숫자로 변환"""
    if isinstance(value, str):
        cleaned_value = value.replace("-", "")
        try:
            return int(cleaned_value)  # 변환 가능하면 정수 변환
        except ValueError:
            return cleaned_value  # 변환 불가능하면 문자열 그대로 반환
    return value  # None 또는 다른 타입 값 유지

def extract_sex_no(ssn_value):
    """SSN 값에서 왼쪽에서 8번째 문자를 추출하여 sex_no 변수에 저장"""
    if isinstance(ssn_value, str) and len(ssn_value) >= 8:
        sex_digit = ssn_value[7]  # 8번째 문자 (0-based index: 7)
        if sex_digit in {"1", "3", "5"}:
            return 22
        elif sex_digit in {"2", "4", "6"}:
            return 21
    return None  # 값이 없거나 길이가 8 미만인 경우 None 반환

def calculate_bm06(bm01_value, sex_no):
    """BM06 값을 계산"""
    try:
        return round((float(bm01_value) / 100) * (float(bm01_value) / 100) * sex_no, 1)
    except (ValueError, TypeError, ZeroDivisionError):
        return None  # 값이 없거나 계산 불가능한 경우 None 반환

def map_and_transfer_data(hospital_file, lg_file, output_file, column_map, numeric_columns, right_align_columns):
    # 병원결과.xlsx 로드
    wb_hospital = openpyxl.load_workbook(hospital_file)
    ws_hospital = wb_hospital.active
    
    # LG결과.xlsx 로드
    wb_lg = openpyxl.load_workbook(lg_file)
    ws_lg = wb_lg.active

    # Delete all rows from row 4 onwards
    ws_lg.delete_rows(4, ws_lg.max_row)

    # Save the cleared file
    wb_lg.save(output_file)
    print(f"All data from row 4 onward has been deleted in {output_file}.")
    
    # 병원결과 헤더 (3번째 행)
    hospital_headers = {cell.value: col_idx for col_idx, cell in enumerate(ws_hospital[3], start=1) if cell.value}
    
    # LG결과 헤더 (3번째 행)
    lg_headers = {cell.value: col_idx for col_idx, cell in enumerate(ws_lg[3], start=1) if cell.value}
    
    # 컬럼 매핑 적용 (병원결과 헤더 -> column_map -> LG결과 헤더)
    transformed_mapping = {
        hospital_headers[header]: lg_headers[column_map[header]]
        for header in column_map if header in hospital_headers and column_map[header] in lg_headers
    }

    # "SSN", "BM01", "BM06" 열 찾기 (존재하면 열 번호 저장)
    ssn_col_idx = lg_headers.get("SSN", None)
    bm01_col_idx = lg_headers.get("BM01", None)
    bm06_col_idx = lg_headers.get("BM06", None)

    # 병원결과 데이터 읽기 (5번째 행부터)
    data_rows = list(ws_hospital.iter_rows(min_row=5, values_only=True))

    # LG결과 다음 빈 행 찾기 (4번째 행부터)
    start_row = max(4, ws_lg.max_row + 1)

    # 데이터 변환 및 입력
    for row_idx, row in enumerate(data_rows, start=start_row):
        new_row = [None] * len(lg_headers)  # LG결과 형식 맞추기
        sex_no = None  # 성별 코드 변수 초기화
        bm01_value = None  # BM01 값 초기화
        bm06_value = None  # BM06 값 초기화
        
        for hospital_col, lg_col in transformed_mapping.items():
            value = row[hospital_col - 1]
            
            # LG결과 헤더명 가져오기
            lg_column_name = list(lg_headers.keys())[list(lg_headers.values()).index(lg_col)]

            # EMP_NO 변환 적용
            if column_map.get(list(hospital_headers.keys())[list(hospital_headers.values()).index(hospital_col)]) == "EMP_NO":
                value = convert_emp_no(value)
            
            # 숫자형 변환 적용
            if lg_column_name in numeric_columns:
                value = convert_to_numeric(value)

            # SSN 변환 적용 (왼쪽 8자리만 출력)
            if lg_column_name == "SSN":
                value = truncate_ssn(value)

                # 성별 코드 (8번째 문자) 추출 및 변환
                sex_no = extract_sex_no(value)
            
            # MDC_DATE 변환 적용 ('-' 제거 후 숫자로 변환)
            if lg_column_name == "MDC_DATE":
                value = convert_mdc_date(value)

            # BM01 값 저장 (BM06 계산에 필요)
            if lg_column_name == "BM01":
                bm01_value = value
            
            # 값 입력
            cell = ws_lg.cell(row=row_idx, column=lg_col, value=value)

            # CE01, CE02, CD03 열을 오른쪽 정렬
            if lg_column_name in right_align_columns:
                cell.alignment = Alignment(horizontal="right")

        # BM06 값 계산 및 입력
        if bm06_col_idx is not None and bm01_value is not None and sex_no is not None:
            bm06_value = calculate_bm06(bm01_value, sex_no)
            ws_lg.cell(row=row_idx, column=bm06_col_idx, value=bm06_value)

        
    # 변환된 엑셀 저장
    wb_lg.save(output_file)
    print(f"변환된 데이터가 {output_file}에 저장되었습니다.")

# 오른쪽 정렬할 컬럼 목록
right_align_columns = {"CE01", "CE02", "CE03"}


# 숫자로 변환할 컬럼 목록
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

# 실행 예시
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

map_and_transfer_data("병원결과.xlsx", "LG결과.xlsx", "LG결과_변환.xlsx", column_map, numeric_columns, right_align_columns)