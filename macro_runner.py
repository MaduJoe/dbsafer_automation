# -*- coding: utf-8 -*-
import macro_performer
import macro_validator
import spreadsheet
import ast 
import shared
import my_utils
import pyautogui
import time
import log_handler

from database import Database
import progress
start_time = time.time()

# 로거 설정
logger = log_handler.setup_logger()

# 0. 실행/검증 클래스 인스턴스 생성 
performer = macro_performer.MacroPerformer()
validator = macro_validator.Validator()
sheet = spreadsheet.SpreadSheets()

# 1. XLSX 다운로드 
# xlsx_file = sheet.save_sheet()

# 1. 기존 파일 load 
xlsx_file = my_utils.get_xlsx()

# 2. JSON 으로 변형 
json_file = sheet.convert_xlsx_to_json(xlsx_file)
# json_file = my_utils.get_json('manager_20240909_202230')

# 3. JSON 읽기
tc_list = my_utils.read_json(json_file)
tc_result = {}
start_id = 539
end_id = 555
tc_list = tc_list[start_id-1:end_id]
start_filter = False

# Total TCs for Monitoring 
total_tcs = len(tc_list)
# current_tc = 0
pass_count = 0
fail_count = 0

# 4. Macro List 및 Validate 수행 
for idx, data in enumerate(tc_list, start=1):
    print(f"\n==================================== TC Progress : {idx} / {len(tc_list)} ====================================")
    summary = data['summary']
    print(f"Running TC : {summary}")
    # current_tc+=1 
    
    # 내보내기 or 파일로저장하는 경우는 사전에 다운로드 경로에 파일이 존재하지 않도록 삭제함 (2024-09-04, any문으로 변경)
    if any(x in summary for x in ['내보내기', '파일로 저장']):
        my_utils.delete_base_files()
    
    # `~추가` TC 수행할 때 DB초기화 하기 로직 추가 
    # if '객체추가' in summary:
    if any(x in summary for x in ['객체추가', '정책추가']): # 2024-10-07 수정
        db = Database()
        db.connect_db()
        db.initialize_tables()  # JENKINS로 돌릴때 주석처리 풀어줘야함. 개인TC생성 및 Play할때는 주석처리 해야할것 권장.
        my_utils.kill_process_by_pname(shared.PROCESS_NAME)
        print("새로운 객체생성을 위한 초기화 ...")
    
    macro_list = ast.literal_eval(data['macro list'])
    precondition = data['precondition']
    group1 = data['group1']
    group2 = data['group2']
    group3 = data['group3']
    group = "_".join(filter(lambda x: isinstance(x, str) and x, [group1, group2, group3]))
    print(group)

    # 4.1 Run Macro List
    performer.perform_macro(macro_list, precondition, group)
    
    # 4.2 Run Validate
    validator.type = data['val type']
    validator.summary = summary
    
    # 예외 처리를 추가하여 'val data'에서 발생할 수 있는 오류를 처리
    try:
        validator.data = ast.literal_eval(data['val data'])
    except (ValueError, SyntaxError) as e:
        # print(f"Error while evaluating 'val data': {e}")
        logger.log_exception(logger, e)

        # 기본값으로 None을 설정하거나 빈 값으로 대체할 수 있음
        validator.data = None

    if group3=='transparent' and '가동' in summary:
        time.sleep(8)
    match_result = validator.do_validate()
    
    # 팝업창이 뜨기 때문에 그 다음 TC수행하는데 지장이 없기위해 enter버튼 누른다. 
    # Todo - 필터링TC에서는 굳이 누를 필요 없다 (우선순위는 낮음 - 영향도 적다)
    if validator.type == 'image' :
        time.sleep(1)
        print("Enter Pressed")
        pyautogui.press('enter')
        pyautogui.press('enter')
    
    # pass/fail 기록     
    tc_id = data['id']
    if validator.type=='query':
        if not match_result['query_match_result']:
            tc_result[tc_id] = 0
        else:
            tc_result[tc_id] = 1
     
    elif validator.type=='image':
        if not match_result['image_match_result']:
            tc_result[tc_id] = 0
        else:
            tc_result[tc_id] = 1

    # 1. 로그인 TC(precondition=='non-login')는 팝업창을 띄우고 끝나므로 > 매니저 프로세스 Kill 
    # 2. 검증 실패(Fail)로 끝나는 경우 존재 > 매니저 프로세스 Kill 
    # if precondition == 'non-login' or tc_result[tc_id] == 0:
    #     my_utils.kill_process_by_pname(shared.PROCESS_NAME)
    if tc_result[tc_id] == 1:
        pass_count+=1
    else:
        fail_count+=1
    # Log the progress locally
    progress.log_progress_to_file(idx, total_tcs, pass_count, fail_count, tc_id)

    # Upload the log file every 1000 test cases
    # if current_tc % 1000 == 0:
    if idx % 1 == 0:
        progress.upload_log_file_via_sftp()

    print(f"TC Result (번호:성공(1)/실패(0)) : {tc_result}")


end_time = time.time()
print(end_time- start_time)

# 5. 결과값 마킹 - Pass/Fail에 대해서 Sheet1_Result 시트에 작성
my_utils.mark_result(tc_result, [start_time, end_time])
time.sleep(1)
# 6. 실패 리스트에 대해 JSON 생성
sheet.convert_fail_to_json()
