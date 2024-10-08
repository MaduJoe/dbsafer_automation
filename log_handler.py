import logging
import traceback

# 로그 설정 함수
def setup_logger():
    # 로그 파일 핸들러를 생성하고 UTF-8로 인코딩 설정
    file_handler = logging.FileHandler('error_log.log', encoding='utf-8')
    file_handler.setLevel(logging.ERROR)

    # 로그 포맷 설정
    formatter = logging.Formatter('%(asctime)s - %(message)s', datefmt='%Y_%m%d_%H:%M:%S')
    file_handler.setFormatter(formatter)

    # 로거 설정
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.ERROR)
    logger.addHandler(file_handler)
    return logger

# 예외를 받아서 로그를 남기는 함수
def log_exception(logger, e):
    error_message = traceback.format_exc()  # 에러 메시지 추적
    logger.error(f"An error occurred: {error_message}")  # 로그 파일에 에러 기록
