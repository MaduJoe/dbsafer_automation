
############    Validate Class 요약    ###########

# 1. row 읽어오기 (row 정보 : val_type에 따른 value)

# 2. val_type에 따라 검증 방법 및 검증값이 선택된다 

# 3. 결과값이 True / False 인지 받아서 GoogleSheet 에 Write한다

# val_result = False
# if val_type == 'query':
#     query = val 
#     val_result = send_query(query)     
# elif val_type == 'image':
#     query = val
#     val_result = find_image(query)
# write val_result is True/False/None

#################################################
import io
import os
import cv2
import time
import shared
import my_utils
import requests
import pyautogui
import numpy as np

from datetime import datetime
from database import Database
from my_utils import read_config
from googleapiclient.discovery import build
from google.oauth2.service_account import Credentials
from oauth2client.service_account import ServiceAccountCredentials
from PIL import Image
from io import BytesIO
from googleapiclient.http import MediaIoBaseDownload

def capture_screen(output_image_path):
    screenshot = pyautogui.screenshot(region=(155, 90, 1600, 850))
    screenshot.save(output_image_path)

def generate_unique_filename(base_path, img_name, img_type , extension=".png"):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    if img_type=='main':
        # 타임스탬프 기반으로 파일 이름 생성
        filename = f"{os.path.splitext(base_path)[0]}\\{img_name}_main_{timestamp}{extension}"

    else:
        filename = f"{os.path.splitext(base_path)[0]}\\{img_name}_output_{timestamp}{extension}"
        
    return filename

# def match_images_and_save(main_image_path, compare_image_path, output_image_path):
def match_images_and_save(base_image_path, compare_image_path, img_name, summary):
    # compare_image_path 가 존재하지 않으면 False반환 
    if not os.path.exists(compare_image_path):
        print(f"compare_image_path: {compare_image_path} 파일이 존재하지 않습니다. 드라이브에 업로드하세요.")
        return False

    # 고유한 파일 이름 생성
    main_image_path = generate_unique_filename(base_image_path, img_name, img_type='main')
    output_image_path = generate_unique_filename(base_image_path, img_name, img_type='output')

    # 현재 화면 캡처
    capture_screen(main_image_path)
    
    # 이미지 읽기
    main_image = cv2.imread(main_image_path)
    compare_image = cv2.imread(compare_image_path)

    # 그레이스케일로 변환
    main_gray = cv2.cvtColor(main_image, cv2.COLOR_BGR2GRAY)
    compare_gray = cv2.cvtColor(compare_image, cv2.COLOR_BGR2GRAY)

    # 템플릿 매칭
    result = cv2.matchTemplate(main_gray, compare_gray, cv2.TM_CCOEFF_NORMED)

    # 매칭 임계값 설정
    threshold = 0.85
    loc = np.where(result >= threshold)

    # 매칭된 사각형의 개수를 카운트
    summary = summary.replace(' ','')
    match_count = 0
    for pt in zip(*loc[::-1]):
        match_count += 1
        # match_found = True
        # 매칭된 영역에 사각형 그리기
        cv2.rectangle(main_image, pt, (pt[0] + compare_gray.shape[1], pt[1] + compare_gray.shape[0]), (0, 255, 0), 2)

    # 결과 이미지 저장
    cv2.imwrite(output_image_path, main_image)
    # return match_found

    # 필터링 TC수행할 때는 2개이상의 매칭이 필요하다.
    if '탭>필터링' in summary:
        return match_count >= 2
    else:
        return match_count

class LoadImage:
    def __init__(self) -> None:
        my_utils.print_message("Load Image by URL")
        data = my_utils.read_config(shared.CONFIG)
        credintial_file = data['google_sheet']['secret_key']
        scope = data['google_sheet']['sheet_scope']
        creds = ServiceAccountCredentials.from_json_keyfile_name(credintial_file, scope)
        self.drive_service = build('drive', 'v3', credentials=creds)
        
    # 특정 폴더에서 파일 검색
    def search_file_in_folder(self, folder_id, file_name):
        query = f"'{folder_id}' in parents and name contains '{file_name}' and mimeType contains 'image/'"
        results = self.drive_service.files().list(q=query, fields="files(id, name)").execute()
        files = results.get('files', [])
        
        if files:
            print(f"Found {len(files)} file(s) matching '{file_name}':")
            for file in files:
                print(f"File Name: {file['name']}, File ID: {file['id']}")
            return files
        else:
            print("No files found")
            return None

    # 다운로드 URL을 통해 이미지 로드
    def load_image_from_url(self, file_id):
        # Google Drive에서 파일의 다운로드 URL 가져오기
        request = self.drive_service.files().get(fileId=file_id, fields='webContentLink').execute()
        download_url = request.get('webContentLink')

        if download_url:
            response = requests.get(download_url)
            image = Image.open(BytesIO(response.content))
            return image
        else:
            raise Exception("Unable to get download URL for the file.")

    # 다운로드 URL을 통해 이미지 로드
    def load_image_as_cv2(self, file_id):
        # Google Drive에서 파일의 실제 데이터를 다운로드
        request = self.drive_service.files().get_media(fileId=file_id)
        fh = BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()

        fh.seek(0)
        image = Image.open(fh)
        image_np = np.array(image)
        # OpenCV에서는 BGR 포맷을 사용하므로, 필요시 RGB에서 BGR로 변환
        if image_np.shape[2] == 3:  # 컬러 이미지인 경우
            image_cv2 = cv2.cvtColor(image_np, cv2.COLOR_RGB2BGR)
        else:  # 흑백 이미지인 경우
            image_cv2 = image_np
        return image_cv2

             
        
    def list_all_files_in_folder(self, folder_id):
        query = f"'{folder_id}' in parents"
        results = self.drive_service.files().list(q=query, fields="files(id, name, mimeType)").execute()
        files = results.get('files', [])
        
        if files:
            print(f"Found {len(files)} file(s) in the folder:")
            for file in files:
                print(f"File Name: {file['name']}, File ID: {file['id']}, MIME Type: {file['mimeType']}")
            return files
        else:
            print("No files found in the folder.")
            return None

class Validator:
    def __init__(self) -> None:
        self.db = Database()
        self.db.connect_db()
        # self.db.initialize_tables()
        self.type = None
        self.data = None
        self.summary = None
    
    def do_validate(self):
        my_utils.print_message("Start Macro Validate")
        
        match_result = None
        
        # valid-type에 따라 검증 step 수행
        if self.type == 'query':
            sql = self.data
            print(f"query stmt: {sql}")
            query_result = self.db.execute_query(query=sql['query'])
            match_result = {
                'query_match_result': query_result,
            }

        elif self.type == 'image':
            
            gd = GoogleDriveDownloader()
            img_name = self.data['image_name']
            compare_image_path = os.path.join(gd.target_dir, f"{img_name}.png")  # 사전에 저장한 이미지
            # main_image_path = os.path.join(gd.target_dir, f"{img_name}_main.png")  # 현재 화면 캡처 이미지
            # output_image_path = os.path.join(gd.target_dir, f"{img_name}_match.png")  # 매칭 결과 이미지
              
            if not os.path.exists(compare_image_path):
                print(f"Error: Comparison image does not exist at path: {compare_image_path}")
                # TO-DO ; log로 남기기
                gd.download_all_images()
                image_match_result = match_images_and_save(gd.target_dir, compare_image_path, img_name, self.summary)
            else:
                image_match_result = match_images_and_save(gd.target_dir, compare_image_path, img_name, self.summary)

            match_result = {
                'image_match_result': image_match_result,
            }
        else:
            print(f"Validate Type이 `query` 이거나 `image`가 아닙니다.")
            
        print(f"Finished Validate")
        return match_result 

class GoogleDriveDownloader:
    def __init__(self):
        data = read_config(shared.CONFIG)
        google_config = data['google_sheet']
        
        self.credentials_json = google_config['secret_key']
        
        self.product_version = google_config['manager_version']
        
        self.scopes = ['https://www.googleapis.com/auth/drive.readonly']
        
        self.drive_service = self.authenticate()
        
        self.target_dir = google_config['img_target_mgr7'] if self.product_version==7 else google_config['img_target_mgr5']
        
        self.images_id = google_config['img_source_mgr7'] if self.product_version==7 else google_config['img_source_mgr5']
        
        
        # print(self.images_id)

    def authenticate(self):
        creds = Credentials.from_service_account_file(self.credentials_json, scopes=self.scopes)
        return build('drive', 'v3', credentials=creds)


    # # ver 1 
    # def list_images_in_folder(self):
    #     query = f"'{self.images_id}' in parents and mimeType contains 'image/'"
    #     results = self.drive_service.files().list(q=query, fields="files(id, name)").execute()
    #     files = results.get('files', [])
    #     print(files)
    #     return files
    
    # ver 2 
    def list_images_in_folder(self, folder_id=None):
        if folder_id is None:
            folder_id = self.images_id
            print(f"folder_id : {folder_id}")

        # 이미지가 들어있는 폴더를 찾는 쿼리
        query = f"'{folder_id}' in parents and (mimeType = 'application/vnd.google-apps.folder' or mimeType contains 'image/')"
        results = self.drive_service.files().list(q=query, fields="files(id, name, mimeType)").execute()
        files = results.get('files', [])

        images = []
        
        for file in files:
            if file['mimeType'] == 'application/vnd.google-apps.folder':
                # 하위 폴더가 발견되면 재귀적으로 해당 폴더를 탐색
                images.extend(self.list_images_in_folder(file['id']))
            elif 'image/' in file['mimeType']:
                images.append({'id': file['id'], 'name': file['name']})
        
        return images
    

    def download_image(self, file_id, file_name):
        request = self.drive_service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()

        fh.seek(0)
        image = Image.open(fh)
        return image

    def save_image(self, image, file_name):
        target_path = os.path.join(self.target_dir, file_name)
        image.save(target_path)
        print(f"Image saved to {target_path}")

    def download_all_images(self):
        if not os.path.exists(self.target_dir):
            os.makedirs(self.target_dir)

        files = self.list_images_in_folder()
        if not files:
            print("No images found in the folder.")
            return

        for file in files:
            file_id = file['id']
            file_name = file['name']
            print(f"Downloading {file_name}...")

            image = self.download_image(file_id, file_name)
            self.save_image(image, file_name)
        time.sleep(3)
            
                 
if __name__=='__main__':
    # import ast
    # import macro_performer
    
    # # tc_list 가져오기 위해 사용
    # macro = macro_performer.MacroPerformer()
    # repo_url = "https://raw.githubusercontent.com/MaduJoe/gui_automation/main/manager_20240819_145656.json"
    # tc_list = macro.fetch_macro_list(repo_url)
    
    # val = Validator()
    # for data in tc_list[1:]:    
    #     print("\n================== Run & Validate ==================")
    #     # Run Macro 
    #     macro_list = ast.literal_eval(data['macro_list'])
    #     macro.perform_macro(macro_list)
    #     # Validate 
    #     val.type = data['val_type']
    #     val.data = ast.literal_eval(data['val_data'])
    #     match_result = val.do_validate()
    #     print(f"{match_result}")
    
    
    ###############################################
    
    # li = LoadImage()
    # folder_id = "13ZI_y4V-1TGI0F1UygU8Z5egvalNHoPr"
    # file_name = "fail_login"
    # # 폴더에서 파일 검색
    # found_files = li.search_file_in_folder(folder_id, file_name)
    # if found_files:
    #     # 첫 번째 파일 로드 및 표시
    #     first_file = found_files[0]
    #     image_cv2 = li.load_image_as_cv2(first_file['id'])
        
    #     # OpenCV를 사용해 이미지 조작 및 저장
    #     cv2.imshow("Loaded Image", image_cv2)
    #     cv2.waitKey(0)
    #     cv2.destroyAllWindows()
    # else:
    #     print("No matching images found.")
    
    gd = GoogleDriveDownloader()
    # all_imgs = gd.list_images_in_folder()
    gd.download_all_images()
    pass