# -*- coding: utf-8 -*-
import sys
import time
import my_utils
import shared
import requests
import app_starter
import pywinauto.mouse as mo

from time import sleep
from pywinauto.timings import wait_until
from pynput import keyboard as pynput_keyboard
from pynput import mouse as pynput_mouse
from database import Database

from pywinauto.application import Application

class MacroPerformer:
    def __init__(self):
        self.keyboard = pynput_keyboard.Controller()
        self.mouse_listener = None
        self.keyboard_listener = None
        self.play_mode = False
        self.ctrl_c_pressed = False
        self.ctrl_key = None
        
        # db = Database()
        # db.connect_db()
        # db.initialize_tables()  # JENKINS로 돌릴때 주석처리 풀어줘야함. 개인TC생성 및 Play할때는 주석처리 해야할것 권장.
    
    # 주석처리 - 로그인시 필터링 다 고정해두고 돌리는 걸로 수정 예정
    # def fix_filtering(self, group):
    #     print('fix fitering 함수 시작')
    #     pre_text, post_text = group.split('_')[0], group.split('_')[-1]
    #     print(f"pre_text : {pre_text}")
    #     print(f"post_text : {post_text}")
    #     if pre_text == 'object':
    #         if post_text == 'dbms':
    #             coord_list = [(13, 137), (169, 106), (379, 222), (580, 138)]
    #             # for coord in coord_list:
    #             #     mo.click(coords=coord, button='left')
    #             #     time.sleep(0.3)
    #         elif post_text == 'instance':
    #             coord_list = [(13, 137), (179, 119), (382, 147), (584, 135)]

    #         for coord in coord_list:
    #             mo.click(coords=coord, button='left')
    #             time.sleep(0.3)

    def on_click(self, x, y, button, pressed):
        if self.play_mode:
            return False  # Block mouse events

    def on_move(self, x, y):
        if self.play_mode:
            return False  # Block mouse events

    def on_press(self, key):
        try:
            # Ctrl+C 감지
            if key == pynput_keyboard.Key.ctrl_l or key == pynput_keyboard.Key.ctrl_r:
                self.ctrl_c_pressed = True
                self.ctrl_key = key
                
            elif key == pynput_keyboard.KeyCode.from_char('c') and self.ctrl_c_pressed:
                print("Ctrl+C pressed. Exiting the program.")
                raise KeyboardInterrupt  # 직접 예외를 발생시켜 종료합니다.
                
            if not self.play_mode:
                print(f'alphanumeric key {key.char} pressed')
                
        except AttributeError:
            print(f'special key {key} pressed')

    def on_release(self, key):
        # Ctrl 키 해제 시 상태 초기화
        if key == pynput_keyboard.Key.ctrl_l or key == pynput_keyboard.Key.ctrl_r:
            self.ctrl_c_pressed = False
            self.ctrl_key = None
              
        if not self.play_mode:
            print(f'{key} released')

    def start_listeners(self):
        # Mouse listener
        if self.mouse_listener is None or not self.mouse_listener.is_alive():
            self.mouse_listener = pynput_mouse.Listener(on_click=self.on_click, on_move=self.on_move)
            self.mouse_listener.start()

        # Keyboard listener
        if self.keyboard_listener is None or not self.keyboard_listener.is_alive():
            self.keyboard_listener = pynput_keyboard.Listener(on_press=self.on_press, on_release=self.on_release)
            self.keyboard_listener.start()

    def stop_listeners(self):
        if self.mouse_listener and self.mouse_listener.is_alive():
            self.mouse_listener.stop()
            self.mouse_listener.join()

        if self.keyboard_listener and self.keyboard_listener.is_alive():
            self.keyboard_listener.stop()
            self.keyboard_listener.join()
            
    # def perform_macro(self, event_list, precondition, group, summary=None):
    def perform_macro(self, event_list, precondition, group):
        self.play_mode = True
        self.start_listeners()
        
        try:
            # start application 
            app = app_starter.AppStarter()
            
            # login 관련 TC수행시 
            if precondition=='non-login':
                my_utils.kill_process_by_pname(shared.PROCESS_NAME)
                app.start_app(shared.PROCESS_NAME)
                
            # 기타 TC수행시                 
            elif precondition=='login':
                exists = app.check_existing_process()   # 여기서 shared.PID 할당
                group_depth = group.split('_')
                group1, group2 = group_depth[0], group_depth[1]
                if exists:
                    my_utils.maximize_window()
                    time.sleep(1)   # 옆의 Panel이 펼쳐지기 위한 대기시간

                    if group1 == 'object':
                        my_utils.release_services()
                    
                    
                else:
                    app.start_app(shared.PROCESS_NAME)  # 여기서 shared.PID 할당
                    my_utils.login_manager(shared.PID)
                    my_utils.maximize_window()
                    if group1 == 'object':
                        my_utils.release_services()
                        
                # 로그인 and 서비스관련 TC일 경우 서비스 하위 항목 펼치는 로직 필요
                # 임시 주석처리 - 빠르게 TC 테스트 하기 위해 
                my_utils.click_into_group(group)

            
            time.sleep(1)
            my_utils.print_message("Start Macro Peform")
            
            last_event_type = None
            
            for idx, event in enumerate(event_list, start=1):
                my_utils.print_message(f"{idx} / {len(event_list)}")
                # my_utils.focus_window()
                
                event_type, value = event[0], event[1:]
                if value == '' or value == None:
                    continue
                
                if event_type == 'click':
                    ctrl_name, ctrl_type, click_cnt, click_type, coordinates, search_by, wait_time = value[0]
                    if idx==1:
                        # sleep(2) 
                        sleep(1) 
                    else:
                        sleep(wait_time)
                    
                    # button 빼고 거의다 좌표로 찾는다 - 중복되는 요소들이 존재하기 때문에 좌표로 찾아야한다.
                    if group1=='object' and group2=='service':
                        if ctrl_type == 'Pane' or ctrl_type == 'Edit' or ctrl_type == 'Group' or ctrl_type == 'List' or ctrl_type == 'Text' or ctrl_type == 'Custom' or ctrl_type == 'TreeItem':
                            search_by = 'coord'

                    else:
                        # 서비스 객체 제외하고 Edit은 id값으로 찾기
                        if ctrl_type == 'Pane' or ctrl_type == 'Group' or ctrl_type == 'List' or ctrl_type == 'Text' or ctrl_type == 'Custom' or ctrl_type == 'TreeItem':
                            search_by = 'coord'

                    rect = None
                    if search_by == 'id':
                        try:
                            # 현재 창에서 요소를 찾는다. 2024-09-03 수정
                            element = shared.PARENT_WINDOW.child_window(auto_id=ctrl_name, control_type=ctrl_type)  # auto_id
                            if element.exists(timeout=3):
                                rect = element.rectangle()
                                print(f"id로 찾음 - {ctrl_name}")
                                
                            else:
                                raise Exception(f"Element with ID '{ctrl_name}' not found.")
                            
                        except Exception as e:
                            print(f"[ 예외처리 - ID ]: {e}")
                            rect = my_utils.click_with_coordinates(parent_window=shared.PARENT_WINDOW, ctrl_id=ctrl_name, ctrl_type=ctrl_type, coords=coordinates)

                    elif search_by == 'name':
                        try:
                            # 현재 창에서 요소를 찾는다. 2024-09-03 수정
                            element = shared.PARENT_WINDOW.child_window(title=ctrl_name, control_type=ctrl_type)    # title
                            if element.exists(timeout=3):
                                rect = element.rectangle()
                                print(f"name으로 찾음 - {ctrl_name}")
                            else:
                                raise Exception(f"Element with Name '{ctrl_name}' not found.")
                            
                        except Exception as e:
                            print(f"[ 예외처리 - Name ]: {e}")
                            rect = my_utils.click_with_coordinates(parent_window=shared.PARENT_WINDOW, ctrl_name=ctrl_name, ctrl_type=ctrl_type, coords=coordinates)

                    
                    # 예외처리로 받은 좌표
                    if rect == None:
                        center_x, center_y = coordinates[0], coordinates[1]
                        
                    # 요소처리로 받은 좌표 
                    else:
                        center_x, center_y = rect.left + (rect.right - rect.left) // 2, rect.top + (rect.bottom - rect.top) // 2

                    if click_cnt == 'double':
                        mo.double_click(coords=(center_x, center_y), button=click_type)
                        print(f"Double Clicked ({center_x}, {center_y}, {click_type}) Success")

                    else:
                        mo.click(coords=(center_x, center_y), button=click_type)
                        print(f"Clicked ({center_x}, {center_y}, {click_type}) Success")

                # 키보드 - 버튼내림
                elif event_type == 'keydown':
                    key = value[0]
                    wait_time = value[1]
                    
                    if key != 'f4' and last_event_type == 'click':  # 20240910 추가 
                        sleep(2)
                    else:
                        sleep(wait_time)

                    try:
                        if hasattr(pynput_keyboard.Key, key):
                            special_key = getattr(pynput_keyboard.Key, key)
                            self.keyboard.press(special_key)
                            print(f'Pressed S-Key Down: {key}')
                        else:
                            self.keyboard.press(key)
                            print(f'Pressed Key Down: {key}')
                    except ValueError as e:
                        print(f"Error pressing key: {key}. Exception: {e}")
                        
                # 키보드 - 버튼올림
                elif event_type == 'keyup':
                    key = value[0]
                    wait_time = value[1]
                    if key != 'f4':
                        sleep(0.1)
                    else:
                        sleep(wait_time)
                    
                    try:
                        if hasattr(pynput_keyboard.Key, key):
                            special_key = getattr(pynput_keyboard.Key, key)
                            self.keyboard.release(special_key)
                            print(f'Released S-Key Up: {key}')
                        else:
                            self.keyboard.release(key)
                            print(f'Released Key Up: {key}')
                    except ValueError as e:
                        print(f"Error releasing key: {key}. Exception: {e}")
                
                last_event_type = event_type        
                sleep(0.1)  # event 간에 interval - 20240910 0.1 -> 0.3으로 변경 
            sleep(2)    # 모든 event 끝나고 n초간 pause (loading같은 팝업창이 뜰 수 있기 떄문에)
            
        except KeyboardInterrupt:
            print("Macro execution interrupted by Ctrl+C key.") # ctrl + a 눌렀을 때 이 구문 탄다...
        
        finally:
            # if precondition == 'non-login':
            #     my_utils.kill_process_by_pname(shared.PROCESS_NAME)
            self.play_mode = False
            self.stop_listeners()
            print("Listeners stopped.")
            
if __name__=='__main__':
    performer = MacroPerformer()
    events = [('click', ('59648', 'Pane', 'single', 'left', (14, 136), 'id', 2.4)), ('click', ('서비스', 'TreeItem', 'single', 'left', (109, 110), 'name', 2.4)), ('click', ('서비스', 'TreeItem', 'single', 'left', (115, 107), 'name', 3.7)), ('click', ('DBMS', 'TreeItem', 'single', 'right', (180, 121), 'name', 1.3)), ('click', ('그룹 생성', 'MenuItem', 'single', 'left', (210, 133), 'name', 0.8)), ('keydown', 't', 0.8), ('keyup', 't', 0.1), ('keydown', 'e', 0.1), ('keyup', 'e', 0.1), ('keydown', 's', 0.0), ('keyup', 's', 0.1), ('keydown', 't', 0.0), ('keyup', 't', 0.1), ('keydown', 'enter', 0.8), ('keyup', 'enter', 0.1), ('click', ('DBMS', 'TreeItem', 'single', 'left', (167, 126), 'name', 3.2)), ('click', ('test123', 'Text', 'single', 'left', (574, 156), 'name', 3.2)), ('click', ('test123', 'Text', 'single', 'right', (581, 154), 'name', 1.4)), ('keydown', 'esc', 7.2), ('keyup', 'esc', 0.1), ('click', ('그룹에 등록', 'Button', 'single', 'left', (744, 116), 'name', 1.2)), ('click', ('test123', 'Text', 'single', 'left', (776, 397), 'name', 1.6)), ('click', ('4247', 'Button', 'single', 'left', (954, 447), 'id', 0.8)), ('click', ('test', 'Text', 'single', 'left', (777, 689), 'name', 0.9)), ('click', ('8285', 'Button', 'single', 'left', (967, 737), 'id', 0.5)), ('click', ('2419', 'Button', 'single', 'left', (1177, 866), 'id', 0.8)), ('click', ('적용', 'Button', 'single', 'left', (385, 106), 'name', 1.6)), ('keydown', 'f4', 1.9), ('keyup', 'f4', 0.1)]
    # events = [('click', ('59648', 'Pane', 'single', 'left', (15, 131), 'id', 13.0)), ('click', ('서비스', 'TreeItem', 'single', 'left', (174, 106), 'name', 1.5)), ('click', ('객체 추가', 'Button', 'single', 'left', (590, 108), 'name', 2.7)), ('keydown', 't', 1.4), ('keyup', 't', 0.1), ('keydown', 'e', 0.1), ('keyup', 'e', 0.1), ('keydown', 's', 0.2), ('keyup', 's', 0.1), ('keydown', 't', 0.1), ('keyup', 't', 0.1), ('keydown', '1', 0.4), ('keyup', '1', 0.1), ('click', ('7159', 'Group', 'single', 'left', (879, 500), 'id', 2.5)), ('click', ('7303', 'Group', 'single', 'left', (883, 555), 'id', 1.1)), ('keydown', '1', 1.0), ('keyup', '1', 0.1), ('keydown', '.', 0.0), ('keyup', '.', 0.1), ('keydown', '2', 0.1), ('keyup', '2', 0.1), ('keydown', '.', 0.0), ('keyup', '.', 0.1), ('keydown', '3', 0.1), ('keyup', '3', 0.1), ('keydown', '.', 0.1), ('keyup', '.', 0.1), ('keydown', '4', 0.1), ('keyup', '4', 0.1), ('keydown', 'tab', 0.3), ('keyup', 'tab', 0.1), ('keydown', 'tab', 0.1), ('keyup', 'tab', 0.1), ('keydown', '1', 2.0), ('keyup', '1', 0.1), ('keydown', '2', 0.1), ('keyup', '2', 0.0), ('keydown', '3', 0.1), ('keyup', '3', 0.1), ('keydown', '4', 0.1), ('keyup', '4', 0.1), ('click', ('4238', 'Edit', 'single', 'left', (879, 715), 'id', 1.0)), ('keydown', 't', 0.5), ('keyup', 't', 0.1), ('keydown', 'e', 0.1), ('keyup', 'e', 0.1), ('keydown', 's', 0.1), ('keyup', 's', 0.1), ('keydown', 't', 0.0), ('keyup', 't', 0.1), ('click', ('6111', 'ComboBox', 'single', 'left', (954, 784), 'id', 2.3)), ('click', ('사용자 - DBSAFER 간 암호화', 'ListItem', 'single', 'left', (954, 813), 'name', 1.1)), ('click', ('9667', 'ComboBox', 'single', 'left', (948, 839), 'id', 1.1)), ('click', ('사용 (로그 수집 방식)', 'ListItem', 'single', 'left', (941, 868), 'name', 0.9)), ('click', ('21135', 'ComboBox', 'single', 'left', (939, 889), 'id', 0.6)), ('click', ('사용', 'ListItem', 'single', 'left', (933, 925), 'name', 0.8)), ('click', ('2419', 'Button', 'single', 'left', (993, 922), 'id', 1.8)), ('click', ('적용', 'Button', 'single', 'left', (380, 98), 'name', 2.0)), ('keydown', 'f4', 1.8), ('keyup', 'f4', 0.1)]
    # events = [('click', ('59648', 'Pane', 'single', 'left', (9, 138), 'id', 2.4)), ('click', ('DBMS', 'TreeItem', 'single', 'right', (187, 126), 'name', 1.9)), ('click', ('그룹 생성', 'MenuItem', 'single', 'left', (221, 136), 'name', 0.9)), ('keydown', 't', 0.5), ('keyup', 't', 0.1), ('keydown', 'e', 0.0), ('keydown', 's', 0.1), ('keyup', 's', 0.0), ('keyup', 'e', 0.0), ('keydown', 't', 0.4), ('keyup', 't', 0.1), ('keydown', 'enter', 0.4), ('keyup', 'enter', 0.1), ('click', ('DBMS', 'TreeItem', 'single', 'left', (179, 124), 'name', 2.4)), ('click', ('그룹에 등록', 'Button', 'single', 'left', (733, 113), 'name', 2.7)), ('click', ('test1', 'Text', 'single', 'left', (751, 400), 'name', 1.4)), ('click', ('4246', 'Button', 'single', 'left', (958, 416), 'id', 0.7)), ('click', ('test', 'Text', 'single', 'left', (733, 683), 'name', 0.9)), ('click', ('8938', 'Button', 'single', 'left', (966, 707), 'id', 0.8)), ('click', ('2419', 'Button', 'single', 'left', (1202, 870), 'id', 0.8)), ('click', ('적용', 'Button', 'single', 'left', (386, 114), 'name', 2.5)), ('keydown', 'f4', 1.5), ('keyup', 'f4', 0.0)]
    group = "policy_dbms_global"
    group = "policy_ftp_auth"
    group = "policy_terminal_rdp"
    group = "object_service_dbms"
    # group = "object_service_ftp"
    # group = "object_service_terminal"
    # group = "object_service_transparent"
    performer.perform_macro(event_list=events, precondition='login', group=group)
    
#     import ast 
#     # GitHub에서 Macro 리스트 가져오기
#     # TO-DO - json 절대 경로가 아닌 dynamic하게 가제올수 있게끔 repo_url 가져오게끔 변경하기 
    
#     macro = MacroPerformer()
#     repo_url = "https://raw.githubusercontent.com/MaduJoe/gui_automation/main/manager_20240819_145656.json"
#     tc_list = macro.fetch_macro_list(repo_url)  # macro 관련 함수니까 MacroPerformer 클래스에 생성함 
    
#     for idx, data in enumerate(tc_list, start=1):
#         print(f"\n================== {idx} / {len(tc_list)} ==================")
#         macro_list = ast.literal_eval(data['macro_list'])
#         precondition = data['precondition']
#         macro.perform_macro(macro_list, precondition)



    # app = app_starter.AppStarter()
    # exists = app.check_existing_process()
    # if exists:
    #     window_utils.maximize_window()
    # else:
    #     app.start_app(shared.PROCESS_NAME)
        
    # print(app.pname)