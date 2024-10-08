import time
import psutil
import shared
import my_utils
import win32gui
from pywinauto import Application
import ctypes

class AppStarter:
    def __init__(self):
        my_utils.print_message("Start Application")
        data = my_utils.read_config(shared.CONFIG)
        self.pname = data['start']['process_name'] # 0 : sa_manager / 1: pc_assist / 2: enterprise_manager | 각 VM PC에는 고정값 1개로 설정할것이다 
        self.username = data['start']['user_name']
        self.userpw = data['start']['user_pw']
        self.app = None
        
        # 3개의 app중에서 PC Assist는 x86이므로 win32으로 컨트롤 
        if self.pname == 'DBSaferAgt.exe':
            self.backend = 'win32'
        else:
            self.backend = 'uia'

    def run_as_admin(self, exe_path):
        try:
            # 관리자 권한으로 실행
            ctypes.windll.shell32.ShellExecuteW(None, "runas", exe_path, None, None, 1)
        except Exception as e:
            print(f"Failed to run as admin: {e}")
            
    def check_existing_process(self):
        # global PID
        process_status = None
        
        # 프로세스명으로 PID찾아서 
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] == self.pname:
                    self.app = Application(backend=self.backend).connect(process=proc.info['pid'])
                    hwnd = my_utils.get_hwnd_from_pid(proc.info['pid'])
                    if hwnd is not None and win32gui.IsWindow(hwnd):
                        print(f"Existing process {proc.info['name']} with PID {proc.info['pid']}")
                        shared.PID = proc.info['pid']
                        process_status = True
                        time.sleep(1) # 기존 3에서 1로 변경 ; 2024-08-28 16:23:00
                        
                    else:
                        print(f"Process {proc.info['name']} with PID {proc.info['pid']} does not have an active window.")
                        process_status = False

            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
                print(f"Could not Find process {proc.info['name']} with PID {proc.info['pid']}: {e}")
            
        if process_status==None:
            print("프로세스가 존재하지 않습니다.")
            
        return process_status

    def start_app(self, pname):
        start_success = False
        try:
            # app = Application(backend=self.backend)
            app = Application(backend='uia')
            if pname == 'DBSaferAgt.exe':
                # app.start(r"C:\Program Files (x86)\PNPSECURE\DBSAFER AGENT\dsastlch.exe")
                self.run_as_admin(r"C:\Program Files (x86)\PNPSECURE\DBSAFER AGENT\dsastlch.exe")

            elif pname == 'nodesafermgr.exe':
                # app.start(r"C:\Program Files\PNP SECURE\NODESAFER\nodesafermgr.exe")
                self.run_as_admin(r"C:\Program Files\PNP SECURE\NODESAFER\nodesafermgr.exe")
                

            elif pname == 'Enterprise Manager.exe':
                # app.start(r"C:\Program Files\PNP SECURE\Enterprise Manager 7\Enterprise Manager.exe")
                self.run_as_admin(r"C:\Program Files\PNP SECURE\Enterprise Manager 7\Enterprise Manager.exe")
                
                
            else:
                raise Exception(f"Process Name이 정확한지 확인하세요: {pname}")
            print(f"{pname}을 실행합니다")
            time.sleep(3)
            
            shared.PID = my_utils.get_process_id()
            self.app = app.connect(process=shared.PID)
            
            start_success = True
            return start_success

        except FileNotFoundError as e:
            print(f"파일을 찾을 수 없습니다: {e}")

        except Exception as e:
            print(f"오류가 발생했습니다: {e}")
            
            
if __name__=='__main__':
    app = AppStarter()
    app.start_app(app.pname)