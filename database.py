import shared
import pymysql
import my_utils
from pymysql import MySQLError
import log_handler

class Database:
    def __init__(self) -> None:
        self.connection = None
        self.config = my_utils.read_config(shared.CONFIG)
        self.pgm = True
        # self.table = None   # table은 list형태 일 수 있다. 

        self.logger = log_handler.setup_logger()
    
    # DB 연결 
    def connect_db(self):
        # print("\n================== Connect DB Stat ==================")
        if self.connection == None:
            db_config = self.config['db_connection']
            self.connection = pymysql.connect(
                host = db_config['db_ip'],
                port = db_config['db_port'],
                user = db_config['db_user'],
                password = db_config['db_password'],
                database = db_config['db_name']
            )
            print(f"Success to Connect {db_config['db_name']} !")

        else:
            # my_utils.print_message("기존에 연결된 DB Connection이 존재합니다.")
            print(f"Existing Pre-connected Session: {db_config['db_name']} !")

    # 쿼리 실행
    def execute_query(self, query):
        try:
            with self.connection.cursor() as cursor:
                cursor.execute(query)
                self.connection.commit()    # 커밋을 명시적으로 수행
                result = cursor.fetchall()
                return result
        except MySQLError as e:
            # TO-DO ; 로깅 남기고 아예 프로세스 Kill  >> 결과값이 안남기 때문에 
            print(f"쿼리 실행 에러: {e}")
            log_handler.log_exception(self.logger, e)
            
    # 테이블 초기화
    def initialize_tables(self):
        words = ['policy','host','service','sid', 'ip', 'id', 'app', 'ldap', 'cert', 'datamask', 'table', 'formal', 'time', 'cmd', 'event', 'alert']

        conditions = " OR ".join([f"table_name LIKE '%{word}%'" for word in words])

        query = f"""
        SELECT table_name 
        FROM information_schema.tables 
        WHERE table_schema = 'dbsafer3' 
        AND ({conditions});
        """

        table_list = self.execute_query(query)

        for table in table_list:
            table = table[0]
            try:
                
                # # 테이블의 모든 컬럼명 가져오기
                # query = f"SHOW COLUMNS FROM `{table}`"
                # columns = self.execute_query(query)

                # name, group_name, object_name 컬럼만 선택
                query = f"""
                SELECT COLUMN_NAME
                FROM INFORMATION_SCHEMA.COLUMNS
                WHERE TABLE_SCHEMA = 'dbsafer3'
                AND TABLE_NAME = '{table}'
                AND COLUMN_NAME IN ('name', 'group_name', 'object_name', 'comment', 'title');
                """
                columns = self.execute_query(query)
                
                # 모든 컬럼명에 대해 처리
                for col in columns:
                    col_name = col[0]
                    
                    # 각 컬럼에서 'test'가 포함된 데이터를 삭제하는 쿼리
                    delete_query = f"DELETE FROM `{table}` WHERE `{col_name}` LIKE '%test%'"

                    # 삭제 쿼리 실행
                    with self.connection.cursor() as cursor:
                        cursor.execute(delete_query)
                        self.connection.commit()
                        
                        # 삭제된 행(row) 수 확인
                        if cursor.rowcount > 0:
                            print(f"Rows containing 'test' in column '{col_name}' of table '{table}' have been deleted.")
                    
                    # # 삭제 쿼리 실행
                    # self.execute_query(delete_query)
                    # print(f"Rows containing 'test' in column '{col_name}' of table '{table}' have been deleted.")
                    
            except MySQLError as e:
                print(f"Error processing table {table}: {e}")

if __name__=='__main__':
    db = Database()
    db.connect_db()
    db.initialize_tables()


    # table_list = [
    #     'policy_db_access', 'policy_db_auth', 'policy_db_global', 
    #     'policy_ftp_access', 'policy_ftp_auth', 
    #     'policy_shell_access', 'policy_shell_auth', 
    #     'client_host', 
    #     'services', 'servicegroups',
    #     'sid', 'sid_group',
    #     'client_ip', 'client_ip_group',
    #     'client_id', 'client_id_group',
    #     'client_app', 'client_app_group',
    #     'dbsafer_ldap_list', 
    #     'datamask', 'datamask_group',
    #     'table', 'table_group'
    #     'formal_query', 'formal_query_group',
    #     'time', 
    #     'client_cmd', 'client_cmd_group',
    #     'client_event_group',
    # ]

