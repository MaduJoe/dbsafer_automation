import paramiko
from datetime import datetime
import my_utils
import shared

def test():
    config = my_utils.read_config(shared.CONFIG)
    db_config = config['db_connection']
    db_ip = db_config['db_ip']
    print(db_ip)


# Function to upload the log file via SFTP
def upload_log_file_via_sftp():
    try:

        config = my_utils.read_config(shared.CONFIG)
        db_config = config['db_connection']
        db_ip = db_config['db_ip']

        # Set up the SFTP connection
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        # ssh.connect(hostname='192.168.106.71', username='root', password='dbsafer00')
        ssh.connect(hostname=db_ip, username='root', password='dbsafer00')

        sftp = ssh.open_sftp()

        # Upload the log file
        src_log_file = 'progress_log.txt'
        remote_log_file = f'/etc/log_progress/{src_log_file}'

        sftp.put(src_log_file, remote_log_file)

        sftp.close()
        ssh.close()

        print("Log file uploaded successfully to remote server via SFTP")
    
    except Exception as e:
        print(f"Error uploading log file: {e}")


# Example usage
def log_progress_to_file(current_tc, total_tcs, pass_count, fail_count, tc_id):
    progress = (current_tc / total_tcs) * 100
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')  # Generate current timestamp
    with open('progress_log.txt', 'a') as f:
        # f.write(f"Progress: {progress:.2f}% - Passed: {pass_count}, Failed: {fail_count}\n")
        if current_tc==1:
            f.write(f"\n===========================================================================\n")
        f.write(f"[{timestamp}] Progress: {progress:.2f}% [{current_tc}/{total_tcs}] - Passed: {pass_count}, Failed: {fail_count}, TC No: {tc_id} \n")


if __name__=='__main__':
    test()