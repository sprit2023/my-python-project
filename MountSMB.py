import os
import subprocess
import time

def ping_host(host):
    """
    检测是否可以 Ping 通指定的主机
    """
    try:
        # 使用 ping 命令检测网络连通性
        output = subprocess.run(["ping", "-c", "1", host], capture_output=True, text=True, timeout=5)
        if output.returncode == 0:
            return True
        else:
            return False
    except Exception as e:
        print(f"Ping 失败: {e}")
        return False

def mount_smb_share(share_url):
    """
    挂载 SMB 共享文件夹
    """
    try:
        # 使用 osascript 调用 AppleScript 挂载 SMB 共享
        mount_script = f'tell application "Finder" to mount volume "{share_url}"'
        subprocess.run(["osascript", "-e", mount_script], check=True)
        
        # 获取挂载点名称并添加到侧边栏
        add_script = 'tell application "Finder" to make new Finder window to folder "Scan" of startup disk'
        subprocess.run(["osascript", "-e", add_script], check=True)
        print(f"成功挂载: {share_url}")
    except subprocess.CalledProcessError as e:
        print(f"挂载失败: {e}")

def main():
    host = "192.168.8.254"  # 要检测的主机
    share_url = "smb://192.168.8.254/Scan"  # SMB 共享文件夹地址

    while True:
        if ping_host(host):
            print(f"主机 {host} 可以 Ping 通，尝试挂载 SMB 共享...")
            mount_smb_share(share_url)
            break  # 挂载成功后退出循环
        else:
            print(f"主机 {host} 无法 Ping 通，等待 10 秒后重试...")
            time.sleep(10)  # 等待 10 秒后重试

if __name__ == "__main__":
    main()
