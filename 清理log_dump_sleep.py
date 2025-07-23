import os
import shutil
import ctypes
import sys
import subprocess
import logging
import time

# 设置日志记录
logging.basicConfig(
    filename='clean_system_log.txt',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)


# 检查并安装 pywin32
def check_and_install_pywin32():
    try:
        import win32evtlog
        logging.info("pywin32 已安装。")
        print("pywin32 已安装。")
        return True
    except ImportError:
        logging.warning("未检测到 pywin32，正在尝试安装...")
        print("未检测到 pywin32，正在尝试安装...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
            logging.info("pywin32 安装成功！")
            print("pywin32 安装成功！")
            import win32evtlog  # 重新导入以确认安装成功
            return True
        except Exception as e:
            logging.error(f"安装 pywin32 失败：{e}")
            print(f"安装 pywin32 失败：{e}")
            print("请手动安装 pywin32：pip install pywin32")
            return False


# 确保脚本以管理员权限运行
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception as e:
        logging.error(f"检查管理员权限失败：{e}")
        print(f"检查管理员权限失败：{e}")
        return False


def run_as_admin():
    if not is_admin():
        logging.info("此脚本需要管理员权限，正在尝试以管理员身份重新运行...")
        print("此脚本需要管理员权限，正在尝试以管理员身份重新运行...")
        try:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            logging.info("已发起管理员权限提升请求。")
            sys.exit()
        except Exception as e:
            logging.error(f"尝试以管理员身份运行失败：{e}")
            print(f"尝试以管理员身份运行失败：{e}")
            print("请手动以管理员身份运行脚本！")
            time.sleep(5)  # 暂停5秒以便查看错误
            sys.exit(1)


# 清理指定目录
def clean_directory(directory):
    if not os.path.exists(directory):
        logging.info(f"目录 {directory} 不存在，跳过清理。")
        print(f"目录 {directory} 不存在，跳过清理。")
        return

    try:
        for item in os.listdir(directory):
            item_path = os.path.join(directory, item)
            try:
                if os.path.isfile(item_path) or os.path.islink(item_path):
                    os.unlink(item_path)
                    logging.info(f"已删除文件：{item_path}")
                    print(f"已删除文件：{item_path}")
                elif os.path.isdir(item_path):
                    shutil.rmtree(item_path)
                    logging.info(f"已删除目录：{item_path}")
                    print(f"已删除目录：{item_path}")
            except PermissionError as e:
                logging.error(f"权限不足，无法删除：{item_path} - {e}")
                print(f"权限不足，无法删除：{item_path}")
            except Exception as e:
                logging.error(f"删除 {item_path} 时出错：{e}")
                print(f"删除 {item_path} 时出错：{e}")
    except Exception as e:
        logging.error(f"访问目录 {directory} 时出错：{e}")
        print(f"访问目录 {directory} 时出错：{e}")


# 清理Windows事件日志
def clear_event_logs():
    try:
        import win32evtlog
    except ImportError:
        logging.error("无法导入 win32evtlog，跳过事件日志清理。")
        print("无法导入 win32evtlog，跳过事件日志清理。")
        return

    log_types = ["Application", "Security", "System"]  # Windows日志类别
    for log_type in log_types:
        try:
            handle = win32evtlog.OpenEventLog(None, log_type)
            win32evtlog.ClearEventLog(handle, None)
            win32evtlog.CloseEventLog(handle)
            logging.info(f"已清除 {log_type} 日志。")
            print(f"已清除 {log_type} 日志。")
        except Exception as e:
            logging.error(f"清除 {log_type} 日志时出错：{e}")
            print(f"清除 {log_type} 日志时出错：{e}")


def main():
    # 检查并请求管理员权限
    run_as_admin()

    # 检查并安装 pywin32
    if not check_and_install_pywin32():
        logging.warning("由于缺少 pywin32，事件日志清理功能将不可用，但将继续清理目录。")
        print("由于缺少 pywin32，事件日志清理功能将不可用，但将继续清理目录。")

    # 要清理的目录列表
    directories = [
        r"C:\Windows\System32\SleepStudy",
        r"C:\Windows\System32\SleepStudy\ScreenOn",
        r"C:\Windows\LiveKernelReports"
    ]

    # 清理目录
    for directory in directories:
        logging.info(f"开始清理目录：{directory}")
        print(f"开始清理目录：{directory}")
        clean_directory(directory)

    # 清理Windows事件日志
    logging.info("开始清理Windows事件日志...")
    print("\n开始清理Windows事件日志...")
    clear_event_logs()

    logging.info("清理完成！")
    print("\n清理完成！")
    print("日志已保存到 clean_system_log.txt")
    time.sleep(3)  # 暂停3秒以便查看结果


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logging.error(f"脚本执行出错：{e}")
        print(f"脚本执行出错：{e}")
        print("错误日志已保存到 clean_system_log.txt")
        time.sleep(5)  # 暂停5秒以便查看错误