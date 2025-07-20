import wmi
import os
from datetime import datetime

def check_device_issues():
    # 初始化 WMI
    c = wmi.WMI()
    issues_found = []
    
    # 查询设备管理器中的所有设备
    for device in c.Win32_PnPEntity():
        # 检查设备状态是否有问题（黄色感叹号或未知设备）
        if device.Status != "OK" or device.ConfigManagerErrorCode != 0 or device.Name is None:
            device_info = f"设备名称: {device.Name or '未知设备'}, 状态: {device.Status or '未知'}, 错误代码: {device.ConfigManagerErrorCode or '无'}"
            issues_found.append(device_info)
    
    # 获取当前时间
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # 写入日志文件
    log_file = "设备YB.log"
    with open(log_file, "w", encoding="utf-8") as f:
        if issues_found:
            f.write(f"检查时间: {current_time}\n")
            f.write("检测到以下设备存在问题:\n")
            for issue in issues_found:
                f.write(f"- {issue}\n")
        else:
            f.write(f"检查时间: {current_time}\n")
            f.write("未检测到带有黄色感叹号或未知设备的问题。\n")
    
    print(f"检查完成，结果已保存到 {log_file}")

if __name__ == "__main__":
    check_device_issues()
