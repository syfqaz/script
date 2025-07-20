import win32com.client
import os
import re
import base64

def clean_filename(filename):
    # 移除非法字符
    filename = re.sub(r'[<>:"/\\|?*]', '', filename)
    # 缩短文件名长度
    if len(filename) > 100:
        filename = filename[:100]
    return filename

# 创建一个Outlook应用实例
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# 获取指定账户
account = None
for acc in outlook.Folders:
    if acc.Name == "bosstao.th@gmail.com":
        account = acc
        break

if not account:
    raise Exception("未找到指定的账户")

# 获取账户下的"Inbox"文件夹
inbox = account.Folders["收件箱"]

# 指定导出邮件的保存路径
save_path = r"E:\self_python_script\下载outlook邮件"

# 确保路径存在
if not os.path.exists(save_path):
    os.makedirs(save_path)

# 遍历"Inbox"文件夹中的邮件
for i, message in enumerate(inbox.Items):
    try:
        # 生成邮件的保存文件名
        subject = message.Subject
        safe_subject = clean_filename(subject)
        filename = f"{i}_{safe_subject}.html"
        filepath = os.path.join(save_path, filename)

        # 将邮件内容以HTML格式保存到文件
        html_body = message.HTMLBody
        # 查找所有cid引用并替换为Base64编码的图片数据
        attachments = message.Attachments
        for attachment in attachments:
            cid = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E")
            if cid:
                attachment_file = attachment.FileName
                attachment.SaveAsFile(os.path.join(save_path, attachment_file))
                with open(os.path.join(save_path, attachment_file), "rb") as img_file:
                    base64_string = base64.b64encode(img_file.read()).decode('utf-8')
                html_body = html_body.replace(f"cid:{cid}", f"data:image/{os.path.splitext(attachment_file)[1][1:]};base64,{base64_string}")
                os.remove(os.path.join(save_path, attachment_file))  # 删除临时保存的文件

        with open(filepath, "w", encoding="utf-8") as file:
            file.write(html_body)
        print(f"邮件 '{subject}' 已保存到 {filepath}")
    except Exception as e:
        print(f"保存邮件时出错: {e}")

print("所有邮件已处理完毕。")
