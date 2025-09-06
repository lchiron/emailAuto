#!/usr/bin/env python3
"""
Thunderbird邮件发送工具

通过AppleScript控制Thunderbird自动发送邮件(macOS)
或通过其他方法发送邮件
"""

import os
import sys
import subprocess
import platform
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


class ThunderbirdSender:
    def __init__(self):
        self.system = platform.system()
        
    def is_thunderbird_running(self):
        """检查Thunderbird是否正在运行"""
        if self.system == "Darwin":
            try:
                result = subprocess.run(
                    ['osascript', '-e', 'tell application "System Events" to (name of processes) contains "Thunderbird"'],
                    capture_output=True, text=True, check=True
                )
                return result.stdout.strip() == "true"
            except:
                return False
        return False
    
    def launch_thunderbird(self):
        """启动Thunderbird"""
        if self.system == "Darwin":
            try:
                subprocess.run(['open', '-a', 'Thunderbird'], check=True)
                return True
            except:
                return False
        return False
    
    def send_email_via_applescript(self, to_addr, subject, body_text, body_html=None):
        """通过AppleScript发送邮件 (macOS)"""
        if self.system != "Darwin":
            raise NotImplementedError("AppleScript只支持macOS")
        
        # 确保Thunderbird正在运行
        if not self.is_thunderbird_running():
            if not self.launch_thunderbird():
                raise Exception("无法启动Thunderbird")
            # 等待启动
            import time
            time.sleep(3)
        
        # 清理收件人地址 - 只保留邮箱地址
        clean_to_addr = to_addr
        if '<' in to_addr and '>' in to_addr:
            # 提取尖括号中的邮箱地址
            start = to_addr.find('<') + 1
            end = to_addr.find('>')
            clean_to_addr = to_addr[start:end]
        elif '"' in to_addr:
            # 移除引号
            clean_to_addr = to_addr.replace('"', '').strip()
            if '<' in clean_to_addr and '>' in clean_to_addr:
                start = clean_to_addr.find('<') + 1
                end = clean_to_addr.find('>')
                clean_to_addr = clean_to_addr[start:end]
        
        # 转义AppleScript字符串
        def escape_applescript_string(s):
            if s is None:
                return '""'
            # 只转义必要的字符
            s = str(s).replace('\\', '\\\\').replace('"', '\\"')
            return '"' + s + '"'
        
        to_escaped = escape_applescript_string(clean_to_addr)
        subject_escaped = escape_applescript_string(subject)
        body_escaped = escape_applescript_string(body_text)
        
        applescript = f'''
        tell application "Thunderbird"
            activate
            delay 1
        end tell
        
        tell application "System Events"
            tell process "Thunderbird"
                -- 创建新邮件 (Cmd+Shift+M)
                keystroke "m" using {{command down, shift down}}
                delay 3
                
                -- 确保焦点在收件人字段
                -- 填写收件人
                keystroke {to_escaped}
                delay 0.5
                
                -- 移到主题字段 (Tab x 2)
                keystroke tab
                delay 0.3
                keystroke tab  
                delay 0.3
                
                -- 填写主题
                keystroke {subject_escaped}
                delay 0.5
                
                -- 移到正文字段 (Tab)
                keystroke tab
                delay 0.3
                
                -- 填写正文
                keystroke {body_escaped}
                delay 1
                
                -- 发送邮件 (Cmd+Return 或 Cmd+Shift+Return)
                keystroke return using {{command down}}
                delay 0.5
            end tell
        end tell
        '''
        
        print(f"发送邮件:")
        print(f"  收件人: {clean_to_addr}")
        print(f"  主题: {subject}")
        print(f"  正文长度: {len(body_text)} 字符")
        
    def send_email_via_applescript_with_confirmation(self, to_addr, subject, body_text, auto_send=True):
        """通过AppleScript发送邮件，可选择是否自动发送"""
        if self.system != "Darwin":
            raise NotImplementedError("AppleScript只支持macOS")
        
        # 确保Thunderbird正在运行
        if not self.is_thunderbird_running():
            if not self.launch_thunderbird():
                raise Exception("无法启动Thunderbird")
            import time
            time.sleep(3)
        
        # 清理收件人地址
        clean_to_addr = to_addr
        if '<' in to_addr and '>' in to_addr:
            start = to_addr.find('<') + 1
            end = to_addr.find('>')
            clean_to_addr = to_addr[start:end]
        elif '"' in to_addr:
            clean_to_addr = to_addr.replace('"', '').strip()
            if '<' in clean_to_addr and '>' in clean_to_addr:
                start = clean_to_addr.find('<') + 1
                end = clean_to_addr.find('>')
                clean_to_addr = clean_to_addr[start:end]
        
        # 转义AppleScript字符串 - 修复版本
        def escape_applescript_string(s):
            if s is None:
                return '""'
            # 移除所有换行符、回车符和制表符
            s = str(s).replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
            # 合并多个空格为单个空格
            import re
            s = re.sub(r'\s+', ' ', s).strip()
            # 转义双引号和反斜杠
            s = s.replace('\\', '\\\\').replace('"', '\\"')
            return '"' + s + '"'
        
        to_escaped = escape_applescript_string(clean_to_addr)
        subject_escaped = escape_applescript_string(subject)
        body_escaped = escape_applescript_string(body_text)
        
        # 根据auto_send参数决定是否包含发送命令
        send_command = '''
                -- 发送邮件 (Cmd+Return)
                keystroke return using {command down}
                delay 3
                
                -- 额外等待确保发送完成
                delay 2''' if auto_send else '''
                -- 不自动发送，留待手动发送'''
        
        applescript = f'''
        tell application "Thunderbird"
            activate
            delay 3
        end tell
        
        tell application "System Events"
            tell process "Thunderbird"
                -- 创建新邮件 (Cmd+Shift+M)
                keystroke "m" using {{command down, shift down}}
                delay 5
                
                -- 确保光标在收件人字段，输入邮箱地址
                keystroke {to_escaped}
                delay 2
                
                -- 移到主题字段 (Tab x 2，跳过抄送字段)
                keystroke tab
                delay 1
                keystroke tab  
                delay 1
                
                -- 输入主题
                keystroke {subject_escaped}
                delay 2
                
                -- 移到正文字段 (Tab)
                keystroke tab
                delay 1
                
                -- 输入正文内容
                keystroke {body_escaped}
                delay 2{send_command}
                
                -- 额外等待确保邮件完全处理完成
                delay 3
            end tell
        end tell
        '''
        
        print(f"{'自动发送' if auto_send else '准备'}邮件:")
        print(f"  收件人: {clean_to_addr}")
        print(f"  主题: {subject}")
        print(f"  正文: '{body_text}'")
        if not auto_send:
            print("  ⚠️  需要手动点击发送按钮")
        
        try:
            result = subprocess.run(
                ['osascript', '-e', applescript],
                capture_output=True, text=True, check=True
            )
            print("✅ AppleScript执行成功")
            return True
        except subprocess.CalledProcessError as e:
            print(f"❌ AppleScript执行失败: {e.stderr}")
            raise Exception(f"AppleScript执行失败: {e.stderr}")
    
    def create_draft_file(self, to_addr, subject, body_text, body_html=None, filename=None):
        """创建邮件草稿文件"""
        if not filename:
            import time
            filename = f"draft_{int(time.time())}.eml"
        
        msg = MIMEMultipart('alternative')
        msg['To'] = to_addr
        msg['Subject'] = subject
        msg['From'] = "Edward Li <eli23@lululemon.com>"
        
        from email.utils import formatdate, make_msgid
        msg['Date'] = formatdate(localtime=True)
        msg['Message-ID'] = make_msgid()
        
        # 添加纯文本部分
        text_part = MIMEText(body_text, 'plain', 'utf-8')
        msg.attach(text_part)
        
        # 添加HTML部分（如果提供）
        if body_html:
            html_part = MIMEText(body_html, 'html', 'utf-8')
            msg.attach(html_part)
        
        # 保存到文件
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(msg.as_string())
        
        return filename
    
    def send_email(self, to_addr, subject, body_text, body_html=None):
        """发送邮件 - 主接口"""
        try:
            # 优先尝试AppleScript方法 (macOS)
            if self.system == "Darwin":
                return self.send_email_via_applescript(to_addr, subject, body_text, body_html)
            else:
                # 其他系统创建草稿文件
                filename = self.create_draft_file(to_addr, subject, body_text, body_html)
                print(f"已创建邮件草稿: {filename}")
                print("请手动使用Thunderbird发送此邮件")
                return True
                
        except Exception as e:
            print(f"发送邮件失败: {e}")
            # 备用方案：创建草稿文件
            try:
                filename = self.create_draft_file(to_addr, subject, body_text, body_html)
                print(f"已创建邮件草稿: {filename}")
                print("请手动使用Thunderbird发送此邮件")
                return True
            except Exception as e2:
                print(f"创建草稿文件也失败: {e2}")
                return False


def test_email_sender():
    """测试邮件发送功能"""
    sender = ThunderbirdSender()
    
    to_addr = "luluprod@service-now.com"
    subject = "Re: RITM1602185 - approve"
    body_text = """Ref:MSG85395759


Edward Li | China Devops
eli23@lululemon.com
lululemon Store Support Centre
18F lululemon, Tower2, No. 3 Hongqiao Road
Xu hui District, Shanghai City, China
lululemon.com"""
    
    body_html = """<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
</head>
<body>
Ref:MSG85395759
<div>
<br>
<b>Edward Li | China Devops</b><br>
<b>eli23@lululemon.com</b><br>
<b>lululemon Store Support Centre</b><br>
<b>18F lululemon, Tower2, No. 3 Hongqiao Road</b><br>
<b>Xu hui District, Shanghai City, China</b><br>
<b><a href="http://www.lululemon.com/">lululemon.com</a></b>
</div>
</body>
</html>"""
    
    print("测试邮件发送...")
    result = sender.send_email(to_addr, subject, body_text, body_html)
    
    if result:
        print("邮件发送成功！")
    else:
        print("邮件发送失败！")


if __name__ == "__main__":
    test_email_sender()