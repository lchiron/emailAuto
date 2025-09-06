#!/usr/bin/env python3
"""
Outlook PWA邮件发送工具

通过AppleScript控制Outlook PWA自动发送邮件(macOS)
支持Safari或Chrome中的Outlook PWA应用
"""

import os
import sys
import subprocess
import platform
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time


class OutlookPWASender:
    def __init__(self, browser="Safari"):
        self.system = platform.system()
        self.browser = browser  # "Safari" 或 "Google Chrome"
        
    def is_browser_running(self):
        """检查浏览器是否正在运行"""
        if self.system == "Darwin":
            try:
                result = subprocess.run(
                    ['osascript', '-e', f'tell application "System Events" to (name of processes) contains "{self.browser}"'],
                    capture_output=True, text=True, check=True
                )
                return result.stdout.strip() == "true"
            except:
                return False
        return False
    
    def launch_browser(self):
        """启动浏览器"""
        if self.system == "Darwin":
            try:
                subprocess.run(['open', '-a', self.browser], check=True)
                return True
            except:
                return False
        return False
    
    def open_outlook_pwa(self):
        """打开Outlook PWA（如果尚未打开）"""
        if self.system != "Darwin":
            raise NotImplementedError("此功能只支持macOS")
        
        # 确保浏览器正在运行
        if not self.is_browser_running():
            if not self.launch_browser():
                raise Exception(f"无法启动{self.browser}")
            time.sleep(3)
        
        # 检查是否已经打开了Outlook PWA页面
        try:
            # 使用AppleScript检查当前标签页是否是Outlook
            check_script = f'''
            tell application "{self.browser}"
                activate
                set currentURL to URL of active tab of front window
                return currentURL
            end tell
            '''
            
            result = subprocess.run(['osascript', '-e', check_script], 
                                  capture_output=True, text=True, check=True)
            current_url = result.stdout.strip()
            
            # 如果当前页面不是Outlook，尝试查找已打开的Outlook标签
            if "outlook.office.com" not in current_url:
                find_outlook_script = f'''
                tell application "{self.browser}"
                    set outlookTabFound to false
                    repeat with w in windows
                        repeat with t in tabs of w
                            if URL of t contains "outlook.office.com" then
                                set active tab index of w to index of t
                                set index of w to 1
                                set outlookTabFound to true
                                exit repeat
                            end if
                        end repeat
                        if outlookTabFound then exit repeat
                    end repeat
                    
                    if not outlookTabFound then
                        -- 如果没找到Outlook标签，在当前标签中打开
                        set URL of active tab of front window to "https://outlook.office.com/mail/"
                    end if
                end tell
                '''
                
                subprocess.run(['osascript', '-e', find_outlook_script], 
                             capture_output=True, text=True, check=True)
                time.sleep(5)  # 等待页面加载
            
            return True
            
        except Exception as e:
            print(f"检查或打开Outlook PWA失败: {e}")
            # 备用方案：在新标签中打开
            try:
                subprocess.run(['open', '-a', self.browser, "https://outlook.office.com/mail/"], check=True)
                time.sleep(5)
                return True
            except Exception as e2:
                print(f"备用方案也失败: {e2}")
                return False
    
    def send_email_via_outlook_pwa(self, to_addr, subject, body_text, auto_send=True):
        """通过Outlook PWA发送邮件"""
        if self.system != "Darwin":
            raise NotImplementedError("AppleScript只支持macOS")
        
        # 确保Outlook PWA已打开
        if not self.open_outlook_pwa():
            raise Exception("无法打开Outlook PWA")
        
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
        
        # 转义AppleScript字符串
        def escape_applescript_string(s):
            if s is None:
                return '""'
            s = str(s).replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
            import re
            s = re.sub(r'\s+', ' ', s).strip()
            s = s.replace('\\', '\\\\').replace('"', '\\"')
            return '"' + s + '"'
        
        to_escaped = escape_applescript_string(clean_to_addr)
        subject_escaped = escape_applescript_string(subject)
        body_escaped = escape_applescript_string(body_text)
        
        # 根据auto_send参数决定是否包含发送命令
        send_command = '''
                -- 发送邮件 (Ctrl+Enter 或点击发送按钮)
                keystroke return using {command down}
                delay 3
                ''' if auto_send else '''
                -- 不自动发送，留待手动发送'''
        
        # Outlook PWA AppleScript - 专注于在当前页面操作
        applescript = f'''
        tell application "{self.browser}"
            activate
            delay 2
        end tell
        
        tell application "System Events"
            tell process "{self.browser}"
                -- 确保焦点在Outlook页面
                delay 2
                
                -- 尝试多种方法点击"New email"按钮
                set emailCreated to false
                
                -- 方法1: 尝试使用快捷键 Cmd+N (在Outlook PWA中通常有效)
                try
                    keystroke "n" using {{command down}}
                    delay 5
                    set emailCreated to true
                end try
                
                -- 方法2: 如果快捷键失败，尝试查找按钮
                if not emailCreated then
                    try
                        -- 尝试点击包含"New"的按钮
                        click button "New email"
                        delay 5
                        set emailCreated to true
                    on error
                        try
                            -- 尝试其他可能的按钮名称
                            click button "New message"
                            delay 5
                            set emailCreated to true
                        on error
                            try
                                click button "Compose"
                                delay 5
                                set emailCreated to true
                            end try
                        end try
                    end try
                end if
                
                -- 方法3: 如果按钮方法也失败，尝试点击可能的位置
                if not emailCreated then
                    try
                        -- 尝试点击左侧导航区域的新建邮件位置
                        click at {{150, 150}}
                        delay 3
                        set emailCreated to true
                    end try
                end if
                
                -- 如果所有自动方法都失败，给用户提示
                if not emailCreated then
                    display dialog "请手动点击'New email'或'新建邮件'按钮，然后点击确定继续" buttons {{"确定"}} default button "确定"
                    delay 2
                end if
                
                -- 等待新邮件编辑窗口完全加载
                delay 3
                
                -- 开始填写邮件内容
                -- 输入收件人
                keystroke {to_escaped}
                delay 2
                
                -- 移到主题字段
                keystroke tab
                delay 1
                
                -- 输入主题
                keystroke {subject_escaped}
                delay 2
                
                -- 移到正文字段
                keystroke tab
                delay 1
                
                -- 输入正文内容
                keystroke {body_escaped}
                delay 2{send_command}
                
                -- 完成
                delay 2
            end tell
        end tell
        '''
        
        print(f"{'自动发送' if auto_send else '准备'}邮件 (Outlook PWA):")
        print(f"  浏览器: {self.browser}")
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
        """创建邮件草稿文件 (备用方案)"""
        if not filename:
            import time
            filename = f"outlook_draft_{int(time.time())}.eml"
        
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
    
    def send_email(self, to_addr, subject, body_text, body_html=None, auto_send=True):
        """发送邮件 - 主接口"""
        try:
            # 尝试通过Outlook PWA发送
            if self.system == "Darwin":
                return self.send_email_via_outlook_pwa(to_addr, subject, body_text, auto_send)
            else:
                # 其他系统创建草稿文件
                filename = self.create_draft_file(to_addr, subject, body_text, body_html)
                print(f"已创建邮件草稿: {filename}")
                print("请手动使用Outlook发送此邮件")
                return True
                
        except Exception as e:
            print(f"通过Outlook PWA发送邮件失败: {e}")
            # 备用方案：创建草稿文件
            try:
                filename = self.create_draft_file(to_addr, subject, body_text, body_html)
                print(f"已创建邮件草稿: {filename}")
                print("请手动使用Outlook发送此邮件")
                return True
            except Exception as e2:
                print(f"创建草稿文件也失败: {e2}")
                return False


def test_outlook_pwa_sender():
    """测试Outlook PWA邮件发送功能"""
    
    # 可以选择浏览器：Safari 或 Google Chrome
    browser_choice = input("选择浏览器 (1: Safari, 2: Chrome): ").strip()
    
    if browser_choice == "2":
        browser = "Google Chrome"
    else:
        browser = "Safari"
    
    sender = OutlookPWASender(browser=browser)
    
    to_addr = "luluprod@service-now.com"
    subject = "Re: RITM1602185 - approve"
    body_text = "Ref:MSG85395759"
    
    print("测试Outlook PWA邮件发送...")
    print(f"使用浏览器: {browser}")
    
    # 测试模式选择
    mode_choice = input("发送模式 (1: 自动发送, 2: 只准备): ").strip()
    auto_send = (mode_choice == "1")
    
    result = sender.send_email(to_addr, subject, body_text, auto_send=auto_send)
    
    if result:
        print("✅ 邮件处理成功！")
    else:
        print("❌ 邮件处理失败！")


if __name__ == "__main__":
    test_outlook_pwa_sender()