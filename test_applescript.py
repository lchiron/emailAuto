#!/usr/bin/env python3
"""
Simple AppleScript test for Thunderbird email sending
"""

import subprocess
import time

def test_applescript():
    """测试AppleScript邮件发送功能"""
    
    # 简单的测试数据
    to_addr = "luluprod@service-now.com"
    subject = "Re: RITM1602185 - approve"  
    body_text = "Ref:MSG85395759"
    
    # 转义AppleScript字符串
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
    
    to_escaped = escape_applescript_string(to_addr)
    subject_escaped = escape_applescript_string(subject)
    body_escaped = escape_applescript_string(body_text)
    
    print(f"测试数据:")
    print(f"  收件人: {to_escaped}")
    print(f"  主题: {subject_escaped}")
    print(f"  正文: {body_escaped}")
    
    # 简化的AppleScript，增加更多调试信息
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
            
            -- 输入收件人 (应该默认在收件人字段)
            keystroke {to_escaped}
            delay 2
            
            -- 移到主题字段 (Tab键两次，跳过抄送字段)
            keystroke tab
            delay 1
            keystroke tab
            delay 1
            
            -- 输入主题
            keystroke {subject_escaped}
            delay 2
            
            -- 移到正文字段 (Tab键)
            keystroke tab
            delay 1
            
            -- 输入正文
            keystroke {body_escaped}
            delay 2
            
            -- 不自动发送，只是准备邮件
        end tell
    end tell
    '''
    
    print("\\n开始执行AppleScript...")
    try:
        result = subprocess.run(
            ['osascript', '-e', applescript],
            capture_output=True, text=True, check=True
        )
        print("✅ AppleScript执行成功")
        print("请检查Thunderbird中的邮件是否正确填写了各个字段")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ AppleScript执行失败:")
        print(f"Return code: {e.returncode}")
        print(f"Stdout: {e.stdout}")
        print(f"Stderr: {e.stderr}")
        return False

if __name__ == "__main__":
    test_applescript()