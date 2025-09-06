#!/usr/bin/env python3
"""
专门测试AppleScript主题字段传递
"""

import subprocess
import time

def test_subject_passing():
    """测试主题字段传递给AppleScript"""
    
    # 测试数据 - 确保包含Re:前缀
    to_addr = "luluprod@service-now.com"
    subject = "Re: RITM1234567 - approve"  # 明确包含Re:前缀
    body_text = "Ref:MSG12345678"
    
    # 转义AppleScript字符串
    def escape_applescript_string(s):
        if s is None:
            return '""'
        s = str(s).replace('\n', ' ').replace('\r', ' ').replace('\t', ' ')
        import re
        s = re.sub(r'\s+', ' ', s).strip()
        s = s.replace('\\', '\\\\').replace('"', '\\"')
        return '"' + s + '"'
    
    to_escaped = escape_applescript_string(to_addr)
    subject_escaped = escape_applescript_string(subject)
    body_escaped = escape_applescript_string(body_text)
    
    print(f"🧪 测试AppleScript主题传递:")
    print(f"  原始主题: {subject}")
    print(f"  转义后主题: {subject_escaped}")
    print(f"  收件人: {to_escaped}")
    print(f"  正文: {body_escaped}")
    
    # AppleScript - 只准备邮件，不发送，以便检查
    applescript = f'''
    tell application "Thunderbird"
        activate
        delay 3
    end tell
    
    tell application "System Events"
        tell process "Thunderbird"
            -- 创建新邮件
            keystroke "m" using {{command down, shift down}}
            delay 5
            
            -- 输入收件人
            keystroke {to_escaped}
            delay 2
            
            -- 移到主题字段 (Tab x 2)
            keystroke tab
            delay 1
            keystroke tab
            delay 1
            
            -- 输入主题 (这是关键测试点)
            keystroke {subject_escaped}
            delay 2
            
            -- 移到正文字段
            keystroke tab
            delay 1
            
            -- 输入正文
            keystroke {body_escaped}
            delay 2
            
            -- 不发送，只是准备邮件供检查
        end tell
    end tell
    '''
    
    print("\\n🚀 执行AppleScript...")
    try:
        result = subprocess.run(
            ['osascript', '-e', applescript],
            capture_output=True, text=True, check=True
        )
        print("✅ AppleScript执行成功")
        print("📧 请检查Thunderbird中新建的邮件:")
        print(f"   - 主题字段是否显示: {subject}")
        print(f"   - 收件人字段是否显示: {to_addr}")  
        print(f"   - 正文字段是否显示: {body_text}")
        print("⚠️  如果主题缺少'Re:'前缀，这可能是AppleScript字符转义问题")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ AppleScript执行失败:")
        print(f"错误信息: {e.stderr}")
        return False

if __name__ == "__main__":
    test_subject_passing()