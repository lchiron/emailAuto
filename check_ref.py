#!/usr/bin/env python3
"""
检查原始邮件内容中的Ref信息
"""

import sys
import os
sys.path.append('/Users/liqilong/Work/emailAuto')

from email_auto_approve import EmailAutoApprover

def check_email_content():
    """检查邮件内容中的Ref信息"""
    
    approver = EmailAutoApprover()
    
    # 查找mbox文件
    watch_folder = approver.config.get('DEFAULT', 'watch_folder')
    full_watch_path = approver.find_thunderbird_mail_folder(watch_folder)
    folder_name = watch_folder.split('/')[-1]
    mbox_file = os.path.join(full_watch_path, folder_name)
    
    # 解析邮件
    emails = approver.parse_mbox_file(mbox_file)
    
    print(f"找到 {len(emails)} 封邮件")
    
    for i, email_info in enumerate(emails):
        if approver.is_approval_needed(email_info):
            print(f"\n📧 邮件 {i+1}: {email_info.get('subject', 'N/A')}")
            print(f"发件人: {email_info.get('from', 'N/A')}")
            
            body = email_info.get('body', '')
            print(f"正文长度: {len(body)} 字符")
            
            # 显示正文内容的一部分来查看Ref格式
            print("正文内容:")
            print("-" * 50)
            print(body[:2000])  # 显示前2000个字符
            print("-" * 50)
            
            # 搜索Ref信息
            import re
            ref_patterns = [
                r'Ref:\s*([^\s\n\r]+)',  # Ref: 后面跟非空白字符
                r'Reference:\s*([^\s\n\r]+)',  # Reference: 后面跟非空白字符
                r'MSG\d+',  # MSG加数字
            ]
            
            print("\n🔍 搜索到的Ref信息:")
            for pattern in ref_patterns:
                matches = re.findall(pattern, body, re.IGNORECASE)
                if matches:
                    print(f"  模式 '{pattern}': {matches}")

if __name__ == "__main__":
    check_email_content()