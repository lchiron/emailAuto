#!/usr/bin/env python3
"""
Debug mbox parsing to see what emails are found
"""

import sys
import os
sys.path.append('/Users/liqilong/Work/emailAuto')

from email_auto_approve import EmailAutoApprover

def debug_mbox_parsing():
    """调试mbox解析功能"""
    
    approver = EmailAutoApprover()
    
    # 查找mbox文件
    watch_folder = approver.config.get('DEFAULT', 'watch_folder')
    full_watch_path = approver.find_thunderbird_mail_folder(watch_folder)
    folder_name = watch_folder.split('/')[-1]
    mbox_file = os.path.join(full_watch_path, folder_name)
    
    print(f"📧 解析mbox文件: {mbox_file}")
    
    # 解析邮件
    emails = approver.parse_mbox_file(mbox_file)
    
    print(f"\\n📊 解析结果: 找到 {len(emails)} 封邮件")
    
    for i, email_info in enumerate(emails):
        print(f"\\n📧 邮件 {i+1}:")
        print(f"  主题: {email_info.get('subject', 'N/A')}")
        print(f"  发件人: {email_info.get('from', 'N/A')}")
        print(f"  收件人: {email_info.get('to', 'N/A')}")
        print(f"  日期: {email_info.get('date', 'N/A')}")
        
        # 检查是否需要批准
        needs_approval = approver.is_approval_needed(email_info)
        print(f"  需要批准: {'✅ 是' if needs_approval else '❌ 否'}")
        
        # 显示正文的一部分
        body = email_info.get('body', '')
        if body:
            preview = body[:100].replace('\\n', ' ').replace('\\r', ' ').strip()
            print(f"  正文预览: {preview}...")
        
        # 如果需要批准，检查是否已处理过
        if needs_approval:
            processed_emails = approver.load_processed_emails()
            message_id = email_info.get('message_id', '')
            if not message_id:
                message_id = f"{email_info.get('subject', '')}|{email_info.get('from', '')}"
            
            already_processed = message_id in processed_emails
            print(f"  已处理过: {'✅ 是' if already_processed else '❌ 否'}")

if __name__ == "__main__":
    debug_mbox_parsing()