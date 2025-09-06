#!/usr/bin/env python3
"""
检查邮件识别逻辑
"""

import sys
import os
sys.path.append('/Users/liqilong/Work/emailAuto')

from email_auto_approve import EmailAutoApprover

def check_email_detection():
    """检查邮件识别逻辑"""
    
    print("🔍 检查邮件识别逻辑")
    print("=" * 60)
    
    approver = EmailAutoApprover()
    
    # 查找文件夹
    watch_folder = approver.config.get('DEFAULT', 'watch_folder')
    full_watch_path = approver.find_thunderbird_mail_folder(watch_folder)
    
    folder_name = watch_folder.split('/')[-1]
    mbox_file = os.path.join(full_watch_path, folder_name)
    
    # 解析所有邮件
    emails = approver.parse_mbox_file(mbox_file)
    
    print(f"📊 邮件列表 (前10封):")
    for i, email in enumerate(emails[:10]):
        subject = email.get('subject', 'N/A')
        from_addr = email.get('from', 'N/A')
        
        print(f"\\n  {i+1}. {subject}")
        print(f"     发件人: {from_addr}")
        
        # 测试批准检测逻辑
        needs_approval = approver.is_approval_needed(email)
        print(f"     需要批准: {'✅ 是' if needs_approval else '❌ 否'}")
        
        # 显示检测逻辑
        subject_lower = subject.lower()
        keywords = ['approve', 'approval', 'ritm', 'servicenow']
        matched_keywords = [kw for kw in keywords if kw in subject_lower]
        
        if matched_keywords:
            print(f"     匹配关键词: {matched_keywords}")
        else:
            print(f"     未匹配关键词: {keywords}")
    
    print(f"\\n💡 当前批准检测关键词: ['approve', 'approval', 'ritm', 'servicenow']")
    print(f"💡 如果需要调整检测逻辑，可以修改 is_approval_needed 函数")

if __name__ == "__main__":
    check_email_detection()