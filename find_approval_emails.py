#!/usr/bin/env python3
"""
查找需要批准的邮件并测试完整流程
"""

import sys
import os
sys.path.append('/Users/liqilong/Work/emailAuto')

from email_auto_approve import EmailAutoApprover

def find_approval_emails():
    """查找需要批准的邮件"""
    
    print("🔍 查找需要批准的邮件")
    print("=" * 60)
    
    approver = EmailAutoApprover()
    
    # 查找文件夹
    watch_folder = approver.config.get('DEFAULT', 'watch_folder')
    full_watch_path = approver.find_thunderbird_mail_folder(watch_folder)
    
    folder_name = watch_folder.split('/')[-1]
    mbox_file = os.path.join(full_watch_path, folder_name)
    
    print(f"📧 解析mbox文件: {mbox_file}")
    
    # 解析所有邮件
    emails = approver.parse_mbox_file(mbox_file)
    print(f"📊 总邮件数: {len(emails)}")
    
    # 查找需要批准的邮件
    approval_needed = []
    for email in emails:
        if approver.is_approval_needed(email):
            approval_needed.append(email)
    
    print(f"⚠️ 需要批准: {len(approval_needed)} 封")
    
    # 检查已处理的邮件
    processed_emails = approver.load_processed_emails()
    print(f"✅ 已处理过: {len(processed_emails)} 封")
    
    # 显示需要批准的邮件详情
    new_approvals = []
    for email in approval_needed:
        message_id = email.get('message_id', '')
        if not message_id:
            message_id = f"{email.get('subject', '')}|{email.get('from', '')}"
        
        if message_id not in processed_emails:
            new_approvals.append(email)
    
    print(f"🆕 新的批准请求: {len(new_approvals)} 封")
    
    if new_approvals:
        print(f"\n📋 待处理的批准邮件:")
        for i, email in enumerate(new_approvals[:5]):  # 只显示前5封
            print(f"  {i+1}. {email.get('subject', 'N/A')}")
            print(f"     发件人: {email.get('from', 'N/A')}")
            print(f"     日期: {email.get('date', 'N/A')}")
            
            # 检查正文中的Ref信息
            body = email.get('body', '')
            if body:
                import re
                msg_matches = re.findall(r'MSG\\d+', body)
                if msg_matches:
                    print(f"     Ref: {msg_matches[-1]}")
            print()
        
        if len(new_approvals) > 5:
            print(f"     ... 还有 {len(new_approvals) - 5} 封邮件")
        
        return new_approvals[0]  # 返回第一封待处理的邮件
    else:
        print("💡 没有新的批准请求需要处理")
        return None

if __name__ == "__main__":
    test_email = find_approval_emails()
    if test_email:
        print(f"\\n🎯 可以用这封邮件进行测试:")
        print(f"主题: {test_email.get('subject', 'N/A')}")
    else:
        print("\\n⚠️ 没有可测试的邮件")