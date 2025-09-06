#!/usr/bin/env python3
"""
Quick test of email processing without file monitoring
"""

import sys
import os
sys.path.append('/Users/liqilong/Work/emailAuto')

from email_auto_approve import EmailAutoApprover

def test_email_processing():
    """测试邮件处理功能"""
    
    approver = EmailAutoApprover()
    
    # 查找mbox文件
    watch_folder = approver.config.get('DEFAULT', 'watch_folder')
    full_watch_path = approver.find_thunderbird_mail_folder(watch_folder)
    
    if not full_watch_path:
        print("❌ 找不到Thunderbird邮件文件夹")
        return False
    
    print(f"✅ 找到目标目录: {full_watch_path}")
    
    # 获取mbox文件路径
    folder_name = watch_folder.split('/')[-1]
    mbox_file = os.path.join(full_watch_path, folder_name)
    
    print(f"📧 mbox文件路径: {mbox_file}")
    
    if not os.path.exists(mbox_file):
        print("❌ mbox文件不存在")
        return False
    
    print(f"📊 文件大小: {os.path.getsize(mbox_file)} 字节")
    
    # 处理mbox文件
    print("\\n🔄 开始处理邮件...")
    result = approver.process_mbox_file(mbox_file)
    
    if result:
        print("✅ 邮件处理完成")
    else:
        print("⚠️ 没有发现需要处理的邮件")
    
    return result

if __name__ == "__main__":
    test_email_processing()