#!/usr/bin/env python3
"""
直接测试本地邮件处理
"""

import sys
import os
sys.path.append('/Users/liqilong/Work/emailAuto')

from email_auto_approve import EmailAutoApprover

def test_local_email_processing():
    """测试本地邮件处理"""
    
    print("🎯 测试本地邮件处理")
    print("=" * 60)
    
    approver = EmailAutoApprover()
    
    # 直接使用绝对路径
    local_needapprove_path = "/Users/liqilong/Library/Thunderbird/Profiles/xoxled0d.default-release/Mail/Local Folders/Archives.sbd/ServiceNow.sbd/NeedApprove"
    
    print(f"📁 处理文件: {local_needapprove_path}")
    
    if os.path.exists(local_needapprove_path):
        size = os.path.getsize(local_needapprove_path)
        print(f"📄 文件大小: {size} bytes")
        
        if size > 0:
            print("📧 开始处理邮件...")
            result = approver.process_mbox_file(local_needapprove_path)
            
            if result:
                print("✅ 邮件处理成功!")
            else:
                print("❌ 邮件处理失败")
        else:
            print("📧 文件为空，没有邮件需要处理")
    else:
        print("❌ 文件不存在")

if __name__ == "__main__":
    test_local_email_processing()