#!/usr/bin/env python3
"""
检查NeedApprove文件夹的实际内容
"""

import sys
import os
sys.path.append('/Users/liqilong/Work/emailAuto')

from email_auto_approve import EmailAutoApprover

def check_needapprove_folder():
    """检查NeedApprove文件夹内容"""
    
    print("🔍 检查NeedApprove文件夹内容")
    print("=" * 60)
    
    # 直接检查文件路径
    needapprove_path = "/Users/liqilong/Library/Thunderbird/Profiles/xoxled0d.default-release/Mail/Local Folders/Archives.sbd/ServiceNow.sbd"
    
    print(f"📂 ServiceNow.sbd目录: {needapprove_path}")
    
    if os.path.exists(needapprove_path):
        print("✅ 目录存在")
        items = os.listdir(needapprove_path)
        print(f"📋 目录内容: {items}")
        
        # 检查NeedApprove文件
        needapprove_file = os.path.join(needapprove_path, "NeedApprove")
        if os.path.exists(needapprove_file):
            size = os.path.getsize(needapprove_file)
            print(f"📄 NeedApprove文件: {size} bytes")
            
            if size > 0:
                print("📧 文件有内容，让程序解析:")
                
                approver = EmailAutoApprover()
                try:
                    emails = approver.parse_mbox_file(needapprove_file)
                    print(f"解析到 {len(emails)} 封邮件")
                    
                    for i, email in enumerate(emails):
                        print(f"  {i+1}. {email.get('subject', 'N/A')}")
                        print(f"     发件人: {email.get('from', 'N/A')}")
                        print(f"     需要批准: {'是' if approver.is_approval_needed(email) else '否'}")
                except Exception as e:
                    print(f"解析失败: {e}")
            else:
                print("📧 文件为空")
        else:
            print("❌ NeedApprove文件不存在")
        
        # 检查.msf文件内容
        msf_file = os.path.join(needapprove_path, "NeedApprove.msf")
        if os.path.exists(msf_file):
            print(f"\n📄 .msf文件存在，大小: {os.path.getsize(msf_file)} bytes")
            print("💡 .msf文件包含邮件索引信息，但邮件内容在mbox文件中")
    else:
        print("❌ 目录不存在")
    
    # 测试程序的文件夹查找功能
    print(f"\n🔍 测试程序查找功能:")
    approver = EmailAutoApprover()
    watch_folder = approver.config.get('DEFAULT', 'watch_folder')
    print(f"配置路径: {watch_folder}")
    
    found_path = approver.find_thunderbird_mail_folder(watch_folder)
    if found_path:
        print(f"✅ 程序找到: {found_path}")
    else:
        print("❌ 程序找不到路径")

if __name__ == "__main__":
    check_needapprove_folder()