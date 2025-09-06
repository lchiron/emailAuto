#!/usr/bin/env python3
"""
验证IMAP ServiceNow文件夹并测试移动功能
"""

import sys
import os
sys.path.append('/Users/liqilong/Work/emailAuto')

from email_auto_approve import EmailAutoApprover

def verify_servicenow_folder():
    """验证ServiceNow IMAP文件夹"""
    
    print("🔍 验证ServiceNow IMAP文件夹")
    print("=" * 60)
    
    approver = EmailAutoApprover()
    
    # 检查指定的路径
    target_path = "/Users/liqilong/Library/thunderbird/Profiles/xoxled0d.default-release/webaccountMail/outlook.office365.com/ServiceNow.sbd"
    
    print(f"📂 目标路径: {target_path}")
    
    if not os.path.exists(target_path):
        print("❌ 路径不存在")
        return False
    
    print("✅ 路径存在")
    
    # 列出目录内容
    print(f"\n📋 目录内容:")
    try:
        items = os.listdir(target_path)
        for item in items:
            item_path = os.path.join(target_path, item)
            if os.path.isfile(item_path):
                size = os.path.getsize(item_path)
                print(f"  📄 {item} ({size:,} bytes)")
            elif os.path.isdir(item_path):
                print(f"  📁 {item}/")
    except Exception as e:
        print(f"❌ 无法列出目录内容: {e}")
        return False
    
    # 测试程序的文件夹查找功能
    print(f"\n🔍 测试程序查找功能:")
    watch_folder = approver.config.get('DEFAULT', 'watch_folder')
    print(f"配置的watch_folder: {watch_folder}")
    
    found_path = approver.find_thunderbird_mail_folder(watch_folder)
    if found_path:
        print(f"✅ 程序找到路径: {found_path}")
        
        # 检查mbox文件
        folder_name = watch_folder.split('/')[-1]  # ServiceNow
        mbox_file = os.path.join(found_path, folder_name)
        
        print(f"📧 mbox文件: {mbox_file}")
        
        if os.path.exists(mbox_file):
            size = os.path.getsize(mbox_file)
            print(f"✅ mbox文件存在 ({size:,} bytes)")
            
            # 解析邮件
            try:
                emails = approver.parse_mbox_file(mbox_file)
                print(f"📊 解析到 {len(emails)} 封邮件")
                
                # 显示邮件列表
                for i, email in enumerate(emails[:5]):  # 只显示前5封
                    print(f"  {i+1}. {email.get('subject', 'N/A')}")
                    print(f"     发件人: {email.get('from', 'N/A')}")
                    print(f"     需要批准: {'是' if approver.is_approval_needed(email) else '否'}")
                    print()
                
                if len(emails) > 5:
                    print(f"  ... 还有 {len(emails) - 5} 封邮件")
                    
            except Exception as e:
                print(f"❌ 解析邮件失败: {e}")
        else:
            print(f"❌ mbox文件不存在: {mbox_file}")
    else:
        print("❌ 程序无法找到文件夹")
        return False
    
    # 测试目标文件夹路径
    print(f"\n🎯 测试目标文件夹:")
    processed_destination = approver.config.get('DEFAULT', 'processed_destination')
    print(f"配置的processed_destination: {processed_destination}")
    
    target_found = approver.find_thunderbird_mail_folder(processed_destination)
    if target_found:
        print(f"✅ 找到目标路径: {target_found}")
    else:
        print("⚠️ 目标文件夹不存在，程序会尝试创建")
    
    return True

if __name__ == "__main__":
    verify_servicenow_folder()