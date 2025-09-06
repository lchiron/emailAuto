#!/usr/bin/env python3
"""
测试使用Outlook PWA的完整邮件处理流程
"""

import sys
import os
sys.path.append('/Users/liqilong/Work/emailAuto')

from email_auto_approve import EmailAutoApprover

def test_complete_outlook_workflow():
    """测试完整的Outlook PWA工作流程"""
    
    print("🎯 测试完整的Outlook PWA邮件自动批准流程")
    print("=" * 60)
    
    # 创建邮件自动批准器
    approver = EmailAutoApprover()
    
    # 显示当前配置
    email_client = approver.config.get('DEFAULT', 'email_client', fallback='thunderbird')
    outlook_browser = approver.config.get('DEFAULT', 'outlook_browser', fallback='Safari')
    send_mode = approver.config.get('DEFAULT', 'send_mode', fallback='draft')
    
    print(f"📧 邮件客户端: {email_client}")
    if email_client.lower() == 'outlook':
        print(f"🌐 使用浏览器: {outlook_browser}")
    print(f"🚀 发送模式: {send_mode}")
    print("=" * 60)
    
    # 查找并处理邮件
    watch_folder = approver.config.get('DEFAULT', 'watch_folder')
    full_watch_path = approver.find_thunderbird_mail_folder(watch_folder)
    
    if not full_watch_path:
        print("❌ 找不到Thunderbird邮件文件夹")
        return False
    
    print(f"✅ 找到目标目录: {full_watch_path}")
    
    folder_name = watch_folder.split('/')[-1]
    mbox_file = os.path.join(full_watch_path, folder_name)
    
    if not os.path.exists(mbox_file):
        print("❌ mbox文件不存在")
        return False
    
    print(f"📊 mbox文件: {mbox_file}")
    print(f"📊 文件大小: {os.path.getsize(mbox_file)} 字节")
    
    # 处理邮件（只处理一封进行测试）
    print("\n🔄 开始处理邮件...")
    
    try:
        # 解析邮件
        emails = approver.parse_mbox_file(mbox_file)
        print(f"📧 找到 {len(emails)} 封邮件")
        
        # 处理第一封需要批准的邮件
        processed = False
        for email_info in emails:
            if approver.is_approval_needed(email_info):
                # 检查是否已处理过
                message_id = email_info.get('message_id', '')
                if not message_id:
                    message_id = f"{email_info.get('subject', '')}|{email_info.get('from', '')}"
                
                processed_emails = approver.load_processed_emails()
                if message_id in processed_emails:
                    print(f"⏭️ 跳过已处理的邮件: {email_info['subject']}")
                    continue
                
                print(f"\n📧 处理邮件: {email_info['subject']}")
                print(f"📤 发件人: {email_info['from']}")
                
                # 创建回复
                reply_msg = approver.create_approval_reply(email_info)
                if reply_msg:
                    print(f"✉️ 回复主题: {reply_msg['Subject']}")
                    print(f"📤 回复收件人: {reply_msg['To']}")
                    
                    # 发送邮件
                    success = approver.send_reply_via_email_client(reply_msg, email_info)
                    if success:
                        print("✅ 邮件处理成功")
                        processed = True
                    else:
                        print("❌ 邮件处理失败")
                    
                    break  # 只处理一封进行测试
        
        if not processed:
            print("⚠️ 没有找到需要处理的新邮件")
        
        return processed
        
    except Exception as e:
        print(f"❌ 处理过程中出现错误: {e}")
        return False

if __name__ == "__main__":
    test_complete_outlook_workflow()