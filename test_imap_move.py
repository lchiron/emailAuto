#!/usr/bin/env python3
"""
测试IMAP邮件处理和移动到Processed目录功能
"""

import sys
import os
sys.path.append('/Users/liqilong/Work/emailAuto')

from email_auto_approve import EmailAutoApprover

def test_imap_email_processing():
    """测试IMAP邮件处理和移动功能"""
    
    print("🎯 测试IMAP邮件处理和移动到Processed功能")
    print("=" * 60)
    
    approver = EmailAutoApprover()
    
    # 显示配置
    watch_folder = approver.config.get('DEFAULT', 'watch_folder')
    move_enabled = approver.config.getboolean('DEFAULT', 'move_processed_emails', fallback=False)
    processed_destination = approver.config.get('DEFAULT', 'processed_destination', fallback='')
    email_client = approver.config.get('DEFAULT', 'email_client', fallback='thunderbird')
    
    print(f"📧 邮件客户端: {email_client}")
    print(f"📁 监控文件夹: {watch_folder}")
    print(f"🔄 移动功能: {'启用' if move_enabled else '禁用'}")
    print(f"📂 目标文件夹: {processed_destination}")
    print("=" * 60)
    
    # 查找监控文件夹
    full_watch_path = approver.find_thunderbird_mail_folder(watch_folder)
    
    if not full_watch_path:
        print("❌ 找不到监控文件夹")
        print(f"请确保存在文件夹: {watch_folder}")
        return False
    
    print(f"✅ 找到监控文件夹: {full_watch_path}")
    
    # 检查文件夹类型
    if "webaccountMail" in full_watch_path:
        print("🌐 文件夹类型: IMAP网络邮件 (会同步到服务器)")
    elif "Local Folders" in full_watch_path:
        print("💾 文件夹类型: 本地文件夹 (不同步到服务器)")
    
    # 检查mbox文件
    folder_name = watch_folder.split('/')[-1]
    mbox_file = os.path.join(full_watch_path, folder_name)
    
    if not os.path.exists(mbox_file):
        print(f"❌ 邮件文件不存在: {mbox_file}")
        return False
    
    print(f"📊 邮件文件: {mbox_file}")
    print(f"📊 文件大小: {os.path.getsize(mbox_file):,} 字节")
    
    # 解析邮件
    try:
        emails = approver.parse_mbox_file(mbox_file)
        print(f"📧 找到 {len(emails)} 封邮件")
        
        # 检查需要处理的邮件
        need_approval = [email for email in emails if approver.is_approval_needed(email)]
        print(f"⚠️ 需要批准: {len(need_approval)} 封")
        
        # 检查已处理的邮件
        processed_emails = approver.load_processed_emails()
        print(f"✅ 已处理过: {len(processed_emails)} 封")
        
        if not need_approval:
            print("💡 没有需要处理的新邮件")
            return True
        
        # 处理第一封邮件进行测试
        test_email = need_approval[0]
        
        print(f"\n🧪 测试处理邮件:")
        print(f"  主题: {test_email.get('subject', 'N/A')}")
        print(f"  发件人: {test_email.get('from', 'N/A')}")
        
        # 检查是否已处理过
        message_id = test_email.get('message_id', '')
        if not message_id:
            message_id = f"{test_email.get('subject', '')}|{test_email.get('from', '')}"
        
        if message_id in processed_emails:
            print("⏭️ 该邮件已处理过，跳过")
            return True
        
        # 创建回复邮件
        reply_msg = approver.create_approval_reply(test_email)
        if not reply_msg:
            print("❌ 创建回复邮件失败")
            return False
        
        print(f"✉️ 生成回复: {reply_msg.get('Subject', 'N/A')}")
        
        # 发送回复（使用draft模式进行测试）
        print("📤 发送回复邮件...")
        send_success = approver.send_reply_via_email_client(reply_msg, test_email)
        
        if send_success:
            print("✅ 回复邮件发送成功")
            
            # 测试移动功能
            if move_enabled:
                print("📁 测试邮件移动功能...")
                move_success = approver.move_processed_email(test_email, mbox_file)
                
                if move_success:
                    print("✅ 邮件移动成功")
                    
                    # 验证移动结果
                    print("🔍 验证移动结果:")
                    
                    # 检查源文件邮件是否减少
                    remaining_emails = approver.parse_mbox_file(mbox_file)
                    print(f"  源文件剩余邮件: {len(remaining_emails)} 封")
                    
                    # 检查目标文件是否增加
                    target_path = approver.find_thunderbird_mail_folder(processed_destination)
                    if target_path:
                        target_folder_name = processed_destination.split('/')[-1]
                        target_mbox = os.path.join(target_path, target_folder_name)
                        if os.path.exists(target_mbox):
                            target_emails = approver.parse_mbox_file(target_mbox)
                            print(f"  目标文件邮件数: {len(target_emails)} 封")
                        else:
                            print("  目标文件不存在")
                    else:
                        print("  找不到目标文件夹")
                    
                else:
                    print("❌ 邮件移动失败")
                    return False
            else:
                print("ℹ️ 邮件移动功能已禁用")
            
            return True
        else:
            print("❌ 回复邮件发送失败")
            return False
            
    except Exception as e:
        print(f"❌ 测试过程中发生错误: {e}")
        return False

if __name__ == "__main__":
    success = test_imap_email_processing()
    if success:
        print("\n🎉 测试完成！")
    else:
        print("\n❌ 测试失败！")