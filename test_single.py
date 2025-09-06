#!/usr/bin/env python3
"""
测试单个邮件的详细处理过程
"""

import sys
import os
sys.path.append('/Users/liqilong/Work/emailAuto')

from email_auto_approve import EmailAutoApprover

def test_single_email_processing():
    """测试单个邮件的详细处理过程"""
    
    approver = EmailAutoApprover()
    
    # 查找mbox文件
    watch_folder = approver.config.get('DEFAULT', 'watch_folder')
    full_watch_path = approver.find_thunderbird_mail_folder(watch_folder)
    folder_name = watch_folder.split('/')[-1]
    mbox_file = os.path.join(full_watch_path, folder_name)
    
    # 解析邮件
    emails = approver.parse_mbox_file(mbox_file)
    
    # 测试第一封需要批准的邮件
    for email_info in emails:
        if approver.is_approval_needed(email_info):
            print(f"📧 原始邮件信息:")
            print(f"  主题: {email_info.get('subject', 'N/A')}")
            print(f"  发件人: {email_info.get('from', 'N/A')}")
            print(f"  回复到: {email_info.get('reply_to', 'N/A')}")
            
            # 创建回复邮件
            reply_msg = approver.create_approval_reply(email_info)
            
            if reply_msg:
                print(f"\n✉️  回复邮件信息:")
                print(f"  From: {reply_msg.get('From', 'N/A')}")
                print(f"  To: {reply_msg.get('To', 'N/A')}")
                print(f"  Subject: {reply_msg.get('Subject', 'N/A')}")
                print(f"  Date: {reply_msg.get('Date', 'N/A')}")
                
                # 提取纯文本内容
                body_text = ""
                if reply_msg.is_multipart():
                    for part in reply_msg.walk():
                        if part.get_content_type() == "text/plain":
                            try:
                                payload = part.get_payload(decode=True)
                                if payload:
                                    body_text = payload.decode('utf-8', errors='ignore')
                                    break
                            except:
                                continue
                
                print(f"  Body: '{body_text}'")
                
                print(f"\n🔍 AppleScript将发送的内容:")
                print(f"  收件人字段: {reply_msg.get('To', 'N/A')}")
                print(f"  主题字段: {reply_msg.get('Subject', 'N/A')}")  
                print(f"  正文字段: '{body_text}'")
                
            break

if __name__ == "__main__":
    test_single_email_processing()