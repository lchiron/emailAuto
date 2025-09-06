#!/usr/bin/env python3
"""
检查Thunderbird邮件文件夹类型和同步状态
"""

import os
import sys
sys.path.append('/Users/liqilong/Work/emailAuto')

from email_auto_approve import EmailAutoApprover

def check_folder_type():
    """检查当前监控文件夹的类型"""
    
    approver = EmailAutoApprover()
    
    # 获取Thunderbird配置路径
    profile_path = approver.config.get('DEFAULT', 'thunderbird_profile_path')
    watch_folder = approver.config.get('DEFAULT', 'watch_folder')
    
    print(f"🔍 检查Thunderbird邮件文件夹类型")
    print(f"配置路径: {profile_path}")
    print(f"监控文件夹: {watch_folder}")
    print("=" * 60)
    
    # 查找文件夹
    full_watch_path = approver.find_thunderbird_mail_folder(watch_folder)
    
    if not full_watch_path:
        print("❌ 找不到目标文件夹")
        return
    
    print(f"✅ 找到文件夹: {full_watch_path}")
    
    # 分析路径结构判断文件夹类型
    if "Local Folders" in full_watch_path:
        folder_type = "本地文件夹 (Local Folders)"
        sync_risk = "无风险 - 完全本地存储"
        recommendation = "✅ 安全使用，不会影响服务器"
    elif "webaccountMail" in full_watch_path or "ImapMail" in full_watch_path:
        folder_type = "IMAP网络邮件文件夹"
        sync_risk = "⚠️  高风险 - 会同步到服务器"
        recommendation = "建议复制邮件到Local Folders处理"
    else:
        # 检查是否在Mail目录下的特定账户文件夹
        mail_parts = full_watch_path.split(os.sep)
        if "Mail" in mail_parts:
            mail_index = mail_parts.index("Mail")
            if len(mail_parts) > mail_index + 1:
                account_folder = mail_parts[mail_index + 1]
                if "Local Folders" in account_folder:
                    folder_type = "本地文件夹"
                    sync_risk = "无风险"
                    recommendation = "✅ 安全使用"
                else:
                    folder_type = f"网络邮件账户 ({account_folder})"
                    sync_risk = "⚠️  可能同步到服务器"
                    recommendation = "建议确认账户类型"
            else:
                folder_type = "未知类型"
                sync_risk = "未知风险"
                recommendation = "建议手动确认"
        else:
            folder_type = "未知类型"
            sync_risk = "未知风险" 
            recommendation = "建议手动确认"
    
    print(f"📁 文件夹类型: {folder_type}")
    print(f"🔄 同步风险: {sync_risk}")
    print(f"💡 建议: {recommendation}")
    
    # 检查文件夹中的邮件数量
    folder_name = watch_folder.split('/')[-1]
    mbox_file = os.path.join(full_watch_path, folder_name)
    
    if os.path.exists(mbox_file):
        file_size = os.path.getsize(mbox_file)
        print(f"📊 邮件文件大小: {file_size:,} 字节")
        
        # 解析邮件数量
        try:
            emails = approver.parse_mbox_file(mbox_file)
            print(f"📧 邮件数量: {len(emails)} 封")
        except Exception as e:
            print(f"❌ 解析邮件失败: {e}")
    else:
        print("📧 邮件文件不存在")
    
    print("\n" + "=" * 60)
    print("💡 安全建议:")
    print("1. 如果是IMAP文件夹，考虑使用Local Folders")
    print("2. 程序当前只读取邮件，不修改文件，相对安全")
    print("3. 但建议备份重要邮件")

if __name__ == "__main__":
    check_folder_type()