#!/usr/bin/env python3
"""
检查实际的Thunderbird文件夹结构，特别是IMAP文件夹
"""

import sys
import os
sys.path.append('/Users/liqilong/Work/emailAuto')

from email_auto_approve import EmailAutoApprover

def explore_thunderbird_structure():
    """探索Thunderbird文件夹结构"""
    
    approver = EmailAutoApprover()
    profile_path = approver.config.get('DEFAULT', 'thunderbird_profile_path')
    
    print("🔍 探索Thunderbird文件夹结构")
    print(f"配置路径: {profile_path}")
    print("=" * 80)
    
    if not os.path.exists(profile_path):
        print("❌ Thunderbird配置路径不存在")
        return
    
    # 遍历所有profile目录
    for profile_item in os.listdir(profile_path):
        profile_dir = os.path.join(profile_path, profile_item)
        if not os.path.isdir(profile_dir):
            continue
        
        print(f"\n📁 Profile: {profile_item}")
        
        mail_dir = os.path.join(profile_dir, 'Mail')
        if not os.path.exists(mail_dir):
            print("  ❌ 没有Mail目录")
            continue
        
        print(f"  📂 Mail目录: {mail_dir}")
        
        # 遍历Mail目录下的所有内容
        for mail_item in os.listdir(mail_dir):
            mail_item_path = os.path.join(mail_dir, mail_item)
            
            if os.path.isdir(mail_item_path):
                print(f"    📁 {mail_item}/")
                
                # 如果是webaccountMail，详细探索
                if mail_item == 'webaccountMail':
                    print(f"      🌐 IMAP账户目录:")
                    try:
                        for server_item in os.listdir(mail_item_path):
                            server_path = os.path.join(mail_item_path, server_item)
                            if os.path.isdir(server_path):
                                print(f"        📧 {server_item}/")
                                
                                # 探索服务器目录下的文件夹
                                try:
                                    for folder_item in os.listdir(server_path):
                                        folder_path = os.path.join(server_path, folder_item)
                                        if os.path.isdir(folder_path):
                                            print(f"          📁 {folder_item}/")
                                        elif os.path.isfile(folder_path):
                                            size = os.path.getsize(folder_path)
                                            print(f"          📄 {folder_item} ({size:,} bytes)")
                                            
                                            # 如果是ServiceNow相关，特别标注
                                            if 'servicenow' in folder_item.lower():
                                                print(f"              ⭐ ServiceNow相关文件!")
                                except PermissionError:
                                    print(f"        ❌ 权限不足，无法访问 {server_item}")
                                except Exception as e:
                                    print(f"        ❌ 错误: {e}")
                    except Exception as e:
                        print(f"      ❌ 无法访问webaccountMail: {e}")
                
                # 如果是Local Folders，也详细探索
                elif mail_item == 'Local Folders':
                    print(f"      💾 本地文件夹:")
                    try:
                        for local_item in os.listdir(mail_item_path):
                            local_path = os.path.join(mail_item_path, local_item)
                            if os.path.isdir(local_path):
                                print(f"        📁 {local_item}/")
                            elif os.path.isfile(local_path):
                                size = os.path.getsize(local_path)
                                print(f"        📄 {local_item} ({size:,} bytes)")
                    except Exception as e:
                        print(f"      ❌ 无法访问Local Folders: {e}")
                
                # 其他目录简单列出
                else:
                    try:
                        item_count = len(os.listdir(mail_item_path))
                        print(f"      📊 包含 {item_count} 个项目")
                    except:
                        print(f"      ❌ 无法访问")
            
            elif os.path.isfile(mail_item_path):
                size = os.path.getsize(mail_item_path)
                print(f"    📄 {mail_item} ({size:,} bytes)")
    
    print("\n" + "=" * 80)
    print("💡 查找建议:")
    print("1. 查找包含'outlook'或'office365'的文件夹")
    print("2. 查找包含'ServiceNow'的文件或文件夹")
    print("3. 检查webaccountMail下的服务器目录")

if __name__ == "__main__":
    explore_thunderbird_structure()