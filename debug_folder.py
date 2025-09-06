#!/usr/bin/env python3
"""
调试webaccountMail文件夹查找问题
"""

import sys
import os
sys.path.append('/Users/liqilong/Work/emailAuto')

from email_auto_approve import EmailAutoApprover

def debug_folder_finding():
    """调试文件夹查找问题"""
    
    print("🔧 调试webaccountMail文件夹查找")
    print("=" * 60)
    
    approver = EmailAutoApprover()
    
    # 手动构建路径进行测试
    profile_path = "/Users/liqilong/Library/Thunderbird/Profiles"
    mail_base_dir = "/Users/liqilong/Library/Thunderbird/Profiles/xoxled0d.default-release/Mail"
    webaccount_dir = os.path.join(mail_base_dir, "webaccountMail")
    server_dir = os.path.join(webaccount_dir, "outlook.office365.com")
    
    print(f"📂 webaccountMail目录: {webaccount_dir}")
    print(f"  存在: {os.path.exists(webaccount_dir)}")
    
    if os.path.exists(webaccount_dir):
        print(f"  内容: {os.listdir(webaccount_dir)}")
    
    print(f"📂 server目录: {server_dir}")
    print(f"  存在: {os.path.exists(server_dir)}")
    
    if os.path.exists(server_dir):
        print(f"  内容: {os.listdir(server_dir)}")
        
        # 查找ServiceNow文件夹
        servicenow_patterns = ['ServiceNow', 'ServiceNow.sbd']
        for pattern in servicenow_patterns:
            test_path = os.path.join(server_dir, pattern)
            print(f"📁 测试路径: {test_path}")
            print(f"  存在: {os.path.exists(test_path)}")
            
            if os.path.exists(test_path):
                if os.path.isdir(test_path):
                    print(f"  类型: 目录")
                    print(f"  内容: {os.listdir(test_path)}")
                elif os.path.isfile(test_path):
                    print(f"  类型: 文件")
                    print(f"  大小: {os.path.getsize(test_path)} bytes")
    
    # 测试递归查找函数
    print(f"\n🔍 测试递归查找函数:")
    
    if os.path.exists(server_dir):
        folder_parts = ['ServiceNow']
        print(f"查找: {folder_parts}")
        result = approver._find_folder_recursive(server_dir, folder_parts)
        print(f"结果: {result}")
        
        # 如果第一次失败，尝试直接查找.sbd目录
        if not result:
            servicenow_sbd = os.path.join(server_dir, 'ServiceNow.sbd')
            if os.path.exists(servicenow_sbd):
                print(f"✅ 直接找到: {servicenow_sbd}")
                return servicenow_sbd
    
    return None

if __name__ == "__main__":
    debug_folder_finding()