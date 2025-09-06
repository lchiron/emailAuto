#!/usr/bin/env python3
"""
Thunderbird文件夹检测工具

帮助用户查找和设置正确的Thunderbird邮件文件夹路径
"""

import os
import platform
from pathlib import Path
import sys

class ThunderbirdFolderDetector:
    def __init__(self):
        self.system = platform.system()
        self.user_home = Path.home()
    
    def get_thunderbird_profiles_path(self):
        """获取Thunderbird配置文件路径"""
        if self.system == "Windows":
            return self.user_home / "AppData" / "Roaming" / "Thunderbird" / "Profiles"
        elif self.system == "Darwin":  # macOS
            return self.user_home / "Library" / "Thunderbird" / "Profiles"
        else:  # Linux
            return self.user_home / ".thunderbird"
    
    def scan_thunderbird_structure(self):
        """扫描Thunderbird文件结构"""
        profiles_path = self.get_thunderbird_profiles_path()
        
        print(f"扫描Thunderbird配置目录: {profiles_path}")
        print("=" * 60)
        
        if not profiles_path.exists():
            print("❌ Thunderbird配置目录不存在")
            print("请确保Thunderbird已安装并至少运行过一次")
            return False
        
        profiles_found = False
        
        for profile_item in profiles_path.iterdir():
            if profile_item.is_dir():
                profiles_found = True
                print(f"\\n📁 配置文件: {profile_item.name}")
                
                mail_dir = profile_item / "Mail"
                if mail_dir.exists():
                    print("  ✅ Mail目录存在")
                    self.scan_mail_accounts(mail_dir)
                else:
                    print("  ❌ Mail目录不存在")
        
        if not profiles_found:
            print("❌ 没有找到任何Thunderbird配置文件")
            return False
        
        return True
    
    def scan_mail_accounts(self, mail_dir):
        """扫描邮件账户目录"""
        for account_item in mail_dir.iterdir():
            if account_item.is_dir():
                print(f"    📧 邮件账户: {account_item.name}")
                
                if account_item.name == "Local Folders":
                    self.scan_local_folders(account_item)
                else:
                    # 扫描IMAP/POP账户文件夹
                    self.scan_account_folders(account_item, "    ")
    
    def scan_local_folders(self, local_folders_dir):
        """扫描本地文件夹"""
        print("      🔍 扫描本地文件夹...")
        
        # 查找Archive文件夹
        archive_paths = self.find_folder_variations(local_folders_dir, "Archive")
        
        if archive_paths:
            for archive_path in archive_paths:
                print(f"      📂 找到Archive: {archive_path.relative_to(local_folders_dir)}")
                
                # 查找ServiceNow文件夹
                servicenow_paths = self.find_folder_variations(archive_path, "ServiceNow")
                
                if servicenow_paths:
                    for servicenow_path in servicenow_paths:
                        print(f"        📂 找到ServiceNow: {servicenow_path.relative_to(local_folders_dir)}")
                        
                        # 查找NeedApprove文件夹
                        needapprove_paths = self.find_folder_variations(servicenow_path, "NeedApprove")
                        
                        if needapprove_paths:
                            for needapprove_path in needapprove_paths:
                                print(f"          ✅ 找到目标文件夹: {needapprove_path}")
                                
                                # 检查文件夹内容
                                eml_files = list(needapprove_path.glob("*.eml"))
                                if eml_files:
                                    print(f"          📧 包含 {len(eml_files)} 个邮件文件")
                                else:
                                    print("          📭 文件夹为空")
                        else:
                            print("          ❌ 未找到NeedApprove文件夹")
                            print("          💡 请在Thunderbird中创建此文件夹")
                else:
                    print("        ❌ 未找到ServiceNow文件夹")
                    print("        💡 请在Archive下创建ServiceNow文件夹")
        else:
            print("      ❌ 未找到Archive文件夹")
            print("      💡 请在本地文件夹中创建Archive文件夹")
            
            # 显示现有的本地文件夹
            print("      📋 现有本地文件夹:")
            self.scan_account_folders(local_folders_dir, "        ")
    
    def find_folder_variations(self, parent_dir, folder_name):
        """查找文件夹的各种变体"""
        if not parent_dir.exists():
            return []
        
        found_paths = []
        
        try:
            # 检查父目录是否真的是目录
            if not parent_dir.is_dir():
                # 如果是文件，检查是否有对应的.sbd目录
                sbd_path = parent_dir.parent / f"{parent_dir.name}.sbd"
                if sbd_path.exists() and sbd_path.is_dir():
                    parent_dir = sbd_path
                else:
                    return []
            
            # 直接查找文件夹
            direct_path = parent_dir / folder_name
            if direct_path.exists() and direct_path.is_dir():
                found_paths.append(direct_path)
            
            # 查找同名文件（Thunderbird邮件文件）
            file_path = parent_dir / folder_name
            if file_path.exists() and file_path.is_file():
                # 检查是否有对应的.sbd目录
                sbd_path = parent_dir / f"{folder_name}.sbd"
                if sbd_path.exists() and sbd_path.is_dir():
                    found_paths.append(sbd_path)
                else:
                    # 如果只有文件没有.sbd目录，也算找到了（叶子文件夹）
                    found_paths.append(file_path)
            
            # 查找.sbd目录
            sbd_path = parent_dir / f"{folder_name}.sbd"
            if sbd_path.exists() and sbd_path.is_dir() and sbd_path not in found_paths:
                found_paths.append(sbd_path)
            
            # 模糊匹配（忽略大小写）
            for item in parent_dir.iterdir():
                item_name_lower = item.name.lower()
                folder_name_lower = folder_name.lower()
                
                if item_name_lower == folder_name_lower and item not in found_paths:
                    if item.is_dir():
                        found_paths.append(item)
                    elif item.is_file():
                        # 检查是否有对应的.sbd目录
                        sbd_path = parent_dir / f"{item.name}.sbd"
                        if sbd_path.exists() and sbd_path.is_dir():
                            found_paths.append(sbd_path)
                        else:
                            found_paths.append(item)
                
                # 也检查.sbd结尾的目录
                elif item_name_lower == f"{folder_name_lower}.sbd" and item.is_dir() and item not in found_paths:
                    found_paths.append(item)
        
        except (PermissionError, OSError) as e:
            print(f"      ❌ 访问权限错误: {e}")
        
        return found_paths
    
    def scan_account_folders(self, account_dir, indent=""):
        """扫描账户文件夹内容"""
        try:
            for item in account_dir.iterdir():
                if item.is_dir():
                    print(f"{indent}📁 {item.name}")
                elif item.suffix in ['.msf', '']:
                    # 这些是Thunderbird的邮件文件
                    print(f"{indent}📄 {item.name}")
        except PermissionError:
            print(f"{indent}❌ 权限不足，无法访问")
    
    def create_folder_guide(self):
        """创建文件夹设置指南"""
        print("\\n" + "=" * 60)
        print("📋 Thunderbird文件夹设置指南")
        print("=" * 60)
        
        print("\\n1. 在Thunderbird中，右键点击'本地文件夹'（Local Folders）")
        print("2. 选择'新建文件夹'")
        print("3. 创建以下文件夹结构:")
        print("   📁 Archive")
        print("   ├── 📁 ServiceNow")
        print("   │   └── 📁 NeedApprove")
        
        print("\\n4. 将需要自动批准的ServiceNow邮件移动或复制到'NeedApprove'文件夹中")
        print("5. 程序将监控此文件夹并自动处理新邮件")
        
        print("\\n💡 提示:")
        print("- 可以通过拖拽方式将邮件移动到目标文件夹")
        print("- 也可以设置邮件过滤规则自动将ServiceNow邮件移动到此文件夹")
        print("- 程序会记录已处理的邮件，避免重复处理")

def main():
    print("🔍 Thunderbird邮件文件夹检测工具")
    print("帮助您找到并设置ServiceNow自动批准邮件文件夹")
    
    detector = ThunderbirdFolderDetector()
    
    # 扫描Thunderbird结构
    if detector.scan_thunderbird_structure():
        detector.create_folder_guide()
    else:
        print("\\n❌ 无法检测到Thunderbird安装")
        print("请确保:")
        print("1. Thunderbird已正确安装")
        print("2. Thunderbird至少运行过一次以创建配置文件")
        print("3. 使用管理员权限运行此工具（如果需要）")
        return 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())