#!/usr/bin/env python3
"""
邮件发送权限设置指南

解决AppleScript权限问题，实现真正的自动邮件发送
"""

import subprocess
import os

def check_accessibility_permission():
    """检查辅助功能权限"""
    print("🔍 检查系统权限...")
    
    try:
        result = subprocess.run([
            'osascript', '-e', 
            'tell application "System Events" to keystroke "test"'
        ], capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✅ 辅助功能权限已授予")
            return True
        else:
            print("❌ 需要授予辅助功能权限")
            return False
    except Exception as e:
        print(f"❌ 权限检查失败: {e}")
        return False

def setup_permissions():
    """设置权限指南"""
    print("\n" + "="*60)
    print("🔧 macOS权限设置指南")
    print("="*60)
    
    print("\n为了让程序自动控制Thunderbird发送邮件，需要授予以下权限：")
    print("\n📋 步骤1: 授予辅助功能权限")
    print("1. 打开 系统偏好设置 / 系统设置")
    print("2. 点击 安全性与隐私 / 隐私与安全性")
    print("3. 点击左侧的 辅助功能 / 无障碍")
    print("4. 点击锁形图标并输入密码")
    print("5. 找到并勾选以下应用：")
    print("   - Terminal (终端)")
    print("   - Python")
    print("   - osascript")
    print("   - 或者运行此程序的应用")
    
    print("\n📋 步骤2: 重新启动Terminal")
    print("1. 完全关闭Terminal")
    print("2. 重新打开Terminal")
    print("3. 重新运行程序")
    
    print("\n📋 备用方案：手动发送模式")
    print("如果权限设置困难，程序会：")
    print("1. 自动生成标准格式的回复邮件(.eml文件)")
    print("2. 您可以双击.eml文件在Thunderbird中打开")
    print("3. 点击发送按钮即可")
    
    print("\n💡 自动化建议：")
    print("- 可以设置Thunderbird的自动发送规则")
    print("- 或使用邮件客户端的定时发送功能")

def test_thunderbird_control():
    """测试Thunderbird控制"""
    print("\n🧪 测试Thunderbird控制权限...")
    
    try:
        # 简单测试：激活Thunderbird
        result = subprocess.run([
            'osascript', '-e', 
            'tell application "Thunderbird" to activate'
        ], capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✅ 可以控制Thunderbird")
            return True
        else:
            print(f"❌ 无法控制Thunderbird: {result.stderr}")
            return False
    except Exception as e:
        print(f"❌ 测试失败: {e}")
        return False

def enable_draft_mode():
    """启用草稿模式"""
    print("\n📝 启用邮件草稿模式...")
    
    config_file = "config.ini"
    if os.path.exists(config_file):
        # 读取配置文件
        with open(config_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 添加草稿模式设置
        if 'draft_mode' not in content:
            content += "\n# 邮件发送模式: auto=自动发送, draft=生成草稿\ndraft_mode = draft\n"
            
            with open(config_file, 'w', encoding='utf-8') as f:
                f.write(content)
            
            print("✅ 已启用草稿模式")
            print("   程序将生成.eml文件供手动发送")
        else:
            print("✅ 草稿模式配置已存在")
    else:
        print("❌ 找不到配置文件")

def main():
    print("🚀 ServiceNow邮件自动发送 - 权限设置工具")
    print()
    
    # 检查权限
    has_permission = check_accessibility_permission()
    
    if has_permission:
        # 测试Thunderbird控制
        can_control = test_thunderbird_control()
        if can_control:
            print("\n🎉 所有权限正常！程序可以自动发送邮件")
            print("运行 ./start.sh 开始自动监控和发送")
        else:
            print("\n⚠️  可以控制系统，但无法控制Thunderbird")
            setup_permissions()
    else:
        # 提供设置指南
        setup_permissions()
        
        # 启用草稿模式作为备用方案
        print("\n" + "="*60)
        print("🔄 启用备用方案...")
        enable_draft_mode()
    
    print("\n" + "="*60)
    print("💡 推荐操作：")
    print("1. 按照上述指南设置权限")
    print("2. 重启Terminal后重新运行程序")
    print("3. 或使用草稿模式 + 手动发送")

if __name__ == "__main__":
    main()