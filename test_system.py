#!/usr/bin/env python3
"""
系统测试脚本

测试邮件自动批准程序的各个组件
"""

import os
import sys
import tempfile
import shutil
from pathlib import Path

def test_config():
    """测试配置文件加载"""
    print("测试配置文件...")
    try:
        from email_auto_approve import EmailAutoApprover
        approver = EmailAutoApprover()
        print("✓ 配置文件加载成功")
        
        # 检查关键配置项
        profile_path = approver.config.get('DEFAULT', 'thunderbird_profile_path')
        watch_folder = approver.config.get('DEFAULT', 'watch_folder')
        from_email = approver.config.get('EMAIL', 'from_email')
        
        print(f"  - Thunderbird路径: {profile_path}")
        print(f"  - 监控文件夹: {watch_folder}")
        print(f"  - 发件人邮箱: {from_email}")
        
        return True
    except Exception as e:
        print(f"✗ 配置文件测试失败: {e}")
        return False

def test_email_parsing():
    """测试邮件解析功能"""
    print("\\n测试邮件解析...")
    try:
        from email_auto_approve import EmailAutoApprover
        approver = EmailAutoApprover()
        
        # 使用示例邮件文件
        eml_file = "Re_ RITM1602185 - approve - Edward Li <eli23@lululemon.com> - 2025-09-06 0804.eml"
        if not os.path.exists(eml_file):
            print(f"✗ 示例邮件文件不存在: {eml_file}")
            return False
        
        email_info = approver.parse_email_file(eml_file)
        if email_info:
            print("✓ 邮件解析成功")
            print(f"  - 主题: {email_info.get('subject', 'N/A')}")
            print(f"  - 发件人: {email_info.get('from', 'N/A')}")
            print(f"  - 收件人: {email_info.get('to', 'N/A')}")
            
            # 测试批准识别
            needs_approval = approver.is_approval_needed(email_info)
            print(f"  - 需要批准: {needs_approval}")
            
            return True
        else:
            print("✗ 邮件解析失败")
            return False
            
    except Exception as e:
        print(f"✗ 邮件解析测试失败: {e}")
        return False

def test_reply_creation():
    """测试回复邮件创建"""
    print("\\n测试回复邮件创建...")
    try:
        from email_auto_approve import EmailAutoApprover
        approver = EmailAutoApprover()
        
        # 模拟邮件信息
        email_info = {
            'subject': 'RITM1602185 - approve',
            'from': 'test@service-now.com',
            'reply_to': 'test@service-now.com',
            'message_id': '<test123@example.com>'
        }
        
        reply_msg = approver.create_approval_reply(email_info)
        if reply_msg:
            print("✓ 回复邮件创建成功")
            print(f"  - 收件人: {reply_msg['To']}")
            print(f"  - 主题: {reply_msg['Subject']}")
            print(f"  - 发件人: {reply_msg['From']}")
            
            # 保存测试回复邮件
            test_reply_file = "test_reply.eml"
            with open(test_reply_file, 'w', encoding='utf-8') as f:
                f.write(reply_msg.as_string())
            print(f"  - 测试回复已保存: {test_reply_file}")
            
            return True
        else:
            print("✗ 回复邮件创建失败")
            return False
            
    except Exception as e:
        print(f"✗ 回复邮件创建测试失败: {e}")
        return False

def test_thunderbird_sender():
    """测试Thunderbird发送器"""
    print("\\n测试Thunderbird发送器...")
    try:
        from thunderbird_sender import ThunderbirdSender
        sender = ThunderbirdSender()
        
        print(f"✓ ThunderbirdSender初始化成功")
        print(f"  - 操作系统: {sender.system}")
        
        # 检查Thunderbird是否运行
        if sender.system == "Darwin":
            is_running = sender.is_thunderbird_running()
            print(f"  - Thunderbird运行状态: {is_running}")
        
        # 创建测试草稿文件
        test_draft = sender.create_draft_file(
            "test@example.com",
            "Test Email",
            "This is a test email."
        )
        
        print(f"  - 测试草稿文件: {test_draft}")
        
        return True
        
    except Exception as e:
        print(f"✗ Thunderbird发送器测试失败: {e}")
        return False

def test_dependencies():
    """测试依赖包"""
    print("测试Python依赖...")
    
    dependencies = [
        'watchdog',
        'email',
        'configparser',
        'json',
        'logging'
    ]
    
    all_ok = True
    for dep in dependencies:
        try:
            __import__(dep)
            print(f"✓ {dep}")
        except ImportError:
            print(f"✗ {dep} - 未安装")
            all_ok = False
    
    return all_ok

def create_test_environment():
    """创建测试环境"""
    print("\\n创建测试环境...")
    
    # 创建测试目录结构
    test_dir = "test_thunderbird"
    watch_dir = os.path.join(test_dir, "Mail", "Local Folders", "Archive", "ServiceNow", "NeedApprove")
    
    try:
        os.makedirs(watch_dir, exist_ok=True)
        print(f"✓ 创建测试目录: {watch_dir}")
        
        # 复制示例邮件到测试目录
        sample_email = "Re_ RITM1602185 - approve - Edward Li <eli23@lululemon.com> - 2025-09-06 0804.eml"
        if os.path.exists(sample_email):
            test_email = os.path.join(watch_dir, "test_email.eml")
            shutil.copy2(sample_email, test_email)
            print(f"✓ 复制测试邮件: {test_email}")
        
        return test_dir
        
    except Exception as e:
        print(f"✗ 创建测试环境失败: {e}")
        return None

def cleanup_test_files():
    """清理测试文件"""
    print("\\n清理测试文件...")
    
    test_files = [
        "test_reply.eml",
        "processed_emails.json",
        "config.ini.backup"
    ]
    
    for filename in test_files:
        if os.path.exists(filename):
            try:
                os.remove(filename)
                print(f"✓ 删除: {filename}")
            except Exception as e:
                print(f"✗ 删除失败 {filename}: {e}")
    
    # 删除测试目录
    if os.path.exists("test_thunderbird"):
        try:
            shutil.rmtree("test_thunderbird")
            print("✓ 删除测试目录")
        except Exception as e:
            print(f"✗ 删除测试目录失败: {e}")

def main():
    """主测试函数"""
    print("ServiceNow邮件自动批准程序 - 系统测试")
    print("=" * 50)
    
    tests = [
        ("依赖包检查", test_dependencies),
        ("配置文件", test_config),
        ("邮件解析", test_email_parsing),
        ("回复创建", test_reply_creation),
        ("发送器", test_thunderbird_sender),
    ]
    
    results = []
    
    for test_name, test_func in tests:
        print(f"\\n[{test_name}]")
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"✗ 测试异常: {e}")
            results.append((test_name, False))
    
    # 汇总结果
    print("\\n" + "=" * 50)
    print("测试结果汇总:")
    
    passed = 0
    for test_name, result in results:
        status = "通过" if result else "失败"
        print(f"  {test_name}: {status}")
        if result:
            passed += 1
    
    print(f"\\n总计: {passed}/{len(results)} 项测试通过")
    
    if passed == len(results):
        print("\\n🎉 所有测试通过！程序已准备就绪。")
    else:
        print(f"\\n⚠️  有 {len(results) - passed} 项测试失败，请检查相关组件。")
    
    # 清理测试文件
    cleanup_test_files()
    
    return 0 if passed == len(results) else 1

if __name__ == "__main__":
    sys.exit(main())