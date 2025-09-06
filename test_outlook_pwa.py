#!/usr/bin/env python3
"""
自动测试Outlook PWA邮件发送功能
"""

import sys
import os
sys.path.append('/Users/liqilong/Work/emailAuto')

from outlook_sender import OutlookPWASender

def test_outlook_pwa_auto():
    """自动测试Outlook PWA邮件发送功能"""
    
    # 使用Chrome作为默认浏览器
    browser = "Google Chrome"
    sender = OutlookPWASender(browser=browser)
    
    # 测试数据
    to_addr = "luluprod@service-now.com"
    subject = "Re: RITM1234567 - approve"
    body_text = "Ref:MSG12345678"
    
    print("🧪 测试Outlook PWA邮件发送...")
    print(f"浏览器: {browser}")
    print(f"收件人: {to_addr}")
    print(f"主题: {subject}")
    print(f"正文: {body_text}")
    
    # 使用draft模式进行测试，避免实际发送
    auto_send = False
    print(f"测试模式: {'自动发送' if auto_send else '只准备，不发送'}")
    
    try:
        result = sender.send_email(to_addr, subject, body_text, auto_send=auto_send)
        
        if result:
            print("✅ Outlook PWA邮件处理成功！")
            print("📧 请检查Chrome中的Outlook PWA是否正确准备了邮件")
            return True
        else:
            print("❌ Outlook PWA邮件处理失败！")
            return False
            
    except Exception as e:
        print(f"❌ 测试过程中出现错误: {e}")
        return False

if __name__ == "__main__":
    test_outlook_pwa_auto()