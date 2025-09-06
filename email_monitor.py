#!/usr/bin/env python3
"""
ServiceNow 邮件自动批准监控程序 - 简化版
直接使用绝对路径，避免路径查找问题
"""

import sys
import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

sys.path.append('/Users/liqilong/Work/emailAuto')

from email_auto_approve import EmailAutoApprover

class EmailHandler(FileSystemEventHandler):
    def __init__(self, approver, mbox_file_path):
        self.approver = approver
        self.mbox_file_path = mbox_file_path
        self.last_processed_time = time.time()
        
    def on_modified(self, event):
        # 当mbox文件被修改时触发
        if event.src_path == self.mbox_file_path:
            current_time = time.time()
            # 避免频繁触发，至少间隔10秒
            if current_time - self.last_processed_time > 10:
                self.last_processed_time = current_time
                print(f"📧 检测到邮件文件变化: {event.src_path}")
                self.process_emails()
    
    def process_emails(self):
        """处理邮件"""
        try:
            result = self.approver.process_mbox_file(self.mbox_file_path)
            if result:
                print("✅ 邮件处理完成")
            else:
                print("ℹ️ 没有新邮件需要处理")
        except Exception as e:
            print(f"❌ 处理邮件时出错: {e}")

def main():
    print("🎯 ServiceNow 邮件自动批准监控程序")
    print("=" * 60)
    
    # 直接使用绝对路径
    mbox_file_path = "/Users/liqilong/Library/Thunderbird/Profiles/xoxled0d.default-release/Mail/Local Folders/Archives.sbd/ServiceNow.sbd/NeedApprove"
    watch_dir = "/Users/liqilong/Library/Thunderbird/Profiles/xoxled0d.default-release/Mail/Local Folders/Archives.sbd/ServiceNow.sbd"
    
    if not os.path.exists(mbox_file_path):
        print(f"❌ 邮件文件不存在: {mbox_file_path}")
        return
    
    print(f"📁 监控邮件文件: {mbox_file_path}")
    
    approver = EmailAutoApprover()
    
    # 首次启动时处理现有邮件
    print("🔍 检查现有邮件...")
    handler = EmailHandler(approver, mbox_file_path)
    handler.process_emails()
    
    # 设置文件监控
    observer = Observer()
    observer.schedule(handler, watch_dir, recursive=False)
    observer.start()
    
    print("👁️ 开始监控邮件文件变化...")
    print("按 Ctrl+C 停止监控")
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n🛑 停止监控")
        observer.stop()
    
    observer.join()

if __name__ == "__main__":
    main()