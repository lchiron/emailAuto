#!/usr/bin/env python3
"""
创建测试RITM邮件到NeedApprove目录
"""

import email
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

def create_test_ritm_email():
    """创建测试RITM邮件"""
    
    # 创建邮件对象
    msg = MIMEMultipart()
    msg['From'] = 'luluprod@service-now.com'
    msg['To'] = 'eli23@lululemon.com'
    msg['Subject'] = 'RITM9999999 | Approval Required / Approbation Requise'
    msg['Date'] = email.utils.formatdate(localtime=True)
    msg['Message-ID'] = f'<test.{int(datetime.now().timestamp())}@service-now.com>'
    
    # 邮件正文
    body = """
Approval Required

Hey Edward,

We have a request that requires your approval. Please see below for the request details and use the buttons to approve or deny this request. The request can be fulfilled only once approval is received.

Reference: RITM9999999

Approve
Reject

Best regards,
ServiceNow Team
    """.strip()
    
    msg.attach(MIMEText(body, 'plain'))
    
    # 目标路径
    needapprove_path = "/Users/liqilong/Library/thunderbird/Profiles/xoxled0d.default-release/webaccountMail/outlook.office365.com/ServiceNow.sbd/NeedApprove"
    
    print(f"📧 创建测试邮件到: {needapprove_path}")
    
    # 检查文件是否存在
    if not os.path.exists(needapprove_path):
        print("❌ NeedApprove文件不存在，创建空文件")
        with open(needapprove_path, 'w') as f:
            f.write("")
    
    # 追加邮件到mbox文件
    with open(needapprove_path, 'a') as f:
        # mbox格式需要From行开始
        f.write(f"From - {datetime.now().strftime('%a %b %d %H:%M:%S %Y')}\n")
        f.write(msg.as_string())
        f.write("\n\n")  # mbox邮件之间的分隔
    
    print("✅ 测试邮件已创建")
    
    # 显示文件大小
    size = os.path.getsize(needapprove_path)
    print(f"📄 NeedApprove文件大小: {size} bytes")

if __name__ == "__main__":
    create_test_ritm_email()