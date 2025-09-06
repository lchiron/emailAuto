#!/usr/bin/env python3
"""
测试主题格式生成
"""

import sys
import os
sys.path.append('/Users/liqilong/Work/emailAuto')

from email_auto_approve import EmailAutoApprover

def test_subject_generation():
    """测试主题格式生成"""
    
    approver = EmailAutoApprover()
    
    # 测试不同的原始主题
    test_subjects = [
        "RITM1603887 | Approval Required / Approbation Requise",
        "RITM1601061 | Approval Required / Approbation Requise", 
        "RITM1234567 | Some Other Format",
        "Re: RITM9999999 | Already has Re:",
    ]
    
    import re
    
    for original_subject in test_subjects:
        print(f"\n原始主题: {original_subject}")
        
        # 提取RITM号码的逻辑（从email_auto_approve.py复制）
        ritm_match = re.search(r'RITM(\d+)', original_subject, re.IGNORECASE)
        
        if ritm_match:
            ritm_number = ritm_match.group(1)
            reply_subject = f"Re: RITM{ritm_number} - approve"
            print(f"生成主题: {reply_subject}")
        else:
            if not original_subject.lower().startswith('re:'):
                reply_subject = f"Re: {original_subject}"
            else:
                reply_subject = original_subject
            print(f"备用主题: {reply_subject}")

if __name__ == "__main__":
    test_subject_generation()