# 🎯 ServiceNow邮件自动批准 - 完整使用指南

## ✅ 功能状态

**所有功能已完美实现并测试通过！**

- ✅ **Thunderbird文件夹检测** - 正确识别程序内文件夹路径
- ✅ **mbox邮件解析** - 完美处理Thunderbird邮件格式  
- ✅ **邮件智能识别** - 自动识别ServiceNow批准邮件
- ✅ **回复邮件生成** - 创建标准格式批准回复
- ✅ **AppleScript自动发送** - 控制Thunderbird发送邮件
- ✅ **灵活发送模式** - 支持自动发送或手动确认

## 🚀 快速开始

### 1️⃣ 设置Thunderbird文件夹

在Thunderbird中创建文件夹结构：
```
本地文件夹 (Local Folders)
└── 📁 Archive
    └── 📁 ServiceNow
        └── 📁 NeedApprove  ← 程序监控此文件夹
```

### 2️⃣ 验证设置
```bash
python detect_thunderbird_folders.py
```
应显示：✅ 找到目标文件夹

### 3️⃣ 选择发送模式

编辑 `config.ini`：
```ini
# 自动发送模式（完全自动化）
send_mode = auto

# 或手动确认模式（准备邮件，您手动发送）
send_mode = draft
```

### 4️⃣ 启动程序
```bash
./start.sh
```

## ⚙️ 发送模式说明

### 🤖 自动发送模式 (`send_mode = auto`)
- 程序自动填写收件人、主题、正文
- **自动点击发送按钮**
- 完全无人工干预
- 适合：完全信任程序的场景

### 👤 手动确认模式 (`send_mode = draft`)  
- 程序自动填写收件人、主题、正文
- **不自动发送，等待您手动点击发送**
- 您可以检查内容后再发送
- 适合：需要最终确认的场景

## 🔧 配置选项

```ini
[DEFAULT]
# Thunderbird程序路径（自动检测）
thunderbird_profile_path = /Users/用户名/Library/Thunderbird/Profiles

# 监控的文件夹路径（Thunderbird内的结构）
watch_folder = Archive/ServiceNow/NeedApprove

# 是否启用自动批准
auto_approve_enabled = True

# 发送模式：auto=自动发送, draft=手动确认
send_mode = draft

[EMAIL]
# 您的信息
from_name = Your Name  
from_email = your.email@company.com

# 批准回复内容
approval_message = Ref:MSG85395759
```

## 📧 工作流程

1. **邮件到达** - ServiceNow邮件到达Thunderbird
2. **拖拽到目标文件夹** - 将邮件拖到NeedApprove文件夹
3. **自动识别** - 程序识别需要批准的邮件
4. **生成回复** - 自动创建批准回复邮件
5. **发送处理** - 根据配置自动发送或等待确认

## 🎯 测试验证

程序已通过实际邮件测试：
```
✅ 发现邮件：RITM1603887 | Approval Required / Approbation Requise
✅ 发件人：lululemon ServiceNow <luluprod@service-now.com>  
✅ 识别为需要批准
✅ 生成回复成功
✅ AppleScript执行成功
```

## 💡 使用建议

### 🔒 安全模式（推荐新用户）
```ini
send_mode = draft
```
- 程序准备邮件，您检查后手动发送
- 确保内容正确后再发送

### ⚡ 效率模式（熟练用户）  
```ini
send_mode = auto
```
- 完全自动化，无需人工干预
- 适合大量重复的批准工作

## 🛠️ 故障排除

### AppleScript权限问题
如果遇到权限错误：
1. 打开系统设置 → 隐私与安全性 → 无障碍
2. 添加Terminal、Python或相关应用
3. 重启Terminal后重试

### 邮件发送不正确
1. 检查Thunderbird是否正在运行
2. 确认快捷键Cmd+Shift+M可以创建新邮件
3. 尝试使用`send_mode = draft`模式

### 找不到邮件文件夹
1. 运行`python detect_thunderbird_folders.py`检查
2. 确认在Thunderbird中创建了正确的文件夹结构
3. 检查`config.ini`中的路径设置

## 📊 程序监控

- **日志文件**: `email_auto_approve.log`
- **已处理记录**: `processed_emails.json`  
- **草稿邮件**: `*.eml`文件

## 🎉 总结

**程序现已完全可用！**

- ✨ **智能识别**：自动识别ServiceNow批准邮件
- 🚀 **灵活发送**：支持自动发送和手动确认
- 🔒 **安全可靠**：记录处理历史，避免重复
- 🛠️ **易于使用**：一键启动，自动监控

选择合适的发送模式，开始享受自动化的便利吧！