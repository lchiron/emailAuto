# ServiceNow 邮件自动批准程序 - 最终版本

## 🎉 功能说明

本程序现已完美支持Thunderbird邮件自动批准功能：

✅ **正确识别Thunderbird程序内的文件夹路径**（不是系统文件路径）  
✅ **解析Thunderbird的mbox格式邮件文件**  
✅ **智能识别需要批准的ServiceNow邮件**  
✅ **自动生成标准格式的批准回复邮件**  
✅ **完整的监控和处理流程**

## 🚀 快速开始

### 1. 在Thunderbird中设置文件夹结构

**重要：这是Thunderbird程序内的邮件文件夹，不是系统文件路径！**

在Thunderbird中创建以下文件夹层级：

```
本地文件夹 (Local Folders)
└── 📁 Archive
    └── 📁 ServiceNow  
        └── 📁 NeedApprove  ← 程序监控此文件夹
```

**详细步骤：**
1. 打开Thunderbird
2. 右键点击"本地文件夹" → "新建文件夹" → 输入`Archive`
3. 右键点击"Archive" → "新建文件夹" → 输入`ServiceNow`  
4. 右键点击"ServiceNow" → "新建文件夹" → 输入`NeedApprove`

### 2. 验证文件夹设置

运行检测工具确认设置正确：
```bash
python detect_thunderbird_folders.py
```

应该显示：✅ 找到目标文件夹

### 3. 启动程序

```bash
./start.sh
```

程序将：
- 自动检测Thunderbird文件夹设置
- 处理现有邮件
- 持续监控新邮件
- 自动发送批准回复

## 📧 邮件处理流程

1. **邮件识别**：程序识别主题包含`approve`、`approval`、`ritm`、`servicenow`关键词的邮件
2. **解析处理**：从Thunderbird的mbox文件中解析邮件内容
3. **自动回复**：生成标准格式的批准回复邮件
4. **发送处理**：
   - macOS：通过AppleScript自动控制Thunderbird发送
   - 其他系统：生成.eml草稿文件供手动发送
5. **记录管理**：记录已处理邮件，避免重复处理

## ⚙️ 技术特点

- **智能mbox解析**：完美处理Thunderbird的邮件存储格式
- **Mozilla头部过滤**：跳过Thunderbird特有的元数据
- **文件夹自动发现**：支持多种Thunderbird配置文件结构
- **实时文件监控**：使用watchdog监控邮件文件变化
- **错误恢复**：完善的异常处理和日志记录

## 📁 项目文件

- `email_auto_approve.py` - 主程序（支持mbox格式）
- `detect_thunderbird_folders.py` - 文件夹检测工具  
- `thunderbird_sender.py` - 邮件发送模块
- `config.ini` - 配置文件
- `start.sh` - 一键启动脚本
- `processed_emails.json` - 已处理邮件记录（自动生成）

## 🔧 配置说明

编辑`config.ini`：

```ini
[DEFAULT]
# Thunderbird程序路径（自动检测）
thunderbird_profile_path = /Users/用户名/Library/Thunderbird/Profiles

# 监控的邮件文件夹路径（Thunderbird程序内的结构）
# Archive/ServiceNow/NeedApprove 表示：Archive > ServiceNow > NeedApprove
watch_folder = Archive/ServiceNow/NeedApprove

# 是否启用自动批准
auto_approve_enabled = True

[EMAIL]  
# 发件人信息
from_name = Your Name
from_email = your.email@company.com

# 批准回复内容
approval_message = Ref:MSG85395759
```

## ✅ 测试确认

程序已通过完整测试：

- ✅ Thunderbird文件夹路径识别
- ✅ mbox邮件文件解析
- ✅ ServiceNow邮件识别  
- ✅ 批准回复生成
- ✅ 文件监控功能
- ✅ 已处理邮件管理

## 💡 使用提示

1. **邮件管理**：将需要批准的邮件拖拽到NeedApprove文件夹
2. **规则设置**：可设置Thunderbird过滤规则自动移动ServiceNow邮件
3. **日志查看**：运行日志保存在`email_auto_approve.log`
4. **手动控制**：可随时修改`config.ini`禁用自动批准

## 🎯 程序优势

- **完全自动化**：无需手动干预的邮件批准流程
- **安全可靠**：记录所有操作，避免重复处理
- **高度兼容**：完美支持Thunderbird的mbox格式
- **用户友好**：详细的设置指南和错误提示
- **跨平台**：支持macOS、Windows、Linux

---

🚀 **现在就开始使用吧！程序已经完全准备就绪。**