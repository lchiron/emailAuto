# ServiceNow 邮件自动批准程序使用说明

## 功能说明

本程序可以自动监控Thunderbird邮件目录下的Archive/ServiceNow/NeedApprove文件夹，识别需要批准的ServiceNow邮件，并自动发送批准回复。

## 安装步骤

### 1. 环境要求
- Python 3.6 或更高版本
- Thunderbird邮件客户端
- macOS系统（推荐，支持完整的自动发送功能）

### 2. 安装Python依赖
```bash
# 激活虚拟环境（如果使用的话）
source venv/bin/activate

# 安装依赖
pip install -r requirements.txt
```

## 配置Thunderbird

### 创建文件夹结构

**重要：这里的路径指的是Thunderbird程序内的邮件文件夹，不是系统文件路径。**

在Thunderbird中创建以下文件夹结构：

1. 打开Thunderbird
2. 在左侧邮件文件夹面板中，找到"本地文件夹"（Local Folders）
3. 右键点击"本地文件夹"，选择"新建文件夹"
4. 按以下层级创建文件夹：

```
本地文件夹 (Local Folders)
└── Archive                    ← 第1步：创建Archive文件夹
    └── ServiceNow            ← 第2步：在Archive下创建ServiceNow文件夹  
        └── NeedApprove       ← 第3步：在ServiceNow下创建NeedApprove文件夹
```

### 文件夹创建步骤详解：

**第1步：创建Archive文件夹**
- 右键点击"本地文件夹" → "新建文件夹"
- 文件夹名称：`Archive`
- 点击"确定"

**第2步：创建ServiceNow文件夹**
- 右键点击刚创建的"Archive"文件夹 → "新建文件夹"  
- 文件夹名称：`ServiceNow`
- 点击"确定"

**第3步：创建NeedApprove文件夹**
- 右键点击刚创建的"ServiceNow"文件夹 → "新建文件夹"
- 文件夹名称：`NeedApprove` 
- 点击"确定"

完成后，您应该看到这样的文件夹结构：
```
📁 本地文件夹
├── 📁 Trash
├── 📁 Unsent Messages  
└── 📁 Archive
    └── 📁 ServiceNow
        └── 📁 NeedApprove  ← 程序将监控这个文件夹
```

### 使用邮件文件夹

将需要批准的ServiceNow邮件移动或复制到`NeedApprove`文件夹中：

1. **手动移动**：选中邮件，拖拽到NeedApprove文件夹
2. **复制邮件**：右键邮件 → "复制到" → 选择NeedApprove文件夹
3. **邮件规则**：设置过滤规则自动将ServiceNow邮件移动到此文件夹

### 验证文件夹设置

运行文件夹检测工具来验证设置：
```bash
python detect_thunderbird_folders.py
```

如果设置正确，工具会显示找到目标文件夹的消息。

## 配置说明

编辑 `config.ini` 文件：

```ini
[DEFAULT]
# Thunderbird配置文件路径（会自动检测，通常无需修改）
thunderbird_profile_path = /Users/yourusername/Library/Thunderbird/Profiles

# 监控的邮件文件夹相对路径
watch_folder = Archive/ServiceNow/NeedApprove

# 已处理邮件记录文件
processed_emails = processed_emails.json

# 日志级别: DEBUG, INFO, WARNING, ERROR
log_level = INFO

# 是否启用自动批准
auto_approve_enabled = True

[EMAIL]
# 发件人信息
from_name = Your Name
from_email = your.email@company.com

# 批准回复内容
approval_message = Ref:MSG85395759
```

## 使用方法

### 1. 基本使用
```bash
python email_auto_approve.py
```

### 2. 测试邮件发送功能
```bash
python thunderbird_sender.py
```

### 3. 程序运行说明

程序启动后会：
1. 读取配置文件
2. 扫描现有的邮件文件并处理需要批准的邮件
3. 持续监控目标文件夹，自动处理新增的邮件
4. 生成日志文件 `email_auto_approve.log`

## 文件说明

- `email_auto_approve.py` - 主程序文件
- `thunderbird_sender.py` - 邮件发送工具
- `config.ini` - 配置文件
- `requirements.txt` - Python依赖列表
- `processed_emails.json` - 已处理邮件记录（自动生成）
- `email_auto_approve.log` - 程序运行日志（自动生成）

## 自动发送机制

### macOS系统
程序会通过AppleScript控制Thunderbird自动发送邮件：
1. 检查Thunderbird是否运行，如未运行则自动启动
2. 创建新邮件窗口
3. 自动填写收件人、主题和邮件内容
4. 自动发送邮件

### 其他系统
程序会生成邮件草稿文件（.eml格式），需要手动导入Thunderbird发送。

## 邮件识别规则

程序会识别以下类型的邮件进行自动批准：
- 主题包含 "approve", "approval", "ritm", "servicenow" 关键词的邮件

## 安全注意事项

1. 程序会记录已处理的邮件，避免重复处理
2. 可以通过配置文件禁用自动批准功能
3. 所有操作都会记录在日志文件中
4. 建议定期检查日志文件和处理结果

## 故障排除

### 1. 找不到Thunderbird配置路径
- 手动修改 `config.ini` 中的 `thunderbird_profile_path`
- macOS默认路径：`~/Library/Thunderbird/Profiles`
- Windows默认路径：`%APPDATA%\\Thunderbird\\Profiles`
- Linux默认路径：`~/.thunderbird`

### 2. 邮件文件夹不存在
- 在Thunderbird中手动创建对应的文件夹结构
- 确保路径 `Archive/ServiceNow/NeedApprove` 存在

### 3. 自动发送失败
- 确保Thunderbird已安装并可正常运行
- 检查系统权限设置（macOS需要授权AppleScript访问）
- 查看日志文件了解详细错误信息

### 4. 邮件未被识别
- 检查邮件主题是否包含识别关键词
- 查看日志文件了解处理详情
- 调整日志级别为DEBUG获取更多信息

## 停止程序

按 `Ctrl+C` 停止监控程序。

## 日志查看

```bash
# 查看实时日志
tail -f email_auto_approve.log

# 查看最近的日志
tail -n 50 email_auto_approve.log
```