---
name: outlook-pywin32
description: "通过pywin32本地操作Outlook的命令行工具。支持发送邮件、列出邮件、读取邮件、搜索邮件等。使用命令: outlook-pywin32.py <方法名> --参数 值"
---

# Outlook PyWin32 命令行工具

基于pywin32的Outlook本地自动化工具，无需OAuth，直接操作本地Outlook。

## 前提条件

- Windows系统
- 已安装Outlook客户端
- Python + pywin32 (`pip install pywin32`)

## 用法

```bash
python scripts/outlook-pywin32.py <方法名> --参数 值 --参数2 值2
```

## 可用方法

| 方法 | 说明 | 参数 |
|------|------|------|
| mail-send | 创建邮件并保存到草稿箱 | --to, --subject, --body, --cc, --bcc |
| mail-list | 列出邮件 | --folder, --limit, --account |
| mail-read | 读取邮件 | --folder, --index, --account |
| mail-search | 搜索邮件 | --query, --limit, --account, --start-time, --end-time |
| account-list | 列出所有可用的Outlook账户 | 无 |

## 参数说明

### 通用参数

- `--account`: 邮箱账户地址（可选）
  - 优先级：1. 传入参数 2. 环境变量 OUTLOOK_ACCOUNT 3. config.json 文件

### mail-search 特有参数

- `--query`: 搜索关键词（可选）
- `--start-time`: 起始时间（可选，如 2024-01-01 或 2024-01-01 09:00:00）
  - 只指定日期时，默认设置为 00:00:00
- `--end-time`: 结束时间（可选，如 2024-12-31 或 2024-12-31 18:00:00）
  - 只指定日期时，默认设置为 23:59:59

## 配置文件

在 scripts 目录下创建 `config.json` 文件，可配置默认邮箱账户：

```json
{
  "outlook_account": "xx@cuhk.edu.cn"
}
```

## 环境变量

- `OUTLOOK_ACCOUNT`: 默认邮箱账户地址

## 示例

```bash
# 列出可用的Outlook账户
python scripts/outlook-pywin32.py account-list

# 发送邮件
python scripts/outlook-pywin32.py mail-send --to user@example.com --subject "测试" --body "你好"

# 列出收件箱前10封邮件
python scripts/outlook-pywin32.py mail-list --folder inbox --limit 10

# 列出指定账户的邮件
python scripts/outlook-pywin32.py mail-list --account xx@cuhk.edu.cn

# 读取第1封邮件
python scripts/outlook-pywin32.py mail-read --folder inbox --index 1

# 只按关键词搜索邮件
python scripts/outlook-pywin32.py mail-search --query "发票"

# 只按时间范围搜索邮件
python scripts/outlook-pywin32.py mail-search --start-time 2024-01-01 --end-time 2024-12-31

# 同时使用关键词和时间范围搜索
python scripts/outlook-pywin32.py mail-search --query "会议" --start-time 2024-01-01 --end-time 2024-12-31

# 使用config.json中的账户搜索
python scripts/outlook-pywin32.py mail-search --query "会议"
```

## 添加新方法

在 `scripts/outlook-pywin32.py` 中:

1. 定义新函数 `mail_xxx(参数...)`
2. 添加到 `METHODS` 字典: `"mail-xxx": mail_xxx`

参数自动从函数签名解析，使用argparse处理命令行。
