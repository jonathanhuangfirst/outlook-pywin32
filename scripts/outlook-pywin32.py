#!/usr/bin/env python3
"""
Outlook PyWin32 命令行工具
用法: outlook-pywin32.py <方法名> --参数1 值1 --参数2 值2 ...
"""

import argparse
import datetime
import json
import os
import sys
import win32com.client


def get_outlook_app():
    """获取Outlook应用对象"""
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        return outlook
    except Exception as e:
        print(f"错误: 无法连接Outlook - {e}")
        sys.exit(1)


def get_namespace(outlook):
    """获取MAPI命名空间"""
    return outlook.GetNamespace("MAPI")


def get_account(account: str = None):
    """
    获取邮箱账户，优先级：
    1. 传入的 account 参数
    2. 环境变量 OUTLOOK_ACCOUNT
    3. 同目录下的 config.json 文件
    """
    if account:
        return account
    
    # 从环境变量获取
    env_account = os.environ.get("OUTLOOK_ACCOUNT")
    if env_account:
        return env_account
    
    # 从 config.json 文件获取
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
                return config.get("outlook_account")
        except Exception:
            pass
    
    return None


def get_mail_folder(namespace, folder_name: str = "inbox", account_email: str = None):
    """
    获取指定邮箱账户的文件夹
    
    参数:
        namespace: MAPI命名空间
        folder_name: 文件夹名称
        account_email: 邮箱账户地址，为空则使用默认账户
    """
    folder_map = {
        "inbox": 6,
        "sentitems": 5,
        "drafts": 16,
        "deleteditems": 3,
        "outbox": 4,
    }
    
    folder_id = folder_map.get(folder_name.lower(), 6)
    
    if account_email:
        for account in namespace.Accounts:
            if account.SmtpAddress.lower() == account_email.lower():
                store = account.DeliveryStore
                return store.GetDefaultFolder(folder_id)
        raise Exception(f"未找到邮箱账户: {account_email}")
    else:
        return namespace.GetDefaultFolder(folder_id)


# ============ Mail 方法集合 ============

def mail_send(to: str, subject: str, body: str = "", cc: str = "", bcc: str = ""):
    """
    创建邮件并保存到草稿箱

    参数:
        --to: 收件人邮箱 (必需)
        --subject: 邮件主题 (必需)
        --body: 邮件正文
        --cc: 抄送
        --bcc: 密送
    """
    outlook = get_outlook_app()
    mail = outlook.CreateItem(0)  # 0 = MailItem

    mail.To = to
    mail.Subject = subject
    if body:
        mail.Body = body
    if cc:
        mail.CC = cc
    if bcc:
        mail.BCC = bcc

    mail.Save()
    print(f"邮件已保存到草稿箱 -> {to}: {subject}")
    return {"success": True, "to": to, "subject": subject, "saved_to": "drafts"}


def mail_list(folder: str = "inbox", limit: int = 10, account: str = None):
    """
    列出文件夹中的邮件

    参数:
        --folder: 文件夹名称 (inbox/sentitems/drafts, 默认inbox)
        --limit: 返回数量 (默认10)
        --account: 邮箱账户地址，优先级：1. 传入参数 2. 环境变量 OUTLOOK_ACCOUNT 3. config.json 文件
    """
    outlook = get_outlook_app()
    namespace = get_namespace(outlook)
    account = get_account(account)
    mail_folder = get_mail_folder(namespace, folder, account)

    messages = mail_folder.Items
    messages.Sort("[ReceivedTime]", True)  # 按时间倒序

    results = []
    for i, msg in enumerate(messages):
        if i >= limit:
            break
        results.append({
            "index": i + 1,
            "subject": msg.Subject,
            "sender": msg.SenderName if hasattr(msg, "SenderName") else "N/A",
            "received": str(msg.ReceivedTime) if hasattr(msg, "ReceivedTime") else "N/A",
            "unread": not msg.UnRead if hasattr(msg, "UnRead") else False,
        })

    for r in results:
        status = "📩" if r["unread"] else "📭"
        print(f"{status} [{r['index']}] {r['subject']} - {r['sender']} ({r['received']})")

    return results


def mail_read(folder: str = "inbox", index: int = 1, account: str = None):
    """
    读取指定邮件内容

    参数:
        --folder: 文件夹名称 (默认inbox)
        --index: 邮件索引 (默认1, 从mail-list获取)
        --account: 邮箱账户地址，优先级：1. 传入参数 2. 环境变量 OUTLOOK_ACCOUNT 3. config.json 文件
    """
    outlook = get_outlook_app()
    namespace = get_namespace(outlook)
    account = get_account(account)
    mail_folder = get_mail_folder(namespace, folder, account)

    messages = mail_folder.Items
    messages.Sort("[ReceivedTime]", True)

    if index < 1 or index > messages.Count:
        print(f"错误: 索引 {index} 超出范围 (1-{messages.Count})")
        return None

    msg = messages(index)

    result = {
        "subject": msg.Subject,
        "sender": msg.SenderName if hasattr(msg, "SenderName") else "N/A",
        "sender_email": msg.SenderEmailAddress if hasattr(msg, "SenderEmailAddress") else "N/A",
        "to": msg.To if hasattr(msg, "To") else "N/A",
        "received": str(msg.ReceivedTime) if hasattr(msg, "ReceivedTime") else "N/A",
        "body": msg.Body if hasattr(msg, "Body") else "",
    }

    print(f"主题: {result['subject']}")
    print(f"发件人: {result['sender']} <{result['sender_email']}>")
    print(f"收件人: {result['to']}")
    print(f"时间: {result['received']}")
    print("-" * 50)
    print(result["body"])

    # 标记为已读
    if hasattr(msg, "UnRead"):
        msg.UnRead = False

    return result


def parse_date_for_outlook(date_str: str, is_start: bool = True):
    """
    解析日期字符串并转换为Outlook Restrict方法需要的格式
    
    参数:
        date_str: 日期字符串
        is_start: 是否为起始时间，True则默认00:00:00，False则默认23:59:59
    """
    if not date_str:
        return None
    
    # 带时间的格式
    time_formats = [
        "%Y-%m-%d %H:%M:%S",
        "%Y/%m/%d %H:%M:%S",
    ]
    
    # 只带日期的格式
    date_only_formats = [
        "%Y-%m-%d",
        "%Y/%m/%d",
    ]
    
    dt = None
    has_time = False
    
    # 先尝试带时间的格式
    for fmt in time_formats:
        try:
            dt = datetime.datetime.strptime(date_str, fmt)
            has_time = True
            break
        except ValueError:
            continue
    
    # 如果没有带时间，尝试只带日期的格式
    if not dt:
        for fmt in date_only_formats:
            try:
                dt = datetime.datetime.strptime(date_str, fmt)
                has_time = False
                break
            except ValueError:
                continue
    
    if not dt:
        return None
    
    # 如果没有指定时间，设置默认时间
    if not has_time:
        if is_start:
            dt = datetime.datetime.combine(dt.date(), datetime.datetime.min.time())
        else:
            dt = datetime.datetime.combine(dt.date(), datetime.datetime.max.time())
    
    # 转换为Outlook Restrict需要的格式: mm/dd/yyyy HH:mm:ss (24小时制)
    return dt.strftime("%m/%d/%Y %H:%M:%S")


def mail_search(query: str = "", limit: int = 50, account: str = None, start_time: str = None, end_time: str = None):
    """
    搜索邮件

    参数:
        --query: 搜索关键词 (可选)
        --limit: 返回数量 (默认50)
        --account: 邮箱账户地址，优先级：1. 传入参数 2. 环境变量 OUTLOOK_ACCOUNT 3. config.json 文件
        --start-time: 起始时间 (如 2024-01-01 或 2024-01-01 09:00:00)
        --end-time: 结束时间 (如 2024-12-31 或 2024-12-31 18:00:00)
    """
    outlook = get_outlook_app()
    namespace = get_namespace(outlook)

    account = get_account(account)
    inbox = get_mail_folder(namespace, "inbox", account)
    messages = inbox.Items

    # 使用Outlook SQL查询语法
    sql_parts = []
    
    # 关键词搜索
    if query:
        # 转义单引号
        escaped_query = query.replace("'", "''")
        sql_parts.append(f"(urn:schemas:httpmail:subject LIKE '%{escaped_query}%' OR urn:schemas:httpmail:textdescription LIKE '%{escaped_query}%')")
    
    # 时间范围搜索
    if start_time or end_time:
        start_date_outlook = parse_date_for_outlook(start_time, is_start=True)
        end_date_outlook = parse_date_for_outlook(end_time, is_start=False)
        
        if start_date_outlook:
            sql_parts.append(f"urn:schemas:httpmail:datereceived >= '{start_date_outlook}'")
        if end_date_outlook:
            sql_parts.append(f"urn:schemas:httpmail:datereceived <= '{end_date_outlook}'")
    
    # 构建完整的SQL查询
    filtered_messages = messages
    if sql_parts:
        sql_criteria = " AND ".join(sql_parts)
        full_query = f"@SQL={sql_criteria}"
        try:
            #print(f"使用SQL查询: {full_query}")
            filtered_messages = messages.Restrict(full_query)
        except Exception as e:
            print(f"警告: SQL查询失败，将使用所有邮件 - {e}")
            filtered_messages = messages

    results = []
    for msg in filtered_messages:
        if len(results) >= limit:
            break
        
        results.append({
            "subject": msg.Subject if hasattr(msg, "Subject") else "",
            "sender": msg.SenderName if hasattr(msg, "SenderName") else "",
            "received": str(msg.ReceivedTime) if hasattr(msg, "ReceivedTime") else "N/A",
        })

    for i, r in enumerate(results):
        print(f"[{i + 1}] {r['subject']} - {r['sender']} ({r['received']})")

    return results


def account_list():
    """
    列出所有可用的 Outlook 邮箱账户
    """
    outlook = get_outlook_app()
    namespace = get_namespace(outlook)
    
    accounts = []
    for account in namespace.Accounts:
        accounts.append({
            "name": account.DisplayName,
            "email": account.SmtpAddress,
        })
    
    print("可用的 Outlook 账户:")
    for i, acc in enumerate(accounts):
        print(f"  [{i + 1}] {acc['name']} <{acc['email']}>")
    
    return accounts


# ============ 方法注册表 ============

METHODS = {
    "mail-send": mail_send,
    "mail-list": mail_list,
    "mail-read": mail_read,
    "mail-search": mail_search,
    "account-list": account_list,
}


def parse_args():
    """解析命令行参数"""
    if len(sys.argv) < 2:
        print("用法: outlook-pywin32.py <方法名> --参数 值 ...")
        print(f"可用方法: {', '.join(METHODS.keys())}")
        sys.exit(1)

    method_name = sys.argv[1].lower().replace("_", "-")

    if method_name not in METHODS:
        print(f"错误: 未知方法 '{method_name}'")
        print(f"可用方法: {', '.join(METHODS.keys())}")
        sys.exit(1)

    parser = argparse.ArgumentParser(description=f"Outlook {method_name}")
    parser.add_argument("_method", nargs="?", default=method_name, help=argparse.SUPPRESS)

    # 根据方法添加参数
    func = METHODS[method_name]
    import inspect
    sig = inspect.signature(func)

    for param_name, param in sig.parameters.items():
        if param_name == "self":
            continue
        arg_name = f"--{param_name.replace('_', '-')}"
        default = param.default if param.default != inspect.Parameter.empty else None
        # account、start_time、end_time 参数即使默认值是 None 也不需要是必需的
        required = default is None and param_name not in ("account", "start_time", "end_time")

        parser.add_argument(
            arg_name,
            dest=param_name,
            required=required,
            default=default,
            type=str if param.annotation == str else int if param.annotation == int else str,
            help=f"{param_name}"
        )

    args = parser.parse_args(sys.argv[1:])

    # 转换参数类型
    kwargs = {}
    for param_name, param in sig.parameters.items():
        if param_name == "self":
            continue
        value = getattr(args, param_name, None)
        if value is not None:
            # 类型转换
            if param.annotation == int:
                value = int(value)
            elif param.annotation == bool:
                value = value.lower() in ("true", "1", "yes")
            kwargs[param_name] = value

    return method_name, kwargs


def main():
    """主入口"""
    method_name, kwargs = parse_args()
    func = METHODS[method_name]

    try:
        result = func(**kwargs)
        return result
    except Exception as e:
        print(f"执行错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
