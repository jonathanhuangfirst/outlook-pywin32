from .utils import get_outlook_app, get_namespace
import json
import os


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


def config():
    """
    编辑 config.json 文件的内容
    """
    # 获取 config.json 文件路径
    config_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config.json')
    
    # 读取现有配置
    with open(config_path, 'r', encoding='utf-8') as f:
        config_data = json.load(f)
    
    print("当前配置:")
    for key, value in config_data.items():
        print(f"  {key}: {value}")
    
    # 提示用户输入新的配置
    print("\n请输入新的配置值 (按 Enter 保留当前值):")
    
    # 处理 outlook_account 配置
    current_account = config_data.get('outlook_account', '')
    new_account = input(f"  outlook_account [{current_account}]: ").strip()
    if new_account:
        config_data['outlook_account'] = new_account
    
    # 保存配置到文件
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config_data, f, indent=2, ensure_ascii=False)
    
    print("\n配置已更新成功!")
    print("新配置:")
    for key, value in config_data.items():
        print(f"  {key}: {value}")
    
    return config_data
