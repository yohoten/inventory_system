import json
import os
from datetime import datetime
import hashlib

class SystemConfig:
    def __init__(self):
        self.config_file = 'system_config.json'
        self.default_config = {
            'users': {
                'admin': {
                    'password': self.hash_password('123456'),
                    'role': 'administrator',
                    'created_date': datetime.now().isoformat()
                }
            },
            'system_settings': {
                'company_name': '进销存管理系统',
                'backup_enabled': True,
                'auto_backup_days': 7,
                'low_stock_threshold': 10,
                'theme': 'light'
            },
            'permissions': {
                'administrator': ['all'],
                'operator': ['view', 'add_purchase', 'add_sale', 'export']
            }
        }
        self.load_config()
    
    def hash_password(self, password):
        """密码哈希加密"""
        return hashlib.sha256(password.encode()).hexdigest()
    
    def load_config(self):
        """加载配置文件"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
            except:
                self.config = self.default_config.copy()
                self.save_config()
        else:
            self.config = self.default_config.copy()
            self.save_config()
    
    def save_config(self):
        """保存配置文件"""
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)
    
    def authenticate_user(self, username, password):
        """用户认证"""
        if username in self.config['users']:
            stored_hash = self.config['users'][username]['password']
            return self.hash_password(password) == stored_hash
        return False
    
    def get_user_role(self, username):
        """获取用户角色"""
        if username in self.config['users']:
            return self.config['users'][username]['role']
        return None
    
    def add_user(self, username, password, role='operator'):
        """添加新用户"""
        if username not in self.config['users']:
            self.config['users'][username] = {
                'password': self.hash_password(password),
                'role': role,
                'created_date': datetime.now().isoformat()
            }
            self.save_config()
            return True
        return False
    
    def get_system_setting(self, key, default=None):
        """获取系统设置"""
        return self.config['system_settings'].get(key, default)
    
    def set_system_setting(self, key, value):
        """设置系统参数"""
        self.config['system_settings'][key] = value
        self.save_config()
    
    def check_permission(self, username, permission):
        """检查用户权限"""
        role = self.get_user_role(username)
        if not role:
            return False
        
        user_permissions = self.config['permissions'].get(role, [])
        return 'all' in user_permissions or permission in user_permissions

# 全局配置实例
config = SystemConfig()

# 为了向后兼容，提供模块级别的函数
def get_system_setting(key, default=None):
    """获取系统设置"""
    return config.get_system_setting(key, default)

def set_system_setting(key, value):
    """设置系统参数"""
    config.set_system_setting(key, value)

def authenticate_user(username, password):
    """用户认证"""
    return config.authenticate_user(username, password)

def check_permission(username, permission):
    """检查用户权限"""
    return config.check_permission(username, permission)