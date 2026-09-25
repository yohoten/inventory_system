import os
import shutil
import json
from datetime import datetime, timedelta
import logging
import hashlib

class SystemUtils:
    def __init__(self):
        # 设置日志
        self.setup_logging()
    
    def setup_logging(self):
        """设置操作日志"""
        logging.basicConfig(
            filename='operation_log.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
    
    def log_operation(self, user, operation, details=""):
        """记录操作日志"""
        log_message = f"用户:{user} 操作:{operation}"
        if details:
            log_message += f" 详情:{details}"
        logging.info(log_message)
    
    def generate_password_hash(self, password):
        """生成密码哈希值"""
        return hashlib.sha256(password.encode()).hexdigest()
    
    def verify_password(self, password, hashed_password):
        """验证密码"""
        return self.generate_password_hash(password) == hashed_password
    
    def backup_database(self, db_path='inventory.db'):
        """数据库备份功能"""
        try:
            backup_dir = "backups"
            if not os.path.exists(backup_dir):
                os.makedirs(backup_dir)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = f"{backup_dir}/inventory_backup_{timestamp}.db"
            
            shutil.copy2(db_path, backup_file)
            
            # 清理超过30天的备份文件
            self.cleanup_old_backups(backup_dir, days=30)
            
            return backup_file
        except Exception as e:
            logging.error(f"备份失败: {str(e)}")
            return None
    
    def cleanup_old_backups(self, backup_dir, days=30):
        """清理过期备份文件"""
        cutoff_date = datetime.now() - timedelta(days=days)
        
        for filename in os.listdir(backup_dir):
            if filename.startswith('inventory_backup_') and filename.endswith('.db'):
                file_path = os.path.join(backup_dir, filename)
                file_time = datetime.fromtimestamp(os.path.getctime(file_path))
                
                if file_time < cutoff_date:
                    os.remove(file_path)
                    logging.info(f"已删除过期备份: {filename}")
    
    def export_to_json(self, data, filename):
        """导出数据为JSON格式"""
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            logging.error(f"JSON导出失败: {str(e)}")
            return False
    
    def calculate_inventory_turnover(self, cursor):
        """计算库存周转率"""
        try:
            # 计算平均库存
            cursor.execute("SELECT AVG(quantity) as avg_inventory FROM inventory")
            avg_inventory = cursor.fetchone()[0] or 0
            
            # 计算期间销售
            cursor.execute("""
                SELECT SUM(quantity) as total_sales FROM sale 
                WHERE date >= date('now', '-30 days')
            """)
            total_sales = cursor.fetchone()[0] or 0
            
            if avg_inventory > 0:
                turnover_rate = total_sales / avg_inventory
                return {
                    'turnover_rate': round(turnover_rate, 2),
                    'avg_inventory': avg_inventory,
                    'period_sales': total_sales,
                    'period_days': 30
                }
            return None
        except Exception as e:
            logging.error(f"计算周转率失败: {str(e)}")
            return None
    
    def get_low_stock_items(self, cursor, threshold=10):
        """获取低库存商品"""
        try:
            cursor.execute("""
                SELECT product_name, specification, quantity, unit
                FROM inventory 
                WHERE quantity <= ?
                ORDER BY quantity ASC
            """, (threshold,))
            return cursor.fetchall()
        except Exception as e:
            logging.error(f"获取低库存商品失败: {str(e)}")
            return []
    
    def get_monthly_statistics(self, cursor):
        """获取月度统计数据"""
        try:
            # 月度进货统计
            cursor.execute("""
                SELECT strftime('%Y-%m', date) as month, 
                       SUM(quantity) as total_quantity,
                       COUNT(*) as record_count
                FROM purchase 
                GROUP BY strftime('%Y-%m', date)
                ORDER BY month DESC LIMIT 12
            """)
            purchase_stats = cursor.fetchall()
            
            # 月度出货统计
            cursor.execute("""
                SELECT strftime('%Y-%m', date) as month, 
                       SUM(quantity) as total_quantity,
                       COUNT(*) as record_count
                FROM sale 
                GROUP BY strftime('%Y-%m', date)
                ORDER BY month DESC LIMIT 12
            """)
            sale_stats = cursor.fetchall()
            
            return {
                'purchase': purchase_stats,
                'sale': sale_stats
            }
        except Exception as e:
            logging.error(f"获取月度统计失败: {str(e)}")
            return {'purchase': [], 'sale': []}

# 全局工具实例
utils = SystemUtils()

# 为了向后兼容，提供模块级别的函数
def generate_password_hash(password):
    """生成密码哈希值"""
    return utils.generate_password_hash(password)

def verify_password(password, hashed_password):
    """验证密码"""
    return utils.verify_password(password, hashed_password)

def log_operation(user, operation, details=""):
    """记录操作日志"""
    utils.log_operation(user, operation, details)

def backup_database(db_path='inventory.db'):
    """数据库备份功能"""
    return utils.backup_database(db_path)

def get_low_stock_items(cursor, threshold=10):
    """获取低库存商品"""
    return utils.get_low_stock_items(cursor, threshold)