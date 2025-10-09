import sqlite3
import datetime

class DatabaseManager:
    """管理 SQLite 数据库连接和操作"""
    def __init__(self, db_name="receipts.db"):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self._setup_db()

    def _setup_db(self):
        """初始化数据库表结构"""
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS Receipts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                receipt_no TEXT UNIQUE NOT NULL,
                client_name TEXT NOT NULL,
                issue_date TEXT NOT NULL,
                total_amount_num REAL NOT NULL,
                total_amount_cap TEXT NOT NULL
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS ReceiptItems (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                receipt_id INTEGER,
                item_name TEXT,
                unit TEXT,
                quantity REAL,
                unit_price REAL,
                amount REAL,
                notes TEXT,
                FOREIGN KEY (receipt_id) REFERENCES Receipts(id)
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS Settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        self.conn.commit()

    def generate_receipt_no(self):
        """生成收据编号: 年月日 + 3位流水号"""
        date_str = datetime.datetime.now().strftime("%Y%m%d")
        self.cursor.execute("SELECT COUNT(*) FROM Receipts WHERE receipt_no LIKE ?", (date_str + '%',))
        count = self.cursor.fetchone()[0] + 1
        return f"{date_str}{count:03d}"

    def save_receipt(self, receipt_data, items_data):
        """保存一张完整的收据及其明细"""
        receipt_no = self.generate_receipt_no()
        issue_date = datetime.datetime.now().strftime("%Y-%m-%d")
        
        try:
            self.cursor.execute(
                "INSERT INTO Receipts (receipt_no, client_name, issue_date, total_amount_num, total_amount_cap) VALUES (?, ?, ?, ?, ?)",
                (receipt_no, receipt_data['client_name'], issue_date, receipt_data['total_amount_num'], receipt_data['total_amount_cap'])
            )
            
            # --- 错误修复：使用 self.cursor.lastrowid ---
            receipt_id = self.cursor.lastrowid
            
            for item in items_data:
                # item 是一个包含 6 个元素的列表/元组
                self.cursor.execute(
                    "INSERT INTO ReceiptItems (receipt_id, item_name, unit, quantity, unit_price, amount, notes) VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (receipt_id, item[0], item[1], item[2], item[3], item[4], item[5])
                )
            
            self.conn.commit()
            return receipt_no
        except Exception as e:
            # 简化错误处理
            print(f"Error saving receipt: {e}")
            return None

    def save_setting(self, key, value):
        """保存设置（如收款人信息）"""
        self.cursor.execute(
            "INSERT OR REPLACE INTO Settings (key, value) VALUES (?, ?)",
            (key, value)
        )
        self.conn.commit()

    def load_setting(self, key, default=""):
        """加载设置（如收款人信息）"""
        self.cursor.execute("SELECT value FROM Settings WHERE key = ?", (key,))
        result = self.cursor.fetchone()
        return result[0] if result else default

    def fetch_all_receipts(self, query=""):
        """查询所有收据或根据客户名称模糊查询"""
        if query:
            self.cursor.execute(
                "SELECT receipt_no, client_name, issue_date, total_amount_num FROM Receipts WHERE client_name LIKE ? OR receipt_no LIKE ? ORDER BY id DESC",
                ('%' + query + '%', '%' + query + '%')
            )
        else:
            self.cursor.execute("SELECT receipt_no, client_name, issue_date, total_amount_num FROM Receipts ORDER BY id DESC")
        return self.cursor.fetchall()

    def fetch_receipt_details(self, receipt_no):
        """根据收据编号获取详情，并附加最新的收款人设置信息"""
        self.cursor.execute("SELECT id, client_name, issue_date, total_amount_num, total_amount_cap FROM Receipts WHERE receipt_no = ?", (receipt_no,))
        receipt = self.cursor.fetchone()
        
        if not receipt:
            return None, None
            
        receipt_id, client_name, issue_date, total_amount_num, total_amount_cap = receipt
        
        self.cursor.execute(
            "SELECT item_name, unit, quantity, unit_price, amount, notes FROM ReceiptItems WHERE receipt_id = ?",
            (receipt_id,)
        )
        items = self.cursor.fetchall()
        
        # 附加最新的收款人设置信息
        payee = self.load_setting("payee", "")
        payee_company = self.load_setting("payee_company", "")
        
        return {
            'receipt_no': receipt_no,
            'client_name': client_name,
            'issue_date': issue_date,
            'total_amount_num': total_amount_num,
            'total_amount_cap': total_amount_cap,
            'payee': payee,
            'payee_company': payee_company
        }, items

    def close(self):
        self.conn.close()