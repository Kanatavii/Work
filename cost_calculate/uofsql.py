import sys
import pymysql
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QWidget, QLabel, QLineEdit, QPushButton, QHBoxLayout
)

# 数据库配置
DB_CONFIG = {
    "host": "dbmysql.uofexp.com",
    "port": 63006,
    "user": "uofreader",
    "password": "uV4ad*3k9@",
    "database": "uof"
}

# 固定的列头，去掉 "到港时间"
COLUMN_HEADERS = [
    "搬入时间", "许可时间", "入库时间", "出库入金时间", "转运日期", "转运单号", "回数", "FBA",
    "空运/海运", "取件地", "主单号", "送り状番号", "箱数", "重量", "立方数", "尺寸", "转运公司", "转运备注",
    "现场用 乐天匹配", "郵便番号", "会社名/个人", "荷受人住所", "荷受人電話", "乐天5位数", "担当者",
    "保管天数", "仓储费", "预算所用单号", "福山", "佐川", "TONAMI", "中村包车", "JHSS", "UOF混载",
    "UOF包车", "成本含税", "请求含税"
]

# 从数据库检索数据，根据 "送り状番号"
def fetch_data_by_tracking_number(tracking_number):
    try:
        connection = pymysql.connect(**DB_CONFIG)
        cursor = connection.cursor()

        # 查询语句：从不同表中检索数据，使用 "送り状番号" 进行匹配
        query = f"""
        SELECT 搬入时间, 许可时间, 入库时间, 出库指示入金确认时间, 转运日期, 转运单号, 回数, FBA,
               空运海运, 取件地, 主单号, 送り状番号, 箱数, 重量, 立方数, 尺寸, 转运公司, 转运备注,
               现场用乐天匹配, 郵便番号, 会社名, 荷受人住所, 荷受人電話, 乐天5位数, 担当者,
               保管天数, 仓储费, 预算所用单号, 福山, 佐川, TONAMI, 中村包车, JHSS, UOF混载,
               UOF包车, 成本含税, 请求含税
        FROM bsn_shipperconsignee_language
        WHERE 送り状番号 = %s
        LIMIT 1;
        """
        cursor.execute(query, (tracking_number,))
        row = cursor.fetchone()

        cursor.close()
        connection.close()
        return row

    except Exception as e:
        print(f"数据库错误: {e}")
        return None


# 主窗口类
class DataTableApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("数据可视化展示 - PyQt5")
        self.setGeometry(100, 100, 1500, 800)  # 设置窗口大小
        self.initUI()

    def initUI(self):
        # 主布局
        main_layout = QVBoxLayout()

        # 输入框和按钮布局
        input_layout = QHBoxLayout()
        input_label = QLabel("请输入 送り状番号:")
        self.input_field = QLineEdit()
        self.search_button = QPushButton("搜索")
        self.search_button.clicked.connect(self.search_data)

        input_layout.addWidget(input_label)
        input_layout.addWidget(self.input_field)
        input_layout.addWidget(self.search_button)

        main_layout.addLayout(input_layout)

        # 表格控件
        self.tableWidget = QTableWidget()
        main_layout.addWidget(self.tableWidget)

        # 设置表头
        self.tableWidget.setColumnCount(len(COLUMN_HEADERS))
        self.tableWidget.setHorizontalHeaderLabels(COLUMN_HEADERS)
        self.tableWidget.setRowCount(0)  # 初始没有行数据

        # 设置中心窗口
        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def search_data(self):
        # 获取输入框中的 送り状番号
        tracking_number = self.input_field.text().strip()
        if not tracking_number:
            return

        # 从数据库检索数据
        row_data = fetch_data_by_tracking_number(tracking_number)

        # 清空现有表格数据
        self.tableWidget.setRowCount(0)

        # 填充表格数据
        if row_data:
            self.tableWidget.setRowCount(1)
            for col_index, value in enumerate(row_data):
                if value is not None:  # 避免空值显示 "None"
                    self.tableWidget.setItem(0, col_index, QTableWidgetItem(str(value)))
        else:
            print("未找到对应的数据。")


# 运行应用
if __name__ == "__main__":
    app = QApplication(sys.argv)
    mainWindow = DataTableApp()
    mainWindow.show()
    sys.exit(app.exec_())
