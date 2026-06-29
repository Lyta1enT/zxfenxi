"""字段抽取结果预览面板"""
from typing import Dict, Any

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QTableWidget,
    QTableWidgetItem, QHeaderView
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QBrush


FIELD_LABELS = {
    'personal': {
        'name': '姓名', 'id_number': '证件号码', 'report_time': '报告时间',
        'credit_card_count': '信用卡账户数', 'loan_count': '贷款账户数',
        'overdue_count': '逾期账户数', 'total_balance': '余额',
        'settled_count': '已结清账户数', 'anomaly_notes': '异常备注',
    },
    'corporate': {
        'company_name': '企业名称', 'credit_code': '统一社会信用代码',
        'report_time': '报告时间', 'unsettled_institutions': '未结清机构数',
        'total_balance': '余额', 'short_term_loan': '短期借款',
        'medium_long_term_loan': '中长期借款', 'guarantee_info': '担保信息',
        'public_info': '公共信息',
    },
    'tax': {
        'tax_registration': '纳税登记状态', 'has_penalty': '是否有滞纳金',
        'tax_arrears': '欠税金额', 'invoice_3year': '近三年开票汇总',
        'tax_revenue_3year': '近三年纳税数据', 'tax_anomaly': '税务异常说明',
    },
}


class PreviewPanel(QWidget):
    """结果预览面板"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_type = 'personal'
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        title = QLabel('\U0001f4cb 字段抽取结果')
        title.setStyleSheet("font-size: 14px; font-weight: bold; color: #333; padding: 4px 0;")
        layout.addWidget(title)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(['字段名称', '字段值', '置信度', '状态'])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                border-radius: 4px;
                font-size: 12px;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 6px;
                border: none;
                border-bottom: 1px solid #ddd;
                font-weight: bold;
            }
        """)

        layout.addWidget(self.table)

        self.status_label = QLabel('等待处理...')
        self.status_label.setStyleSheet("color: #999; font-size: 11px; padding: 4px 0;")
        layout.addWidget(self.status_label)

        self.setLayout(layout)

    def show_fields(self, fields: Dict[str, Any], report_type: str):
        """显示字段抽取结果"""
        self._current_type = report_type
        labels = FIELD_LABELS.get(report_type, {})

        visible_fields = [(k, v) for k, v in fields.items() if k in labels]
        self.table.setRowCount(len(visible_fields))

        for row, (key, data) in enumerate(visible_fields):
            label = labels.get(key, key)
            value = data.get('value', '')
            confidence = data.get('confidence', 0)
            note = data.get('note', '')

            # 字段名
            name_item = QTableWidgetItem(label)
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 0, name_item)

            # 字段值
            display = value if value else '[未识别]'
            value_item = QTableWidgetItem(display)
            value_item.setFlags(value_item.flags() & ~Qt.ItemIsEditable)

            if not value:
                value_item.setBackground(QBrush(QColor(0xFF, 0xF2, 0xCC)))
            elif '\u26a0\ufe0f' in str(value):
                value_item.setBackground(QBrush(QColor(0xFC, 0xE4, 0xEC)))

            self.table.setItem(row, 1, value_item)

            # 置信度
            conf_text = f'{confidence:.0%}' if isinstance(confidence, float) else str(confidence)
            conf_item = QTableWidgetItem(conf_text)
            conf_item.setFlags(conf_item.flags() & ~Qt.ItemIsEditable)

            if isinstance(confidence, float) and confidence < 0.6:
                conf_item.setBackground(QBrush(QColor(0xFC, 0xE4, 0xEC)))
            elif isinstance(confidence, float) and confidence < 0.9:
                conf_item.setBackground(QBrush(QColor(0xFF, 0xF2, 0xCC)))

            self.table.setItem(row, 2, conf_item)

            # 状态
            status_text = note if note else ('\u2705 已识别' if value else '\u23f3 待识别')
            status_item = QTableWidgetItem(status_text)
            status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)

            if note == '未识别':
                status_item.setForeground(QBrush(QColor(0xCC, 0x66, 0x00)))
            elif '\u26a0\ufe0f' in str(value):
                status_item.setForeground(QBrush(QColor(0xCC, 0x33, 0x00)))

            self.table.setItem(row, 3, status_item)

    def clear(self):
        """清空预览"""
        self.table.setRowCount(0)
        self.status_label.setText('等待处理...')

    def set_status(self, text: str):
        self.status_label.setText(text)
