import sys
from datetime import datetime, date

from PyQt6 import QtCore, QtGui, QtWidgets
import openpyxl
from openpyxl.utils.datetime import from_excel


class ExcelViewer(QtWidgets.QMainWindow):
    def __init__(self, excel_path: str):
        super().__init__()
        self.setWindowTitle('RQC Data - Desktop Viewer')
        self.resize(1000, 700)

        # Set default font for better visibility
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(10)
        self.setFont(font)

        self.setStyleSheet("""
            QMainWindow {
                background-color: #2b2b2b;
                color: #ffffff;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #4a90e2;
                border-radius: 5px;
                margin-top: 1ex;
                color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #ffffff;
            }
            QPushButton {
                background-color: #4a90e2;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #357abd;
            }
            QPushButton:pressed {
                background-color: #2c5aa0;
            }
            QLineEdit, QDateEdit {
                border: 1px solid #555;
                border-radius: 4px;
                padding: 4px;
                background-color: #3a3a3a;
                color: #ffffff;
            }
            QCheckBox {
                color: #ffffff;
                font-size: 14px;
            }
            QLabel {
                color: #ffffff;
            }
            QTableWidget {
                gridline-color: #555;
                background-color: #3a3a3a;
                alternate-background-color: #454545;
                color: #ffffff;
            }
            QHeaderView::section {
                background-color: #4a90e2;
                color: white;
                padding: 4px;
                border: 1px solid #555;
            }
        """)

        self.excel_path = excel_path
        self.table = QtWidgets.QTableWidget()
        self.search_date = QtWidgets.QDateEdit(calendarPopup=True)
        self.date_filter_checkbox = QtWidgets.QCheckBox('Apply date filter')
        self.search_job = QtWidgets.QLineEdit()
        self.job_filter_checkbox = QtWidgets.QCheckBox('Apply job filter')

        self.search_date.setDisplayFormat('yyyy-MM-dd')
        self.search_date.setDate(QtCore.QDate.currentDate())

        self._setup_ui()
        self.load_data()

    def _setup_ui(self):
        self.stack = QtWidgets.QStackedWidget()
        self.setCentralWidget(self.stack)

        self.home_page = self._create_home_page()
        self.viewer_page = self._create_viewer_page()

        self.stack.addWidget(self.home_page)
        self.stack.addWidget(self.viewer_page)
        self.stack.setCurrentWidget(self.home_page)

    def _create_home_page(self):
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)

        header = QtWidgets.QLabel('RQC Data Desktop Viewer')
        header.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        header.setStyleSheet('font-size: 36px; font-weight: bold; padding: 20px; color: #0d47a1;')
        layout.addWidget(header)

        sub = QtWidgets.QLabel('Browse your RQC workbook, filter by date or job name, then review results in an interactive table.')
        sub.setWordWrap(True)
        sub.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        sub.setStyleSheet('font-size: 16px; color: #1976d2; padding: 0 80px 20px 80px;')
        layout.addWidget(sub)

        start_btn = QtWidgets.QPushButton('Open Viewer')
        start_btn.setFixedSize(240, 60)
        start_btn.setStyleSheet('font-size: 18px; background-color: #4caf50; color: white; border-radius: 10px;')
        start_btn.clicked.connect(lambda: self.stack.setCurrentWidget(self.viewer_page))
        layout.addWidget(start_btn, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)

        layout.addStretch()
        return page

    def _create_viewer_page(self):
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)

        header = QtWidgets.QLabel('RQC Data Viewer')
        header.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        header.setStyleSheet('font-size: 28px; font-weight: bold; padding: 10px; color: #0d47a1;')
        layout.addWidget(header)

        subtitle = QtWidgets.QLabel('Use the filters below to search by date and/or job name')
        subtitle.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet('color: #1976d2; padding-bottom: 12px; font-size: 14px;')
        layout.addWidget(subtitle)

        form_group = QtWidgets.QGroupBox('Filters')
        form_layout = QtWidgets.QFormLayout()
        form_layout.addRow(self.date_filter_checkbox, self.search_date)
        form_layout.addRow(self.job_filter_checkbox, self.search_job)
        form_group.setLayout(form_layout)
        layout.addWidget(form_group)

        button_layout = QtWidgets.QHBoxLayout()
        self.search_button = QtWidgets.QPushButton('Search')
        self.reset_button = QtWidgets.QPushButton('Show All')
        self.back_button = QtWidgets.QPushButton('Back to Home')
        self.search_button.setStyleSheet('background-color: #ff9800; color: white;')
        self.reset_button.setStyleSheet('background-color: #9c27b0; color: white;')
        self.back_button.setStyleSheet('background-color: #f44336; color: white;')
        button_layout.addWidget(self.search_button)
        button_layout.addWidget(self.reset_button)
        button_layout.addWidget(self.back_button)
        layout.addLayout(button_layout)

        self.search_job.textChanged.connect(self.on_search)
        self.search_date.dateChanged.connect(self.on_search)
        self.date_filter_checkbox.stateChanged.connect(self.on_search)
        self.job_filter_checkbox.stateChanged.connect(self.on_search)

        self.notice = QtWidgets.QLabel('')
        self.notice.setStyleSheet('padding: 8px; font-weight: bold; color: #388e3c;')
        layout.addWidget(self.notice)

        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self.table)
        layout.addWidget(scroll)

        return page
    def load_data(self):
        try:
            wb = openpyxl.load_workbook(self.excel_path, data_only=True)
            sheet = wb.active
            self.rows = [tuple(row) for row in sheet.iter_rows(values_only=True)]
            self.headers = self.rows[0] if self.rows else []
            self.data_rows = self.rows[1:] if len(self.rows) > 1 else []
            self._populate_table(self.data_rows)
        except Exception as e:
            self.notice.setText(f'Error loading data: {str(e)}')
            self.headers = []
            self.data_rows = []
            self._populate_table([])

    def _populate_table(self, rows):
        self.table.clear()
        if not rows:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            return

        self.table.setColumnCount(len(rows[0]))
        self.table.setRowCount(len(rows))
        self.table.setHorizontalHeaderLabels([str(h) if h is not None else '' for h in self.headers])

        for r, row in enumerate(rows):
            for c, value in enumerate(row):
                item = QtWidgets.QTableWidgetItem(str(value) if value is not None else '')
                self.table.setItem(r, c, item)

        self.table.resizeColumnsToContents()
        self.table.resizeRowsToContents()

        self.notice.setText(f'Showing {len(rows)} rows')

    def _cell_matches_date(self, cell, search_dt):
        if cell is None:
            return False

        if isinstance(cell, datetime):
            return cell.date() == search_dt

        if isinstance(cell, date):
            return cell == search_dt

        # Excel stores dates as floats; try converting
        if isinstance(cell, (int, float)):
            try:
                dt = from_excel(cell)
                return dt.date() == search_dt
            except Exception:
                pass

        s = str(cell).strip()
        if not s:
            return False

        # Check common date formats
        for fmt in ('%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y', '%Y/%m/%d'):
            try:
                if datetime.strptime(s, fmt).date() == search_dt:
                    return True
            except Exception:
                pass

        # Fallback: substring match against ISO format
        return search_dt.strftime('%Y-%m-%d') in s

    def on_search(self):
        date_text = self.search_date.date().toString('yyyy-MM-dd')
        job_text = self.search_job.text().strip().lower()
        use_date_filter = self.date_filter_checkbox.isChecked()
        use_job_filter = self.job_filter_checkbox.isChecked()

        if not use_date_filter and not use_job_filter:
            self.notice.setText('No filter enabled: showing all data')
            self._populate_table(self.data_rows)
            return

        search_dt = None
        if use_date_filter and date_text:
            try:
                search_dt = datetime.fromisoformat(date_text).date()
            except ValueError:
                search_dt = None

        def matches(row):
            date_ok = True
            job_ok = True

            if use_date_filter and search_dt:
                date_ok = any(self._cell_matches_date(cell, search_dt) for cell in row)

            if use_job_filter and job_text:
                job_ok = any(job_text in str(cell).lower() for cell in row if cell is not None)

            return date_ok and job_ok

        filtered = [r for r in self.data_rows if matches(r)]
        self._populate_table(filtered)
        self.notice.setText(f'Showing {len(filtered)} rows (filtered)')

    def on_reset(self):
        self.search_job.clear()
        self.date_filter_checkbox.setChecked(False)
        self.job_filter_checkbox.setChecked(False)
        self.notice.setText('')
        self.load_data()


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = ExcelViewer(r"D:\study material\dhqp\LotInfo.xlsx")
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
