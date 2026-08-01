"""
财务分析工具 v2.0 - GUI 入口（支持多年财报）
- 多文件添加，自动按年份归类
- 趋势分析：YoY 同比 + CAGR 复合增长率
- 多 Sheet Excel 报告 + openpyxl 原生图表
"""

import sys
import json
import traceback
from pathlib import Path
from datetime import datetime

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTextEdit, QProgressBar,
    QTabWidget, QMessageBox, QGroupBox, QListWidget, QListWidgetItem
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor

from analyzer import (
    extract_from_excel, extract_from_pdf,
    calculate_indicators, calculate_trends,
    generate_report, analyze_files, detect_year_from_filename,
    export_subjects_to_csv
)


# ============================================================
# 工作线程
# ============================================================

class Worker(QThread):
    finished = pyqtSignal(dict, dict, list, dict, list)  # data, indicators, years, trends, errors
    error = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(self, file_list):
        super().__init__()
        self.file_list = file_list

    def run(self):
        try:
            self.progress.emit(f"正在分析 {len(self.file_list)} 个文件...")
            data_by_year, indicators_by_year, years, trends, errors = analyze_files(
                self.file_list)

            if not years:
                self.error.emit(
                    "未能从任何财报中提取到有效数据！\n\n"
                    "可能原因:\n"
                    "1. PDF 是扫描件（需要 OCR）\n"
                    "2. Excel 文件结构不标准\n"
                    "3. 科目名称差异过大\n\n"
                    "提示: 文件名包含 4 位年份（如 2023年报.pdf）可自动识别。")
                return

            self.progress.emit(
                f"分析完成: {len(years)} 年 ({', '.join(map(str, years))}), "
                f"{len(errors)} 个错误")
            self.finished.emit(data_by_year, indicators_by_year, years, trends, errors)
        except Exception as e:
            self.error.emit(f"{str(e)}\n\n{traceback.format_exc()}")


# ============================================================
# 主窗口
# ============================================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("财务分析工具 v2.0 - 本地版")
        self.setGeometry(80, 80, 1200, 850)
        self.file_list = []  # 待分析文件列表
        self.data_by_year = None
        self.indicators_by_year = None
        self.years = None
        self.trends = None
        self.errors = None
        self.init_ui()

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setSpacing(10)

        # 标题
        title = QLabel("💼 财务分析工具 v2.0")
        title.setFont(QFont("Microsoft YaHei", 20, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #1F4E78; padding: 6px;")
        layout.addWidget(title)

        subtitle = QLabel("支持多年财报 · 趋势分析 · 自动图表 · 本地运行数据不外传")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: #888; padding-bottom: 4px;")
        layout.addWidget(subtitle)

        # 作者信息（固定署名）
        author_layout = QHBoxLayout()
        author_label = QLabel("👤  作者:")
        author_label.setStyleSheet(
            "color: #1F4E78; font-weight: bold; font-size: 14px;")
        author_layout.addStretch()
        author_layout.addWidget(author_label)
        author_name = QLabel("Jasmin Alpha Hunter")
        author_name.setStyleSheet(
            "color: #1F4E78; font-weight: bold; font-size: 16px; "
            "padding: 4px 12px; background: #F0F8FF; "
            "border: 1px solid #B0C4DE; border-radius: 4px;")
        author_layout.addWidget(author_name)
        author_layout.addStretch()
        layout.addLayout(author_layout)

        # ① 文件选择（多文件）
        file_group = QGroupBox("① 添加财报文件（可多个，建议文件名含年份如 2023年报.pdf）")
        file_layout = QVBoxLayout(file_group)

        btn_row = QHBoxLayout()
        btn_pdf = QPushButton("📄 添加 PDF 财报")
        btn_pdf.setStyleSheet("padding: 8px;")
        btn_pdf.clicked.connect(lambda: self.add_files("PDF"))
        btn_row.addWidget(btn_pdf)

        btn_excel = QPushButton("📊 添加 Excel 财报")
        btn_excel.setStyleSheet("padding: 8px;")
        btn_excel.clicked.connect(lambda: self.add_files("Excel"))
        btn_row.addWidget(btn_excel)

        btn_clear = QPushButton("🗑 清空列表")
        btn_clear.setStyleSheet("padding: 8px;")
        btn_clear.clicked.connect(self.clear_files)
        btn_row.addWidget(btn_clear)

        btn_row.addStretch()
        file_layout.addLayout(btn_row)

        self.file_list_widget = QListWidget()
        self.file_list_widget.setMaximumHeight(120)
        self.file_list_widget.setStyleSheet(
            "background: #FAFAFA; border: 1px solid #DDD;")
        file_layout.addWidget(self.file_list_widget)

        layout.addWidget(file_group)

        # ② 操作
        op_group = QGroupBox("② 操作")
        op_layout = QHBoxLayout(op_group)

        self.btn_analyze = QPushButton("🚀 开始分析")
        self.btn_analyze.setStyleSheet(
            "background: #1F4E78; color: white; padding: 10px; "
            "font-weight: bold; font-size: 14px;")
        self.btn_analyze.clicked.connect(self.analyze)
        self.btn_analyze.setEnabled(False)
        op_layout.addWidget(self.btn_analyze)

        self.btn_export = QPushButton("💾 导出 Excel 报告（含图表）")
        self.btn_export.setStyleSheet(
            "background: #2E75B6; color: white; padding: 10px; "
            "font-weight: bold; font-size: 14px;")
        self.btn_export.clicked.connect(self.export_report)
        self.btn_export.setEnabled(False)
        op_layout.addWidget(self.btn_export)

        self.btn_export_csv = QPushButton("📋 导出科目表 (CSV)")
        self.btn_export_csv.setStyleSheet(
            "background: #548235; color: white; padding: 10px; "
            "font-weight: bold; font-size: 14px;")
        self.btn_export_csv.setToolTip(
            "导出长格式科目表为 CSV（适合导入数据库 / Excel 透视 / BI 工具）")
        self.btn_export_csv.clicked.connect(self.export_subjects_csv)
        self.btn_export_csv.setEnabled(False)
        op_layout.addWidget(self.btn_export_csv)

        op_layout.addStretch()
        op_layout.addWidget(QLabel("💡 数据仅在本机处理，不会上传到任何服务器"))

        layout.addWidget(op_group)

        # 进度
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # 结果标签
        self.tabs = QTabWidget()

        self.data_view = QTextEdit()
        self.data_view.setReadOnly(True)
        self.data_view.setFont(QFont("Consolas", 10))
        self.tabs.addTab(self.data_view, "📋 提取的数据")

        self.indicator_view = QTextEdit()
        self.indicator_view.setReadOnly(True)
        self.indicator_view.setFont(QFont("Consolas", 10))
        self.tabs.addTab(self.indicator_view, "📈 财务指标")

        self.trend_view = QTextEdit()
        self.trend_view.setReadOnly(True)
        self.trend_view.setFont(QFont("Consolas", 10))
        self.tabs.addTab(self.trend_view, "📊 趋势分析 (YoY/CAGR)")

        layout.addWidget(self.tabs, 1)
        self.statusBar().showMessage("就绪 - 请添加财报文件开始分析")

    def add_files(self, file_type):
        if file_type == "PDF":
            paths, _ = QFileDialog.getOpenFileNames(
                self, "选择 PDF 财报", "", "PDF Files (*.pdf)")
        else:
            paths, _ = QFileDialog.getOpenFileNames(
                self, "选择 Excel 财报", "", "Excel Files (*.xlsx *.xls)")
        for path in paths:
            if path not in self.file_list:
                self.file_list.append(path)
                year = detect_year_from_filename(path)
                year_label = f" [{year}年]" if year else " [年份未识别]"
                item = QListWidgetItem(f"📄 {Path(path).name}{year_label}")
                if year is None:
                    item.setForeground(QColor("#CC6600"))
                self.file_list_widget.addItem(item)
        if self.file_list:
            self.btn_analyze.setEnabled(True)
            self.statusBar().showMessage(
                f"已添加 {len(self.file_list)} 个文件")

    def clear_files(self):
        self.file_list = []
        self.file_list_widget.clear()
        self.btn_analyze.setEnabled(False)
        self.statusBar().showMessage("已清空文件列表")

    def analyze(self):
        if not self.file_list:
            return
        self.btn_analyze.setEnabled(False)
        self.btn_export.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)

        self.worker = Worker(self.file_list)
        self.worker.progress.connect(self.statusBar().showMessage)
        self.worker.finished.connect(self.on_analyze_done)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def on_analyze_done(self, data_by_year, indicators_by_year, years,
                        trends, errors):
        self.data_by_year = data_by_year
        self.indicators_by_year = indicators_by_year
        self.years = years
        self.trends = trends
        self.errors = errors

        # 标签 1: 提取数据
        data_text = json.dumps(data_by_year, ensure_ascii=False,
                               indent=2, default=str)
        self.data_view.setText(data_text)

        # 标签 2: 财务指标
        ind_text = self._format_indicators_text(indicators_by_year, years)
        self.indicator_view.setText(ind_text)

        # 标签 3: 趋势分析
        trend_text = self._format_trends_text(trends, years)
        self.trend_view.setText(trend_text)

        self.btn_analyze.setEnabled(True)
        self.btn_export.setEnabled(True)
        self.btn_export_csv.setEnabled(True)
        self.progress_bar.setVisible(False)

        msg = f"分析完成！\n\n覆盖 {len(years)} 年: {', '.join(map(str, years))}\n"
        msg += f"共计算 {sum(len(v) for v in indicators_by_year.values())} 项指标"
        if errors:
            msg += f"\n\n⚠️ {len(errors)} 个文件解析失败（详见报告）"
        QMessageBox.information(self, "完成", msg)
        self.statusBar().showMessage(
            f"✅ 分析完成 - {len(years)} 年 | {sum(len(v) for v in indicators_by_year.values())} 项指标")

    def _format_indicators_text(self, indicators_by_year, years):
        """格式化指标显示文本"""
        lines = [f"📊 财务指标（按年并列）", "=" * 80, ""]
        for year in years:
            indicators = indicators_by_year.get(year, {})
            lines.append(f"━━━ {year} 年 ({len(indicators)} 项) ━━━")
            for k, v in indicators.items():
                if isinstance(v, str):
                    lines.append(f"  {k:42s} | {v}")
                else:
                    lines.append(f"  {k:42s} | {v:>14,.2f}")
            lines.append("")
        return "\n".join(lines)

    def _format_trends_text(self, trends, years):
        """格式化趋势分析文本"""
        if not trends:
            return "需要至少 1 年数据才能做趋势分析"
        lines = ["📈 趋势分析（YoY 同比 + CAGR 复合增长率）", "=" * 80, ""]
        if len(years) < 2:
            lines.append("⚠️ 数据不足 2 年，仅显示当前值")
            lines.append("")

        for k, t in trends.items():
            lines.append(f"━━━ {k} ━━━")
            for year, v in t["values"]:
                if v is None:
                    lines.append(f"  {year}: -")
                else:
                    lines.append(f"  {year}: {v:>14,.2f}")
            for yoy in t["yoy"]:
                lines.append(
                    f"  {yoy['year']} YoY: {yoy['direction']} {yoy['yoy']:+.2f}%")
            if t["cagr"] is not None:
                n = t["values"][-1][0] - t["values"][0][0]
                lines.append(f"  {n}年 CAGR: {t['cagr']:+.2f}%")
            lines.append("")
        return "\n".join(lines)

    def on_error(self, msg):
        self.btn_analyze.setEnabled(True)
        self.progress_bar.setVisible(False)
        QMessageBox.critical(self, "错误", msg)
        self.statusBar().showMessage("分析失败")

    def export_report(self):
        if not self.data_by_year:
            return
        default_name = (f"财务分析报告_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                        f".xlsx")
        path, _ = QFileDialog.getSaveFileName(
            self, "保存报告", default_name, "Excel Files (*.xlsx)")
        if path:
            try:
                # 报告内不显示作者名（仅在 GUI 界面显示）
                generate_report(
                    self.data_by_year, self.indicators_by_year,
                    self.years, self.trends, path,
                    "财务分析报告", "", self.errors)
                QMessageBox.information(
                    self, "成功",
                    f"报告已保存到:\n{path}\n\n"
                    f"包含 {len(self.years)} 年数据 + 趋势分析 + 图表")
                self.statusBar().showMessage(f"报告已保存: {path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存失败: {str(e)}")

    def export_subjects_csv(self):
        """导出长格式科目表为 CSV"""
        if not self.data_by_year:
            return
        default_name = (f"科目表_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                        f".csv")
        path, _ = QFileDialog.getSaveFileName(
            self, "保存科目表", default_name, "CSV Files (*.csv)")
        if path:
            try:
                ok = export_subjects_to_csv(
                    self.data_by_year, self.years, path)
                if ok:
                    QMessageBox.information(
                        self, "成功",
                        f"科目表已导出到:\n{path}\n\n"
                        f"格式: 长格式（年份/报表类型/会计科目/类别/数值）\n"
                        f"共 {len(self.years)} 年数据\n"
                        f"可直接用 Excel 打开，或导入数据库 / Power BI")
                    self.statusBar().showMessage(f"CSV 已导出: {path}")
                else:
                    QMessageBox.warning(
                        self, "无数据",
                        "未能提取到任何科目数据，请检查原始财报文件")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败: {str(e)}")


def main():
    if hasattr(Qt, 'AA_EnableHighDpiScaling'):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
