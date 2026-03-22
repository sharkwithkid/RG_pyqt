# app_turner.py
"""
리딩게이트 반이동 자동화 — 스텝 전환 방식 UI
사이드바 제거, 상단 스텝 인디케이터, 각 스텝이 전체 너비 사용
"""
from engine import (
    inspect_work_root,
    load_all_school_names_from_db,
    scan_main_engine,
    run_main_engine,
    scan_diff_engine,
    run_diff_engine,
    get_school_domain_from_db,
    get_project_dirs,
    load_notice_templates,
)
from core.roster_log import write_work_result, write_email_sent

import sys
import json
import re
from pathlib import Path
from datetime import date as _date_type

from PyQt6.QtCore import Qt, QDate, QTimer, QThread, QObject, pyqtSignal, QPoint
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import (
    QApplication, QCheckBox, QComboBox, QDateEdit, QDialog,
    QFrame, QGridLayout, QHBoxLayout, QHeaderView, QLabel,
    QLineEdit, QListWidget, QMainWindow, QPushButton, QPlainTextEdit,
    QScrollArea, QStackedWidget, QSpinBox, QTabWidget,
    QTableWidget, QTableWidgetItem, QTextEdit, QVBoxLayout,
    QWidget, QMessageBox, QListWidgetItem, QFileDialog,
    QSizePolicy, QAbstractItemView,
)

# ── 설정/config (app.py와 동일) ──────────────────────────
CONFIG_PATH = Path("app_config.json")

DEFAULT_APP_CONFIG = {
    "work_root": "",
    "roster_log_path": "",
    "worker_name": "",
    "school_start_date": "",
    "work_date": "",
    "roster_col_map": {
        "sheet": "", "header_row": 0, "data_start": 0,
        "col_school": 0, "col_domain": 0, "col_email_arr": 0,
        "col_email_snt": 0, "col_worker": 0, "col_freshmen": 0,
        "col_transfer": 0, "col_withdraw": 0, "col_teacher": 0,
    },
}

def load_app_config() -> dict:
    if not CONFIG_PATH.exists():
        return DEFAULT_APP_CONFIG.copy()
    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        config = DEFAULT_APP_CONFIG.copy()
        if isinstance(data, dict):
            config.update(data)
        return config
    except Exception:
        return DEFAULT_APP_CONFIG.copy()

def save_app_config(data: dict) -> None:
    config = DEFAULT_APP_CONFIG.copy()
    if isinstance(data, dict):
        config.update(data)
    CONFIG_PATH.write_text(
        json.dumps(config, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

# ── QSS ──────────────────────────────────────────────────
APP_QSS = """
QMainWindow, QWidget {
    background: #FFFFFF;
    color: #111827;
    font-size: 13px;
}

QLabel#PageTitle {
    font-size: 20px;
    font-weight: 800;
    color: #0F172A;
}

QLabel#Muted {
    color: #64748B;
    font-size: 12px;
}

QLabel#CardTitle {
    font-size: 14px;
    font-weight: 700;
    color: #0F172A;
}

QFrame#Card {
    background: white;
    border: 1px solid #E5E7EB;
    border-radius: 12px;
}

QFrame#StepBar {
    background: #F8FAFC;
    border-bottom: 1px solid #E5E7EB;
}

QLineEdit, QComboBox, QDateEdit, QTextEdit, QPlainTextEdit {
    background: #FFFFFF;
    border: 1px solid #D1D5DB;
    border-radius: 8px;
    padding: 7px 10px;
    min-height: 20px;
}

QLineEdit:focus, QComboBox:focus, QDateEdit:focus {
    border: 1px solid #2563EB;
}

QListWidget {
    background: #FFFFFF;
    border: 1px solid #D1D5DB;
    border-radius: 8px;
    padding: 4px;
}

QListWidget::item { padding: 5px 8px; border-radius: 5px; }
QListWidget::item:selected { background: #DBEAFE; color: #1D4ED8; }
QListWidget::item:hover { background: #F1F5F9; }

QTableWidget {
    background: #FFFFFF;
    border: 1px solid #E5E7EB;
    border-radius: 8px;
    gridline-color: #F1F5F9;
    selection-background-color: #DBEAFE;
    selection-color: #111827;
}

QPushButton {
    background: #E5E7EB;
    border: none;
    border-radius: 8px;
    padding: 8px 16px;
    font-weight: 600;
    color: #111827;
    min-height: 20px;
}

QPushButton:hover { background: #D1D5DB; }
QPushButton:disabled { background: #F3F4F6; color: #9CA3AF; }

QPushButton#PrimaryButton {
    background: #2563EB;
    color: white;
    border-radius: 8px;
    padding: 10px 24px;
}
QPushButton#PrimaryButton:hover { background: #1D4ED8; }
QPushButton#PrimaryButton:disabled { background: #93C5FD; color: white; }

QPushButton#GhostButton {
    background: transparent;
    border: 1px solid #D1D5DB;
    color: #374151;
}
QPushButton#GhostButton:hover { background: #F9FAFB; }

QPushButton#StepBtn {
    background: transparent;
    border: none;
    border-radius: 6px;
    padding: 6px 14px;
    font-size: 13px;
    font-weight: 500;
    color: #6B7280;
    min-height: 0px;
}
QPushButton#StepBtn:hover { background: #F1F5F9; color: #374151; }

QPushButton#StepBtnActive {
    background: #EFF6FF;
    border: none;
    border-radius: 6px;
    padding: 6px 14px;
    font-size: 13px;
    font-weight: 700;
    color: #2563EB;
    min-height: 0px;
}

QHeaderView::section {
    background: #F8FAFC;
    color: #334155;
    border: none;
    border-bottom: 1px solid #E5E7EB;
    padding: 7px 8px;
    font-weight: 600;
}

QScrollArea { border: none; background: transparent; }

QScrollBar:vertical { background: transparent; width: 6px; }
QScrollBar::handle:vertical { background: #CBD5E1; border-radius: 3px; min-height: 20px; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal { background: transparent; height: 6px; }
QScrollBar::handle:horizontal { background: #CBD5E1; border-radius: 3px; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }

QTabBar::tab {
    background: #EFF2F7;
    color: #475569;
    border: none;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
    padding: 8px 16px;
    margin-right: 4px;
    font-weight: 600;
}
QTabBar::tab:selected { background: white; color: #0F172A; }
QTabWidget::pane { border: none; }
"""

# ── Workers (app.py에서 그대로) ──────────────────────────
class ScanWorker(QObject):
    finished = pyqtSignal(object)
    failed   = pyqtSignal(str)

    def __init__(self, work_root, school_name, school_start_date, work_date,
                 roster_xlsx=None, col_map=None):
        super().__init__()
        self.work_root         = work_root
        self.school_name       = school_name
        self.school_start_date = school_start_date
        self.work_date         = work_date
        self.roster_xlsx       = roster_xlsx
        self.col_map           = col_map

    def run(self):
        try:
            result = scan_main_engine(
                work_root=self.work_root,
                school_name=self.school_name,
                school_start_date=self.school_start_date,
                work_date=self.work_date,
                roster_xlsx=self.roster_xlsx,
                col_map=self.col_map,
            )
            self.finished.emit(result)
        except Exception as e:
            self.failed.emit(str(e))


class RunWorker(QObject):
    finished = pyqtSignal(object)
    failed   = pyqtSignal(str)

    def __init__(self, scan, work_date, school_start_date,
                 layout_overrides=None, school_kind_override=None):
        super().__init__()
        self.scan                 = scan
        self.work_date            = work_date
        self.school_start_date    = school_start_date
        self.layout_overrides     = layout_overrides or {}
        self.school_kind_override = school_kind_override

    def run(self):
        try:
            result = run_main_engine(
                scan=self.scan,
                work_date=self.work_date,
                school_start_date=self.school_start_date,
                layout_overrides=self.layout_overrides,
                school_kind_override=self.school_kind_override,
            )
            self.finished.emit(result)
        except Exception as e:
            self.failed.emit(str(e))


# ── 상단 스텝 인디케이터 ─────────────────────────────────
STEP_LABELS = ["기본 설정", "학교 선택", "스캔", "실행·결과", "안내문"]

class StepIndicator(QWidget):
    step_clicked = pyqtSignal(int)

    def __init__(self):
        super().__init__()
        self._current = 0
        self._build_ui()

    def _build_ui(self):
        bar = QFrame()
        bar.setObjectName("StepBar")
        bar.setFixedHeight(48)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        outer.addWidget(bar)

        row = QHBoxLayout(bar)
        row.setContentsMargins(24, 0, 24, 0)
        row.setSpacing(4)

        # 앱 이름
        app_name = QLabel("리딩게이트 반이동 자동화")
        app_name.setStyleSheet("font-weight: 700; font-size: 13px; color: #0F172A;")
        row.addWidget(app_name)
        row.addSpacing(24)

        self._btns = []
        for i, label in enumerate(STEP_LABELS):
            btn = QPushButton(f"{i+1}. {label}")
            btn.setObjectName("StepBtn")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(lambda _, idx=i: self.step_clicked.emit(idx))
            self._btns.append(btn)
            row.addWidget(btn)

        row.addStretch()

        self.set_step(0)

    def set_step(self, idx: int):
        self._current = idx
        for i, btn in enumerate(self._btns):
            if i == idx:
                btn.setObjectName("StepBtnActive")
            else:
                btn.setObjectName("StepBtn")
            # QSS 갱신 강제
            btn.style().unpolish(btn)
            btn.style().polish(btn)

    def set_step_enabled(self, idx: int, enabled: bool):
        self._btns[idx].setEnabled(enabled)


# ── 스텝 1: 기본 설정 ────────────────────────────────────
class Step1Setup(QWidget):
    def __init__(self):
        super().__init__()
        self._build_ui()

    def _build_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        page = QWidget()
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(0, 0, 0, 0)

        container = QWidget()
        container.setMaximumWidth(640)
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 32, 0, 40)
        layout.setSpacing(20)

        title = QLabel("기본 설정")
        title.setObjectName("PageTitle")
        layout.addWidget(title)

        sub = QLabel("작업 폴더와 기본 정보를 입력하세요.")
        sub.setObjectName("Muted")
        layout.addWidget(sub)

        layout.addSpacing(8)

        # 작업 폴더
        self.work_root_input = QLineEdit()
        self.work_root_input.setPlaceholderText("작업 폴더를 선택하거나 직접 입력")
        self.btn_browse_path = QPushButton("찾아보기")
        self.btn_browse_path.setObjectName("GhostButton")
        self.btn_browse_path.setFixedWidth(90)
        layout.addWidget(self._field("작업 폴더", self.work_root_input, self.btn_browse_path))

        # 명단 파일
        self.roster_log_input = QLineEdit()
        self.roster_log_input.setPlaceholderText("학교 전체 명단 .xlsx 파일 선택")
        self.btn_browse_roster = QPushButton("찾아보기")
        self.btn_browse_roster.setObjectName("GhostButton")
        self.btn_browse_roster.setFixedWidth(90)
        layout.addWidget(self._field("학교 전체 명단 파일", self.roster_log_input, self.btn_browse_roster))

        # 작업자
        self.worker_input = QLineEdit()
        self.worker_input.setPlaceholderText("작업자 이름")
        layout.addWidget(self._field("작업자", self.worker_input))

        # 날짜
        self.open_date_edit = QDateEdit()
        self.open_date_edit.setCalendarPopup(True)
        self.open_date_edit.setDate(QDate.currentDate())
        self.open_date_edit.wheelEvent = lambda e: None

        self.work_date_edit = QDateEdit()
        self.work_date_edit.setCalendarPopup(True)
        self.work_date_edit.setDate(QDate.currentDate())
        self.work_date_edit.wheelEvent = lambda e: None

        date_row = QHBoxLayout()
        date_row.setSpacing(16)
        date_row.addWidget(self._field("개학일", self.open_date_edit))
        date_row.addWidget(self._field("작업일", self.work_date_edit))
        layout.addLayout(date_row)

        layout.addSpacing(8)

        # 버튼
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        self.btn_save = QPushButton("설정 저장")
        self.btn_save.setObjectName("GhostButton")
        self.btn_load = QPushButton("설정 불러오기")
        self.btn_load.setObjectName("GhostButton")
        self.btn_start = QPushButton("다음 →")
        self.btn_start.setObjectName("PrimaryButton")
        self.btn_start.setEnabled(False)
        btn_row.addWidget(self.btn_save)
        btn_row.addWidget(self.btn_load)
        btn_row.addStretch()
        btn_row.addWidget(self.btn_start)
        layout.addLayout(btn_row)

        page_layout.addWidget(container, 0, Qt.AlignmentFlag.AlignHCenter)
        page_layout.addStretch()
        scroll.setWidget(page)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)

        self.work_root_input.textChanged.connect(self._refresh_btn)
        self.worker_input.textChanged.connect(self._refresh_btn)

    def _field(self, label_text, widget, btn=None):
        wrap = QWidget()
        layout = QVBoxLayout(wrap)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)
        lbl = QLabel(label_text)
        lbl.setObjectName("Muted")
        layout.addWidget(lbl)
        if btn:
            row = QHBoxLayout()
            row.setContentsMargins(0, 0, 0, 0)
            row.setSpacing(8)
            row.addWidget(widget, 1)
            row.addWidget(btn)
            layout.addLayout(row)
        else:
            layout.addWidget(widget)
        return wrap

    def _refresh_btn(self):
        ok = bool(self.work_root_input.text().strip()) and bool(self.worker_input.text().strip())
        self.btn_start.setEnabled(ok)


# ── 스텝 2: 학교 선택 ────────────────────────────────────
class Step2School(QWidget):
    def __init__(self):
        super().__init__()
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 32, 40, 40)
        layout.setSpacing(16)

        title = QLabel("학교 선택")
        title.setObjectName("PageTitle")
        layout.addWidget(title)

        sub = QLabel("작업할 학교를 검색해서 선택하세요.")
        sub.setObjectName("Muted")
        layout.addWidget(sub)

        layout.addSpacing(8)

        # 검색
        search_row = QHBoxLayout()
        search_row.setSpacing(10)
        self.school_input = QLineEdit()
        self.school_input.setPlaceholderText("학교명을 입력하세요")
        self.school_input.setFixedHeight(38)
        self.btn_select = QPushButton("선택")
        self.btn_select.setObjectName("PrimaryButton")
        self.btn_select.setEnabled(False)
        self.btn_select.setFixedHeight(38)
        search_row.addWidget(self.school_input, 1)
        search_row.addWidget(self.btn_select)
        layout.addLayout(search_row)

        self.status_label = QLabel("")
        self.status_label.setObjectName("Muted")
        layout.addWidget(self.status_label)

        # 결과 목록
        self.result_list = QListWidget()
        self.result_list.setMaximumHeight(320)
        layout.addWidget(self.result_list)

        layout.addStretch()

        # 하단 버튼
        btn_row = QHBoxLayout()
        self.btn_back = QPushButton("← 이전")
        self.btn_back.setObjectName("GhostButton")
        self.btn_next = QPushButton("다음 →")
        self.btn_next.setObjectName("PrimaryButton")
        self.btn_next.setEnabled(False)
        btn_row.addWidget(self.btn_back)
        btn_row.addStretch()
        btn_row.addWidget(self.btn_next)
        layout.addLayout(btn_row)


# ── 스텝 3: 스캔 ─────────────────────────────────────────
class Step3Scan(QWidget):
    def __init__(self):
        super().__init__()
        self._build_ui()

    def _build_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(40, 32, 40, 40)
        layout.setSpacing(16)

        # 헤더
        head_row = QHBoxLayout()
        head_left = QVBoxLayout()
        head_left.setSpacing(4)
        title = QLabel("스캔 검수")
        title.setObjectName("PageTitle")
        sub = QLabel("입력 파일 구조를 확인하고 의심 행을 검수합니다.")
        sub.setObjectName("Muted")
        head_left.addWidget(title)
        head_left.addWidget(sub)

        btn_col = QVBoxLayout()
        btn_col.setSpacing(8)
        self.btn_scan = QPushButton("파일 내용 스캔")
        self.btn_scan.setObjectName("PrimaryButton")
        self.btn_scan.setFixedHeight(38)

        action_row = QHBoxLayout()
        action_row.setSpacing(8)
        self.btn_show_log = QPushButton("로그 보기")
        self.btn_show_log.setObjectName("GhostButton")
        self.btn_open_file = QPushButton("파일 열기")
        self.btn_open_file.setObjectName("GhostButton")
        self.btn_open_file.setEnabled(False)
        action_row.addWidget(self.btn_show_log)
        action_row.addWidget(self.btn_open_file)

        btn_col.addWidget(self.btn_scan)
        btn_col.addLayout(action_row)

        head_row.addLayout(head_left, 1)
        head_row.addLayout(btn_col)
        layout.addLayout(head_row)

        # 상태
        self.scan_status = QLabel("스캔을 실행해 주세요.")
        self.scan_status.setObjectName("Muted")
        layout.addWidget(self.scan_status)

        # 파일 테이블 — 구분/파일명/시트/자동감지/확인(체크박스)
        self.scan_table = QTableWidget(4, 5)
        self.scan_table.setHorizontalHeaderLabels(["구분", "파일명", "시트", "자동 감지", "확인"])
        self.scan_table.verticalHeader().setVisible(False)
        self.scan_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.scan_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.scan_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.scan_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.scan_table.setColumnWidth(0, 70)
        self.scan_table.setColumnWidth(2, 80)
        self.scan_table.setColumnWidth(3, 80)
        self.scan_table.setColumnWidth(4, 50)
        self.scan_table.setFixedHeight(165)

        self._scan_checkboxes = []
        for r, kind in enumerate(["신입생", "전입생", "전출생", "교직원"]):
            item = QTableWidgetItem(kind)
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.scan_table.setItem(r, 0, item)

            # 확인 체크박스
            chk_widget = QWidget()
            chk_layout = QHBoxLayout(chk_widget)
            chk_layout.setContentsMargins(0, 0, 0, 0)
            chk_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            chk = QCheckBox()
            chk_layout.addWidget(chk)
            self.scan_table.setCellWidget(r, 4, chk_widget)
            self._scan_checkboxes.append(chk)

        layout.addWidget(self.scan_table)

        # 하단 버튼
        btn_row = QHBoxLayout()
        self.btn_back = QPushButton("← 이전")
        self.btn_back.setObjectName("GhostButton")
        self.btn_next = QPushButton("실행으로 →")
        self.btn_next.setObjectName("PrimaryButton")
        self.btn_next.setEnabled(False)
        btn_row.addWidget(self.btn_back)
        btn_row.addStretch()
        btn_row.addWidget(self.btn_next)
        layout.addLayout(btn_row)

        layout.addStretch()
        scroll.setWidget(page)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)


# ── 스텝 4: 실행·결과 ────────────────────────────────────
class Step4Run(QWidget):
    def __init__(self):
        super().__init__()
        self._build_ui()

    def _build_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(40, 32, 40, 40)
        layout.setSpacing(16)

        head_row = QHBoxLayout()
        left = QVBoxLayout()
        left.setSpacing(4)
        title = QLabel("실행 · 결과")
        title.setObjectName("PageTitle")
        sub = QLabel("스캔 결과를 바탕으로 반이동 파일을 생성합니다.")
        sub.setObjectName("Muted")
        left.addWidget(title)
        left.addWidget(sub)

        self.btn_run = QPushButton("실행")
        self.btn_run.setObjectName("PrimaryButton")
        self.btn_run.setFixedSize(100, 38)
        self.btn_run.setEnabled(False)

        head_row.addLayout(left, 1)
        head_row.addWidget(self.btn_run, 0, Qt.AlignmentFlag.AlignBottom)
        layout.addLayout(head_row)

        self.run_status = QLabel("스캔을 먼저 완료해 주세요.")
        self.run_status.setObjectName("Muted")
        layout.addWidget(self.run_status)

        # 결과 파일 콤보
        file_row = QHBoxLayout()
        file_row.setSpacing(10)
        self.run_file_combo = QComboBox()
        self.btn_open_result = QPushButton("파일 열기")
        self.btn_open_result.setObjectName("GhostButton")
        self.btn_show_run_log = QPushButton("로그 보기")
        self.btn_show_run_log.setObjectName("GhostButton")
        file_row.addWidget(self.run_file_combo, 1)
        file_row.addWidget(self.btn_open_result)
        file_row.addWidget(self.btn_show_run_log)
        layout.addLayout(file_row)

        # 시트 탭
        self.run_sheet_tabs = QTabWidget()
        layout.addWidget(self.run_sheet_tabs, 1)

        # 하단 버튼
        btn_row = QHBoxLayout()
        self.btn_back = QPushButton("← 이전")
        self.btn_back.setObjectName("GhostButton")
        self.btn_next = QPushButton("안내문으로 →")
        self.btn_next.setObjectName("PrimaryButton")
        self.btn_next.setEnabled(False)
        btn_row.addWidget(self.btn_back)
        btn_row.addStretch()
        btn_row.addWidget(self.btn_next)
        layout.addLayout(btn_row)

        scroll.setWidget(page)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)


# ── 스텝 5: 안내문 ───────────────────────────────────────
class Step5Notice(QWidget):
    def __init__(self):
        super().__init__()
        self._build_ui()

    def _build_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(40, 32, 40, 40)
        layout.setSpacing(16)

        title = QLabel("안내문")
        title.setObjectName("PageTitle")
        sub = QLabel("안내문을 확인하고 이메일 발송 여부를 기록합니다.")
        sub.setObjectName("Muted")
        layout.addWidget(title)
        layout.addWidget(sub)
        layout.addSpacing(8)

        # 안내문 선택 + 본문
        body_row = QHBoxLayout()
        body_row.setSpacing(16)

        self.notice_list = QListWidget()
        self.notice_list.setFixedWidth(180)
        body_row.addWidget(self.notice_list)

        notice_right = QVBoxLayout()
        notice_right.setSpacing(8)
        self.notice_text = QTextEdit()
        self.notice_text.setReadOnly(True)
        self.notice_text.setMinimumHeight(240)
        self.btn_copy_notice = QPushButton("복사")
        self.btn_copy_notice.setObjectName("GhostButton")
        notice_right.addWidget(self.notice_text, 1)
        notice_right.addWidget(self.btn_copy_notice, 0, Qt.AlignmentFlag.AlignRight)
        body_row.addLayout(notice_right, 1)

        layout.addLayout(body_row, 1)

        # 이메일 발송 기록
        email_row = QHBoxLayout()
        email_row.setSpacing(10)
        email_lbl = QLabel("완료 이메일:")
        email_lbl.setObjectName("Muted")
        self.btn_email_sent = QPushButton("발송 완료")
        self.btn_email_sent.setCheckable(True)
        self.btn_email_hold = QPushButton("보류")
        self.btn_email_hold.setCheckable(True)
        email_row.addWidget(email_lbl)
        email_row.addWidget(self.btn_email_sent)
        email_row.addWidget(self.btn_email_hold)
        email_row.addStretch()

        self.btn_record_roster = QPushButton("명단 기록")
        self.btn_record_roster.setObjectName("PrimaryButton")
        self.btn_record_roster.setEnabled(False)
        email_row.addWidget(self.btn_record_roster)
        layout.addLayout(email_row)

        self.email_status = QLabel("")
        self.email_status.setObjectName("Muted")
        layout.addWidget(self.email_status)

        # 하단 버튼
        btn_row = QHBoxLayout()
        self.btn_back = QPushButton("← 이전")
        self.btn_back.setObjectName("GhostButton")
        self.btn_new = QPushButton("새 작업 시작")
        self.btn_new.setObjectName("GhostButton")
        btn_row.addWidget(self.btn_back)
        btn_row.addStretch()
        btn_row.addWidget(self.btn_new)
        layout.addLayout(btn_row)

        scroll.setWidget(page)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(scroll)


# ── 메인 윈도우 ──────────────────────────────────────────
class TurnerWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("리딩게이트 반이동 자동화")
        self.resize(1100, 780)

        # 상태
        self.worker_name        = ""
        self.work_root          = None
        self.work_date          = None
        self.school_start_date  = None
        self.selected_school    = None
        self.roster_log_path    = None
        self._roster_col_map    = {}
        self._pending_roster    = False
        self._school_names      = []
        self._school_name_set   = set()
        self._last_scan_result  = None
        self._last_run_result   = None
        self._last_scan_logs    = []
        self._last_run_logs     = []
        self._current_outputs   = []
        self._selected_domain   = ""
        self._school_kind_override = None
        self._scan_thread       = None
        self._run_thread        = None
        self._scan_worker       = None
        self._scan_file_paths   = {}
        self._run_worker        = None

        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(250)
        self._search_timer.timeout.connect(self._do_school_search)

        # 스텝 위젯
        self.step1 = Step1Setup()
        self.step2 = Step2School()
        self.step3 = Step3Scan()
        self.step4 = Step4Run()
        self.step5 = Step5Notice()

        # 스텝 인디케이터
        self.step_bar = StepIndicator()

        # 스택
        self.stack = QStackedWidget()
        self.stack.addWidget(self.step1)  # 0
        self.stack.addWidget(self.step2)  # 1
        self.stack.addWidget(self.step3)  # 2
        self.stack.addWidget(self.step4)  # 3
        self.stack.addWidget(self.step5)  # 4

        central = QWidget()
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)
        root.addWidget(self.step_bar)
        root.addWidget(self.stack, 1)
        self.setCentralWidget(central)

        # 스텝 버튼 초기 비활성
        for i in range(1, 5):
            self.step_bar.set_step_enabled(i, False)

        self._wire()
        self._apply_saved_config()

    # ── 시그널 연결 ────────────────────────────────────────
    def _wire(self):
        # 스텝바 클릭
        self.step_bar.step_clicked.connect(self._go_to_step)

        # 스텝1
        self.step1.btn_browse_path.clicked.connect(self._browse_work_root)
        self.step1.btn_browse_roster.clicked.connect(self._browse_roster)
        self.step1.btn_save.clicked.connect(self._save_config)
        self.step1.btn_load.clicked.connect(self._load_config)
        self.step1.btn_start.clicked.connect(self._on_step1_next)

        # 스텝2
        self.step2.school_input.textChanged.connect(self._on_school_input_changed)
        self.step2.result_list.itemClicked.connect(self._on_school_item_clicked)
        self.step2.result_list.itemDoubleClicked.connect(lambda _: self._on_school_confirmed())
        self.step2.btn_select.clicked.connect(self._on_school_confirmed)
        self.step2.btn_back.clicked.connect(lambda: self._go_to_step(0))
        self.step2.btn_next.clicked.connect(lambda: self._go_to_step(2))

        # 스텝3
        self.step3.btn_scan.clicked.connect(self._run_scan)
        self.step3.btn_show_log.clicked.connect(
            lambda: self._show_log("스캔 로그", self._last_scan_logs))
        self.step3.btn_open_file.clicked.connect(self._open_selected_scan_file)
        self.step3.scan_table.cellDoubleClicked.connect(self._on_scan_table_double_click)
        self.step3.btn_back.clicked.connect(lambda: self._go_to_step(1))
        self.step3.btn_next.clicked.connect(lambda: self._go_to_step(3))

        # 스텝4
        self.step4.btn_run.clicked.connect(self._run_main)
        self.step4.btn_show_run_log.clicked.connect(
            lambda: self._show_log("실행 로그", self._last_run_logs))
        self.step4.btn_back.clicked.connect(lambda: self._go_to_step(2))
        self.step4.btn_next.clicked.connect(lambda: self._go_to_step(4))

        # 스텝5
        self.step5.btn_back.clicked.connect(lambda: self._go_to_step(3))
        self.step5.btn_new.clicked.connect(self._new_work)
        self.step5.notice_list.currentRowChanged.connect(self._on_notice_changed)
        self.step5.btn_copy_notice.clicked.connect(self._copy_notice)
        self.step5.btn_record_roster.clicked.connect(self._record_roster)

    # ── 스텝 이동 ──────────────────────────────────────────
    def _go_to_step(self, idx: int):
        self.stack.setCurrentIndex(idx)
        self.step_bar.set_step(idx)

    # ── 설정 저장/불러오기 ─────────────────────────────────
    def _apply_saved_config(self):
        cfg = load_app_config()
        self.step1.work_root_input.setText(cfg.get("work_root", ""))
        self.step1.roster_log_input.setText(cfg.get("roster_log_path", ""))
        self.step1.worker_input.setText(cfg.get("worker_name", ""))
        if cfg.get("school_start_date"):
            try:
                d = _date_type.fromisoformat(cfg["school_start_date"])
                self.step1.open_date_edit.setDate(QDate(d.year, d.month, d.day))
            except Exception:
                pass
        if cfg.get("work_date"):
            try:
                d = _date_type.fromisoformat(cfg["work_date"])
                self.step1.work_date_edit.setDate(QDate(d.year, d.month, d.day))
            except Exception:
                pass
        self._roster_col_map = cfg.get("roster_col_map", {})

    def _save_config(self):
        cfg = load_app_config()
        cfg["work_root"]        = self.step1.work_root_input.text().strip()
        cfg["roster_log_path"]  = self.step1.roster_log_input.text().strip()
        cfg["worker_name"]      = self.step1.worker_input.text().strip()
        qd = self.step1.open_date_edit.date()
        cfg["school_start_date"] = f"{qd.year()}-{qd.month():02d}-{qd.day():02d}"
        qd = self.step1.work_date_edit.date()
        cfg["work_date"]        = f"{qd.year()}-{qd.month():02d}-{qd.day():02d}"
        save_app_config(cfg)
        QMessageBox.information(self, "저장 완료", "기본 설정이 저장됐습니다.")

    def _load_config(self):
        self._apply_saved_config()
        QMessageBox.information(self, "불러오기 완료", "저장된 설정을 불러왔습니다.")

    # ── 파일 찾아보기 ──────────────────────────────────────
    def _browse_work_root(self):
        path = QFileDialog.getExistingDirectory(self, "작업 폴더 선택")
        if path:
            self.step1.work_root_input.setText(path)

    def _browse_roster(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "명단 파일 선택", "", "Excel 파일 (*.xlsx)")
        if path:
            self.step1.roster_log_input.setText(path)
            cfg = load_app_config()
            self._roster_col_map = cfg.get("roster_col_map", {})

    # ── 스텝1 → 스텝2 ──────────────────────────────────────
    def _on_step1_next(self):
        from datetime import date as _date_cls, datetime as _dt

        work_root_text = self.step1.work_root_input.text().strip()
        worker_text    = self.step1.worker_input.text().strip()
        roster_text    = self.step1.roster_log_input.text().strip()

        if not work_root_text:
            QMessageBox.warning(self, "입력 확인", "작업 폴더를 입력하세요.")
            return
        work_root_path = Path(work_root_text)
        if not work_root_path.exists():
            QMessageBox.warning(self, "입력 확인", f"작업 폴더를 찾을 수 없습니다.\n{work_root_text}")
            return
        if not worker_text:
            QMessageBox.warning(self, "입력 확인", "작업자 이름을 입력하세요.")
            return

        inspect = inspect_work_root(work_root_path)
        if not inspect.get("ok", False):
            errors = inspect.get("errors", [])
            QMessageBox.warning(self, "리소스 확인",
                "작업 폴더의 resources 구성이 올바르지 않습니다.\n\n" + "\n".join(errors[:10]))
            return

        if not roster_text:
            QMessageBox.warning(self, "명단 파일 필요",
                "학교 전체 명단 파일(.xlsx)을 지정해야 합니다.\n'찾아보기' 버튼으로 파일을 선택해 주세요.")
            return

        roster_path = Path(roster_text)
        if not roster_path.exists():
            QMessageBox.warning(self, "명단 파일 없음",
                f"지정된 명단 파일을 찾을 수 없습니다.\n{roster_text}")
            return

        cfg = load_app_config()
        col_map = cfg.get("roster_col_map", {})
        if not col_map.get("col_school"):
            QMessageBox.warning(self, "열 매핑 필요",
                "명단 파일의 열 매핑이 완료되지 않았습니다.\n'찾아보기' 버튼으로 열 매핑을 완료해 주세요.")
            return

        self.worker_name       = worker_text
        self.work_root         = work_root_path
        self.roster_log_path   = roster_path
        self._roster_col_map   = col_map

        qd = self.step1.open_date_edit.date()
        self.school_start_date = _date_cls(qd.year(), qd.month(), qd.day())
        qd = self.step1.work_date_edit.date()
        self.work_date         = _date_cls(qd.year(), qd.month(), qd.day())

        # 학교 목록 로드
        self.step2.status_label.setText("학교 목록 로드 중...")
        try:
            names = load_all_school_names_from_db(self.roster_log_path, self._roster_col_map)
            self._school_names    = names
            self._school_name_set = set(names)
            self.step2.status_label.setText(f"학교 {len(names)}개 준비 완료")
        except Exception as e:
            self.step2.status_label.setText(f"학교 목록 로드 실패: {e}")
            self._school_names    = []
            self._school_name_set = set()

        # 스텝2로
        self.step_bar.set_step_enabled(1, True)
        self._go_to_step(1)

    # ── 학교 검색 ──────────────────────────────────────────
    def _on_school_input_changed(self, text):
        self._search_timer.start()

    def _do_school_search(self):
        kw = self.step2.school_input.text().strip().lower()
        self.step2.result_list.clear()
        self.step2.btn_select.setEnabled(False)
        self.step2.btn_next.setEnabled(False)

        if not kw:
            matched = self._school_names[:100]
        else:
            matched = [n for n in self._school_names if kw in n.lower()][:100]

        for name in matched:
            self.step2.result_list.addItem(QListWidgetItem(name))

    def _on_school_item_clicked(self, item):
        self.step2.school_input.blockSignals(True)
        self.step2.school_input.setText(item.text())
        self.step2.school_input.blockSignals(False)
        self.step2.btn_select.setEnabled(item.text() in self._school_name_set)

    def _on_school_confirmed(self):
        name = self.step2.school_input.text().strip()
        if name not in self._school_name_set:
            QMessageBox.warning(self, "학교 선택", "명단에 없는 학교입니다.")
            return

        # 폴더 확인
        school_dirs = [
            p for p in Path(self.work_root).iterdir()
            if p.is_dir() and "resources" not in p.name.lower()
            and not p.name.startswith(".")
        ]
        matched = [p for p in school_dirs if name in p.name]
        if not matched:
            QMessageBox.warning(self, "학교 폴더 없음",
                f"'{name}' 학교 폴더를 작업 폴더에서 찾을 수 없습니다.")
            return

        self.selected_school = name
        dirs = get_project_dirs(self.work_root)
        self._selected_domain = get_school_domain_from_db(
            self.roster_log_path, name, self._roster_col_map) or ""

        self.step2.status_label.setText(f"선택됨: {name}")
        self.step2.btn_next.setEnabled(True)
        self.step_bar.set_step_enabled(2, True)
        self._go_to_step(2)

    # ── 스캔 ───────────────────────────────────────────────
    def _run_scan(self):
        if not self.selected_school:
            QMessageBox.warning(self, "학교 미선택", "학교를 먼저 선택해 주세요.")
            return

        self.step3.btn_scan.setEnabled(False)
        self.step3.scan_status.setText("스캔 중...")

        self._scan_worker = ScanWorker(
            work_root=self.work_root,
            school_name=self.selected_school,
            school_start_date=self.school_start_date,
            work_date=self.work_date,
            roster_xlsx=self.roster_log_path,
            col_map=self._roster_col_map,
        )
        self._scan_thread = QThread()
        self._scan_worker.moveToThread(self._scan_thread)
        self._scan_thread.started.connect(self._scan_worker.run)
        self._scan_worker.finished.connect(self._on_scan_done)
        self._scan_worker.failed.connect(self._on_scan_failed)
        self._scan_worker.finished.connect(self._scan_thread.quit)
        self._scan_worker.failed.connect(self._scan_thread.quit)
        self._scan_worker.finished.connect(self._scan_worker.deleteLater)
        self._scan_worker.failed.connect(self._scan_worker.deleteLater)
        self._scan_thread.finished.connect(self._scan_thread.deleteLater)
        self._scan_thread.start()

    def _on_scan_done(self, result):
        self.step3.btn_scan.setEnabled(True)
        self._last_scan_result = result
        self._last_scan_logs   = result.logs or []

        if not result.ok:
            err = next((l for l in reversed(result.logs or []) if "[ERROR]" in l), "스캔 실패")
            self.step3.scan_status.setText(err)
            return

        # 파일별 경고 수집
        file_warns = {}
        kind_label_map = {
            "신입생": result.freshmen,
            "전입생": result.transfer_in,
            "전출생": result.transfer_out,
            "교직원": result.teachers,
        }
        for kind_label, meta in kind_label_map.items():
            if meta and meta.get("warning"):
                warns = [w for w in meta["warning"].split("\n") if w.strip()]
                if warns:
                    file_warns[kind_label] = warns

        warn_count = sum(len(v) for v in file_warns.values())
        if warn_count:
            self.step3.scan_status.setText(f"⚠  스캔 완료 — 경고 {warn_count}건")
            self.step3.scan_status.setStyleSheet("color: #D97706; font-weight: 600;")
            self._show_scan_warn_popup(file_warns)
        else:
            self.step3.scan_status.setText("✓  스캔 완료 — 이상 없음")
            self.step3.scan_status.setStyleSheet("color: #16A34A; font-weight: 600;")

        # 테이블 채우기
        kind_map = {"신입생": 0, "전입생": 1, "전출생": 2, "교직원": 3}
        file_meta = {
            "신입생": result.freshmen,
            "전입생": result.transfer_in,
            "전출생": result.transfer_out,
            "교직원": result.teachers,
        }
        for kind, row in kind_map.items():
            meta = file_meta.get(kind)
            fname  = meta.get("file_name", "")      if meta else ""
            sheet  = meta.get("sheet_name", "")     if meta else ""
            h_row  = meta.get("header_row", "")     if meta else ""
            d_row  = meta.get("data_start_row", "") if meta else ""
            auto_text = f"{h_row} / {d_row}" if h_row and d_row else ""
            self.step3.scan_table.setItem(row, 1, QTableWidgetItem(fname))
            self.step3.scan_table.setItem(row, 2, QTableWidgetItem(str(sheet)))
            self.step3.scan_table.setItem(row, 3, QTableWidgetItem(auto_text))
            # 체크박스 초기화
            if row < len(self.step3._scan_checkboxes):
                self.step3._scan_checkboxes[row].setChecked(False)

        # 파일 경로 맵 저장 (더블클릭/파일열기용)
        r = self._last_scan_result
        self._scan_file_paths = {
            0: r.freshmen_file,
            1: r.transfer_file,
            2: r.withdraw_file,
            3: r.teacher_file,
        }
        self.step3.btn_open_file.setEnabled(True)
        self.step3.btn_next.setEnabled(True)
        self.step4.btn_run.setEnabled(True)
        self.step_bar.set_step_enabled(3, True)
        self._load_notices()

    def _on_scan_failed(self, msg):
        self.step3.btn_scan.setEnabled(True)
        self.step3.scan_status.setText(f"오류: {msg}")

    # ── 실행 ───────────────────────────────────────────────
    def _run_main(self):
        if not self._last_scan_result or not self._last_scan_result.ok:
            QMessageBox.warning(self, "스캔 필요", "먼저 스캔을 완료해 주세요.")
            return

        self.step4.btn_run.setEnabled(False)
        self.step4.run_status.setText("실행 중...")

        self._run_worker = RunWorker(
            scan=self._last_scan_result,
            work_date=self.work_date,
            school_start_date=self.school_start_date,
            school_kind_override=self._school_kind_override,
        )
        self._run_thread = QThread()
        self._run_worker.moveToThread(self._run_thread)
        self._run_thread.started.connect(self._run_worker.run)
        self._run_worker.finished.connect(self._on_run_done)
        self._run_worker.failed.connect(self._on_run_failed)
        self._run_worker.finished.connect(self._run_thread.quit)
        self._run_worker.failed.connect(self._run_thread.quit)
        self._run_worker.finished.connect(self._run_worker.deleteLater)
        self._run_worker.failed.connect(self._run_worker.deleteLater)
        self._run_thread.finished.connect(self._run_thread.deleteLater)
        self._run_thread.start()

    def _on_run_done(self, result):
        self.step4.btn_run.setEnabled(True)
        self._last_run_result = result
        self._last_run_logs   = result.logs or []
        self._current_outputs = list(result.outputs or [])

        if not result.ok:
            err = next((l for l in reversed(result.logs or []) if "[ERROR]" in l), "실행 실패")
            self.step4.run_status.setText(err)
            return

        self.step4.run_status.setText(f"실행 완료 — {len(self._current_outputs)}개 파일 생성")

        # 파일 콤보 채우기
        self.step4.run_file_combo.blockSignals(True)
        self.step4.run_file_combo.clear()
        for p in self._current_outputs:
            self.step4.run_file_combo.addItem(Path(p).name)
        self.step4.run_file_combo.blockSignals(False)

        self.step4.btn_next.setEnabled(True)
        self.step5.btn_record_roster.setEnabled(True)
        self.step_bar.set_step_enabled(4, True)
        self._load_notices()

    def _on_run_failed(self, msg):
        self.step4.btn_run.setEnabled(True)
        self.step4.run_status.setText(f"오류: {msg}")

    # ── 안내문 ─────────────────────────────────────────────
    def _load_notices(self):
        if not self.work_root:
            return
        try:
            templates = load_notice_templates(self.work_root)
            self.step5.notice_list.clear()
            for key in sorted(templates.keys()):
                self.step5.notice_list.addItem(key)
            self._notice_templates = templates
        except Exception:
            self._notice_templates = {}

    def _on_notice_changed(self, row):
        if row < 0:
            return
        key = self.step5.notice_list.item(row).text()
        text = getattr(self, "_notice_templates", {}).get(key, "")

        # 도메인/학교명 치환
        if self.selected_school:
            text = text.replace("{학교명}", self.selected_school)
        if self._selected_domain:
            text = text.replace("{도메인}", self._selected_domain)
        self.step5.notice_text.setPlainText(text)

    def _copy_notice(self):
        text = self.step5.notice_text.toPlainText()
        if text:
            QApplication.clipboard().setText(text)

    # ── 명단 기록 ──────────────────────────────────────────
    def _record_roster(self):
        if not self.roster_log_path or not self.selected_school or not self._last_run_result:
            return

        kind_flags = {}
        if self._last_run_result:
            r = self._last_run_result
            kind_flags = {
                "신입생": bool(getattr(r, "freshmen_count", 0)),
                "전입생": bool(getattr(r, "transfer_in_count", 0)),
                "전출생": bool(getattr(r, "transfer_out_count", 0)),
                "교직원": bool(getattr(r, "teacher_count", 0)),
            }

        ok, msg = write_work_result(
            xlsx_path=self.roster_log_path,
            school_name=self.selected_school,
            worker=self.worker_name,
            kind_flags=kind_flags,
            col_map=self._roster_col_map,
        )
        self.step5.email_status.setText(msg)

    # ── 유틸 ───────────────────────────────────────────────
    def _show_scan_warn_popup(self, file_warns: dict):
        """파일별 경고를 팝업으로 표시"""
        dlg = QDialog(self)
        dlg.setWindowTitle("스캔 경고")
        dlg.resize(600, 400)
        layout = QVBoxLayout(dlg)

        title = QLabel("⚠  다음 항목을 확인해 주세요.")
        title.setStyleSheet("font-weight: 700; font-size: 14px; color: #D97706;")
        layout.addWidget(title)
        layout.addSpacing(8)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        inner = QWidget()
        inner_layout = QVBoxLayout(inner)
        inner_layout.setSpacing(12)

        for kind_label, warns in file_warns.items():
            # 파일 섹션 제목
            sec_title = QLabel(f"▸ {kind_label}")
            sec_title.setStyleSheet("font-weight: 700; color: #374151;")
            inner_layout.addWidget(sec_title)

            # 경고 목록 — 중복 제거 후 표시
            seen = set()
            for w in warns:
                key = w.split("'")[-2] if "'" in w else w  # 값 기준 dedup
                if key not in seen:
                    seen.add(key)
                    lbl = QLabel(f"  {w}")
                    lbl.setStyleSheet("color: #6B7280; font-size: 12px;")
                    lbl.setWordWrap(True)
                    inner_layout.addWidget(lbl)

        inner_layout.addStretch()
        scroll.setWidget(inner)
        layout.addWidget(scroll, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_ok = QPushButton("확인")
        btn_ok.setObjectName("PrimaryButton")
        btn_ok.clicked.connect(dlg.accept)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)

        dlg.exec()

    def _on_scan_table_double_click(self, row, col):
        path = getattr(self, "_scan_file_paths", {}).get(row)
        if path:
            self._open_file(path)

    def _open_selected_scan_file(self):
        row = self.step3.scan_table.currentRow()
        if row < 0:
            return
        path = getattr(self, "_scan_file_paths", {}).get(row)
        if path:
            self._open_file(path)

    def _open_file(self, path):
        import subprocess, os
        try:
            os.startfile(str(path))
        except AttributeError:
            subprocess.Popen(["xdg-open", str(path)])

    def _show_log(self, title, logs):
        dlg = QDialog(self)
        dlg.setWindowTitle(title)
        dlg.resize(640, 480)
        layout = QVBoxLayout(dlg)
        text = QPlainTextEdit()
        text.setReadOnly(True)
        text.setPlainText("\n".join(logs or ["(로그 없음)"]))
        layout.addWidget(text)
        btn = QPushButton("닫기")
        btn.clicked.connect(dlg.accept)
        layout.addWidget(btn, 0, Qt.AlignmentFlag.AlignRight)
        dlg.exec()

    def _new_work(self):
        reply = QMessageBox.question(
            self, "새 작업", "새 작업을 시작하시겠습니까?\n현재 작업 내용이 초기화됩니다.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply != QMessageBox.StandardButton.Yes:
            return
        self.selected_school   = None
        self._last_scan_result = None
        self._last_run_result  = None
        self._last_scan_logs   = []
        self._last_run_logs    = []
        self._current_outputs  = []
        self.step3.scan_status.setText("스캔을 실행해 주세요.")
        self.step3.btn_next.setEnabled(False)
        self.step4.run_status.setText("스캔을 먼저 완료해 주세요.")
        self.step4.btn_run.setEnabled(False)
        self.step4.btn_next.setEnabled(False)
        self.step5.notice_text.clear()
        self.step5.email_status.setText("")
        for i in range(2, 5):
            self.step_bar.set_step_enabled(i, False)
        self._go_to_step(1)


# ── 실행 ─────────────────────────────────────────────────
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(APP_QSS)
    win = TurnerWindow()
    win.show()
    sys.exit(app.exec())
