"""Reindeer Ledger 0.6.0 RC — интерфейс бухгалтерской ведомости (PySide6).

Исходная бухгалтерская логика: ledger_core.py. XLSX никогда не изменяем.
"""
from __future__ import annotations

import sys
import hashlib
import os
from pathlib import Path

from PySide6.QtCore import (QAbstractTableModel, QModelIndex, QSettings, Qt, QEvent,
                             QSortFilterProxyModel, QUrl, Signal, QTimer, QStandardPaths)
from PySide6.QtGui import (QAction, QDesktopServices, QFont, QKeySequence, QPainter,
                           QPainterPath, QPixmap)
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QComboBox, QDialog, QDialogButtonBox,
    QFileDialog, QFrame, QGridLayout, QHBoxLayout, QHeaderView, QLabel,
    QLineEdit, QMainWindow, QMessageBox, QPlainTextEdit, QPushButton,
    QSizePolicy, QScrollArea, QStackedWidget, QStyledItemDelegate, QTableView, QTableWidget,
    QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget,
)

from checkstate import parse_check_state
from accounting_export import (
    AccountingExportError, GROUPS, SHEET_NAMES, build_accounting_summary,
    compare_with_template, export_accounting_xlsx, page_boundaries_from_text,
)
from ledger_core import (Ledger, MarkupError, compare_purchased, export_pdf,
                         read_ledger, read_purchased_jsonl, short_weight,
                         summarize, write_jsonl)

ROOT = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))
LOGO = ROOT / 'assets' / 'kvramis_logo.jpg'
VERSION = '0.6.0 RC'
COMPANY = 'KV/RaMIS Project Group'
SITE = 'https://kvramispg.vercel.app'
EMAIL = 'kvramisprojectgroup@gmail.com'

PALETTES = {
    'Светлая': dict(bg='#F6F7F9', panel='#FFFFFF', text='#20242B', muted='#626D7B',
                    border='#DDE2E8', soft='#EEF1F5', hover='#FBE9EB', accent='#C72D38',
                    accenthover='#A9222C', table='#FFFFFF', selected='#FBE9EB',
                    subtle='#F5F6F8', disabled='#7C8795'),
    'Тёмная': dict(bg='#15181D', panel='#20242B', text='#F0F2F5', muted='#B3BDC9',
                   border='#424B57', soft='#2C323B', hover='#403038', accent='#ED4D58',
                   accenthover='#D63B47', table='#20242B', selected='#45313A',
                   subtle='#262B33', disabled='#9AA5B1'),
}


def stylesheet(theme: str) -> str:
    c = PALETTES[theme]
    return f'''
    QWidget {{ background:{c['bg']}; color:{c['text']};
              font-family:'Segoe UI','Noto Sans',sans-serif; }}
    QLabel {{ background:transparent; }}
    QLabel[tone="muted"] {{ color:{c['muted']}; }}
    QLabel[tone="danger"] {{ color:{c['accent']}; }}
    QFrame#sidebar, QFrame#card, QFrame#drop {{ background:{c['panel']};
        border:1px solid {c['border']}; border-radius:8px; }}
    QFrame#sidebar {{ border-left:0; border-top:0; border-bottom:0; border-radius:0; }}
    QFrame#drop {{ border:2px dashed {c['border']}; border-radius:9px; }}
    QFrame#drop[hovered="true"] {{ border:2px dashed {c['accent']}; background:{c['hover']}; }}
    QFrame#sidebar QLabel, QFrame#card QLabel, QFrame#drop QLabel {{ background:transparent; }}
    QPushButton {{ background:{c['panel']}; color:{c['text']}; border:1px solid {c['border']};
        border-radius:6px; padding:9px 11px; font-weight:600; }}
    QPushButton:hover {{ background:{c['soft']}; }}
    QPushButton:pressed {{ background:{c['hover']}; }}
    QPushButton:disabled {{ color:{c['disabled']}; background:{c['soft']}; }}
    QPushButton#primary {{ background:{c['accent']}; color:#FFFFFF; border-color:{c['accent']}; }}
    QPushButton#primary:hover {{ background:{c['accenthover']}; border-color:{c['accenthover']}; }}
    QPushButton#primary:disabled {{ background:{c['soft']}; color:{c['disabled']}; border-color:{c['border']}; }}
    QPushButton#nav {{ text-align:left; border:0; background:transparent; padding:11px 10px; font-weight:500; }}
    QPushButton#nav:checked {{ background:{c['hover']}; color:{c['accent']}; font-weight:700; }}
    QPushButton#nav:hover:!checked {{ background:{c['soft']}; }}
    QLineEdit, QComboBox, QPlainTextEdit {{ background:{c['panel']}; color:{c['text']};
        border:1px solid {c['border']}; border-radius:6px; padding:7px 8px;
        selection-background-color:{c['selected']}; selection-color:{c['text']}; }}
    QLineEdit::placeholder {{ color:{c['muted']}; }}
    QComboBox QAbstractItemView {{ background:{c['panel']}; color:{c['text']};
        selection-background-color:{c['selected']}; selection-color:{c['text']}; }}
    QTableView, QTableWidget {{ background:{c['table']}; alternate-background-color:{c['subtle']};
        color:{c['text']}; border:1px solid {c['border']}; gridline-color:{c['border']};
        selection-background-color:{c['selected']}; selection-color:{c['text']}; }}
    QHeaderView::section {{ background:{c['soft']}; color:{c['text']};
        border:0; border-bottom:1px solid {c['border']}; border-right:1px solid {c['border']};
        padding:6px 4px; font-weight:600; }}
    QTableCornerButton::section {{ background:{c['soft']}; }}
    QScrollBar:vertical {{ background:{c['bg']}; width:11px; }}
    QScrollBar::handle:vertical {{ background:{c['border']}; border-radius:5px; min-height:26px; }}
    QScrollBar:horizontal {{ background:{c['bg']}; height:11px; }}
    QScrollBar::handle:horizontal {{ background:{c['border']}; border-radius:5px; min-width:26px; }}
    QScrollBar::add-line, QScrollBar::sub-line {{ height:0px; width:0px; }}
    QTabWidget::pane {{ border:1px solid {c['border']}; background:{c['panel']}; }}
    QTabBar::tab {{ background:{c['soft']}; color:{c['text']}; padding:8px 12px;
       border:1px solid {c['border']}; border-bottom:0; }}
    QTabBar::tab:selected {{ background:{c['panel']}; color:{c['accent']}; font-weight:700; }}
    QToolTip {{ color:{c['text']}; background:{c['panel']}; border:1px solid {c['border']}; }}
    '''


def make_label(text: str, size: int = 13, bold: bool = False,
               muted: bool = False, wrap: bool = True) -> QLabel:
    w = QLabel(text)
    w.setWordWrap(wrap)
    if muted:
        w.setProperty('tone', 'muted')
    f = w.font()
    f.setPixelSize(size)
    f.setBold(bold)
    w.setFont(f)
    return w


def rounded_logo(size: int = 100) -> QPixmap:
    result = QPixmap(size, size)
    result.fill(Qt.GlobalColor.transparent)
    source = QPixmap(str(LOGO))
    if source.isNull():
        return result
    painter = QPainter(result)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    path = QPainterPath()
    path.addEllipse(0, 0, size, size)
    painter.setClipPath(path)
    scaled = source.scaled(size, size, Qt.AspectRatioMode.KeepAspectRatioByExpanding,
                           Qt.TransformationMode.SmoothTransformation)
    painter.drawPixmap((size - scaled.width()) // 2, (size - scaled.height()) // 2, scaled)
    painter.end()
    return result


def stat_text(n: int, kg: float) -> tuple[str, str]:
    return f'{n} гол.', f'{short_weight(float(kg))} кг'


class DropZone(QFrame):
    dropped = Signal(str)

    def __init__(self):
        super().__init__()
        self.setObjectName('drop')
        self.setAcceptDrops(True)
        self.setProperty('hovered', False)
        self.setMinimumHeight(150)
        self.setMaximumHeight(175)
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        for title, size, strong in [
            ('↓', 30, False),
            ('Перетащите ведомость Excel сюда', 17, True),
            ('Файл .xlsx · или воспользуйтесь кнопкой ниже', 12, False),
        ]:
            item = make_label(title, size, strong, muted=not strong)
            item.setAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
            layout.addWidget(item)

    @staticmethod
    def valid(event) -> bool:
        urls = event.mimeData().urls() if event.mimeData().hasUrls() else []
        return (len(urls) == 1 and urls[0].isLocalFile()
                and urls[0].toLocalFile().lower().endswith('.xlsx'))

    def set_hover(self, on: bool):
        self.setProperty('hovered', on)
        self.style().unpolish(self)
        self.style().polish(self)
        self.update()

    def dragEnterEvent(self, event):
        if self.valid(event):
            event.acceptProposedAction()
            self.set_hover(True)
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if self.valid(event):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.set_hover(False)
        event.accept()

    def dropEvent(self, event):
        self.set_hover(False)
        if self.valid(event):
            event.acceptProposedAction()
            self.dropped.emit(event.mimeData().urls()[0].toLocalFile())
        else:
            event.ignore()


class AnimalModel(QAbstractTableModel):
    purchase_changed = Signal()
    HEADERS = ('Закуп', '№', 'Федеральный номер', 'Бирка',
               'Категория', 'Направление', 'Вес, кг')

    def __init__(self, parent=None):
        super().__init__(parent)
        self.animals = []
        self.purchased: set[int] = set()

    def reset(self, ledger: Ledger | None, purchased: set[int]):
        self.beginResetModel()
        self.animals = ledger.animals if ledger is not None else []
        self.purchased = purchased  # разделяем один set с окном
        self.endResetModel()

    def rowCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.animals)

    def columnCount(self, parent=QModelIndex()):
        return 0 if parent.isValid() else len(self.HEADERS)

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole and orientation == Qt.Orientation.Horizontal:
            return self.HEADERS[section]
        return None

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or index.row() >= len(self.animals):
            return None
        a = self.animals[index.row()]
        col = index.column()
        if role == Qt.ItemDataRole.CheckStateRole and col == 0:
            return Qt.CheckState.Checked if a.number in self.purchased else Qt.CheckState.Unchecked
        if role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.EditRole):
            return (
                '', str(a.number), a.fed_number or '—', a.tag or '—',
                a.category, 'Реализация' if a.free_weight is not None else
                'Переработка' if a.industrial_weight is not None else '—',
                short_weight(float(a.slaughter_weight)),
            )[col]
        if role == Qt.ItemDataRole.UserRole:
            return (int(a.number in self.purchased), a.number, a.fed_number.casefold(),
                    a.tag.casefold(), a.category.casefold(), a.operation.casefold(),
                    a.slaughter_weight)[col]
        if role == Qt.ItemDataRole.TextAlignmentRole:
            if col in (0, 1):
                return Qt.AlignmentFlag.AlignCenter
            if col == 6:
                return Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight
            return Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft
        if role == Qt.ItemDataRole.ToolTipRole:
            if col == 0:
                return 'Отметить животное как закупное (по № п/п)'
            if col == 2:
                return f'Федеральный номер: {a.fed_number or "—"}'
            if col == 5:
                return a.operation
            return f'№ {a.number} · {a.category} · {short_weight(float(a.slaughter_weight))} кг'
        return None

    def flags(self, index):
        if not index.isValid():
            return Qt.ItemFlag.NoItemFlags
        result = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
        if index.column() == 0:
            result |= Qt.ItemFlag.ItemIsUserCheckable
        return result

    def setData(self, index, value, role=Qt.ItemDataRole.EditRole):
        if (not index.isValid() or index.column() != 0
                or role != Qt.ItemDataRole.CheckStateRole):
            return False
        number = self.animals[index.row()].number
        should_be_checked = parse_check_state(value)
        if should_be_checked is None:
            return False
        already = number in self.purchased
        if should_be_checked == already:
            return True
        if should_be_checked:
            self.purchased.add(number)
        else:
            self.purchased.discard(number)
        self.dataChanged.emit(index, index,
                              [Qt.ItemDataRole.CheckStateRole, Qt.ItemDataRole.UserRole])
        self.purchase_changed.emit()
        return True

    def refresh_checks(self):
        if self.animals:
            left = self.index(0, 0)
            right = self.index(len(self.animals) - 1, 0)
            self.dataChanged.emit(left, right,
                                  [Qt.ItemDataRole.CheckStateRole, Qt.ItemDataRole.UserRole])


class PurchaseDelegate(QStyledItemDelegate):
    """Один щелчок в любой точке ячейки «Закуп» переключает галочку.

    В отличие от стандартного делегата, не нужно попадать в маленький квадрат.
    Через прокси передаётся Qt.CheckState, а не int(enum).
    """

    def editorEvent(self, event, model, option, index):
        if not index.isValid() or index.column() != 0:
            return super().editorEvent(event, model, option, index)
        if not (index.flags() & Qt.ItemFlag.ItemIsEnabled):
            return False
        mouse_toggle = (event.type() == QEvent.Type.MouseButtonRelease
                        and event.button() == Qt.MouseButton.LeftButton)
        key_toggle = (event.type() == QEvent.Type.KeyPress
                      and event.key() in (Qt.Key.Key_Space, Qt.Key.Key_Select))
        if not mouse_toggle and not key_toggle:
            return False
        current = parse_check_state(model.data(index, Qt.ItemDataRole.CheckStateRole))
        if current is None:
            return False
        desired = Qt.CheckState.Unchecked if current else Qt.CheckState.Checked
        return bool(model.setData(index, desired, Qt.ItemDataRole.CheckStateRole))


class AnimalFilter(QSortFilterProxyModel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.query = ''
        self.category = 'Все категории'
        self.purchase_filter = 'Все животные'
        self.setSortRole(Qt.ItemDataRole.UserRole)

    def filterAcceptsRow(self, source_row, source_parent):
        model = self.sourceModel()
        if model is None:
            return False
        a = model.animals[source_row]
        if self.query and self.query not in ' '.join((
            str(a.number), a.fed_number, a.tag, a.category, a.operation,
        )).casefold():
            return False
        if self.category != 'Все категории' and self.category != a.category:
            return False
        bought = a.number in model.purchased
        if self.purchase_filter == 'Только закупные' and not bought:
            return False
        if self.purchase_filter == 'Остальные' and bought:
            return False
        return True

    def update_filter(self, query, category, purchase_filter):
        self.query = query.casefold().strip()
        self.category = category
        self.purchase_filter = purchase_filter
        self.refresh_filter()

    def refresh_filter(self):
        # Qt >= 6.10: invalidateFilter() устарел. На Qt 6.7–6.9
        # продолжаем пользоваться старым API ради совместимости.
        if hasattr(self, 'beginFilterChange') and hasattr(self, 'endFilterChange'):
            self.beginFilterChange()
            self.endFilterChange()
        else:
            self.invalidateFilter()


class StatCard(QFrame):
    def __init__(self, heading: str):
        super().__init__()
        self.setObjectName('card')
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(13, 10, 13, 10)
        lay.setSpacing(3)
        lay.addWidget(make_label(heading, 12, muted=True))
        self.primary = make_label('—', 19, True, wrap=False)
        self.secondary = make_label('', 12, muted=True, wrap=False)
        lay.addWidget(self.primary)
        lay.addWidget(self.secondary)

    def display(self, n: int, kg: float):
        p, s = stat_text(n, kg)
        self.primary.setText(p)
        self.secondary.setText(s)


class Window(QMainWindow):
    NAV = ('Главная', 'Загрузка ведомости', 'Проверка данных', 'Отчёты',
           'История', 'Обучение', 'Настройки', 'О программе')

    def __init__(self):
        super().__init__()
        self.settings = QSettings('KVRaMIS', 'ReindeerLedger')
        self.theme = self.settings.value('theme', 'Светлая')
        if self.theme not in PALETTES:
            self.theme = 'Светлая'
        self.ledger: Ledger | None = None
        self.purchased: set[int] = set()
        self.dirty = False
        self.saved_markup_path: Path | None = None
        self.recent_paths = self.load_recent()
        self.autosave_timer = QTimer(self)
        self.autosave_timer.setSingleShot(True)
        self.autosave_timer.setInterval(750)
        self.autosave_timer.timeout.connect(self.autosave_markup)
        self.setWindowTitle(f'Reindeer Ledger {VERSION} — {COMPANY}')
        self.setMinimumSize(900, 620)
        self.resize(1150, 740)

        root = QWidget()
        self.setCentralWidget(root)
        outer = QHBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        side = QFrame()
        side.setObjectName('sidebar')
        side.setFixedWidth(214)
        side_layout = QVBoxLayout(side)
        side_layout.setContentsMargins(17, 22, 17, 17)
        side_layout.setSpacing(7)
        logo = QLabel()
        logo.setPixmap(rounded_logo(98))
        side_layout.addWidget(logo)
        side_layout.addWidget(make_label('Reindeer Ledger', 18, True))
        side_layout.addWidget(make_label('Учёт ведомостей', 12, muted=True))
        side_layout.addSpacing(18)
        self.nav_buttons = []
        for ix, name in enumerate(self.NAV):
            btn = QPushButton(f'{ix + 1:02d}   {name}')
            btn.setObjectName('nav')
            btn.setCheckable(True)
            btn.clicked.connect(lambda checked=False, index=ix: self.navigate(index))
            side_layout.addWidget(btn)
            self.nav_buttons.append(btn)
        side_layout.addStretch()
        side_layout.addWidget(make_label(COMPANY, 11, muted=True))
        side_layout.addWidget(make_label('Версия ' + VERSION, 11, muted=True))
        outer.addWidget(side)

        center = QWidget()
        ml = QVBoxLayout(center)
        ml.setContentsMargins(28, 21, 28, 15)
        ml.setSpacing(11)
        upper = QHBoxLayout()
        upper.addStretch()
        upper.addWidget(make_label('Тема: ', 12, muted=True))
        self.theme_switch = QComboBox()
        self.theme_switch.addItems(PALETTES)
        self.theme_switch.setCurrentText(self.theme)
        self.theme_switch.setToolTip('Цветовая тема интерфейса')
        self.theme_switch.currentTextChanged.connect(self.set_theme)
        upper.addWidget(self.theme_switch)
        ml.addLayout(upper)
        self.pages = QStackedWidget()
        ml.addWidget(self.pages, 1)
        self.build_home()
        self.build_load()
        self.build_review()
        self.build_reports()
        self.build_history()
        self.build_guide()
        self.build_settings()
        self.build_about()
        self.status = make_label('Готово к работе', 11, muted=True, wrap=False)
        ml.addWidget(self.status)
        outer.addWidget(center, 1)
        self.set_theme(self.theme)
        self.navigate(0)
        self.refresh_home()
        self.refresh_history()
        for shortcut, callback in (('Ctrl+O', self.pick_file),
                                   ('Ctrl+S', self.save_jsonl),
                                   ('Ctrl+F', self.focus_search)):
            action = QAction(self)
            action.setShortcut(QKeySequence(shortcut))
            action.triggered.connect(callback)
            self.addAction(action)

    def focus_search(self):
        if self.ledger is not None:
            self.navigate(2)
            self.search.setFocus()
            self.search.selectAll()
        else:
            self.navigate(1)

    @staticmethod
    def button(text, callback, primary=False):
        btn = QPushButton(text)
        if primary:
            btn.setObjectName('primary')
        btn.clicked.connect(callback)
        return btn

    def page(self, title, description):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        layout.addWidget(make_label(title, 26, True))
        layout.addWidget(make_label(description, 12, muted=True))
        self.pages.addWidget(page)
        return layout

    @staticmethod
    def card(layout, *widgets):
        card = QFrame()
        card.setObjectName('card')
        inn = QVBoxLayout(card)
        inn.setContentsMargins(14, 12, 14, 12)
        inn.setSpacing(7)
        for w in widgets:
            inn.addWidget(w)
        layout.addWidget(card)
        return card

    def build_home(self):
        lay = self.page('Главная', 'Ведомости и проверка закупных — без командной строки.')
        lay.addSpacing(10)
        self.home_title = make_label('Начните с загрузки Excel-ведомости.', 19, True)
        self.home_description = make_label('Откройте файл .xlsx, проверьте животных и сохраните разметку.')
        self.card(lay, self.home_title, self.home_description)
        row = QHBoxLayout()
        row.addWidget(self.button('Открыть ведомость →', lambda: self.navigate(1), True))
        row.addStretch()
        lay.addLayout(row)
        lay.addSpacing(12)
        lay.addWidget(make_label('Как работать с ведомостью', 16, True))
        workflow = QHBoxLayout()
        workflow.setSpacing(8)
        for title, detail in (
            ('01  Загрузка', 'Выберите Excel-файл'),
            ('02  Проверка', 'Отметьте закупных'),
            ('03  Результат', 'Сохраните бухгалтерский Excel'),
        ):
            panel = QFrame()
            panel.setObjectName('card')
            content = QVBoxLayout(panel)
            content.setContentsMargins(11, 11, 11, 11)
            content.addWidget(make_label(title, 13, True))
            content.addWidget(make_label(detail, 12, muted=True))
            workflow.addWidget(panel, 1)
        lay.addLayout(workflow)
        lay.addSpacing(12)
        lay.addWidget(make_label('Последние файлы', 16, True))
        self.home_recent = make_label('Пока нет недавно открытых ведомостей.', 13, muted=True)
        lay.addWidget(self.home_recent)
        lay.addStretch(1)

    def build_load(self):
        lay = self.page('Загрузка ведомости', 'Перетащите XLSX или выберите файл на компьютере. Оригинал не изменяется.')
        self.drop = DropZone()
        self.drop.dropped.connect(self.load_file)
        lay.addWidget(self.drop)
        btnrow = QHBoxLayout()
        btnrow.addWidget(self.button('Выбрать XLSX…', self.pick_file))
        self.remove_file_button = self.button('Убрать выбранный файл', self.remove_file)
        self.remove_file_button.setEnabled(False)
        btnrow.addWidget(self.remove_file_button)
        btnrow.addStretch()
        lay.addLayout(btnrow)
        self.file_panel = QFrame()
        self.file_panel.setObjectName('card')
        inner = QVBoxLayout(self.file_panel)
        inner.setContentsMargins(14, 11, 14, 11)
        self.loaded_info = make_label('Файл не выбран.', 13, True)
        inner.addWidget(self.loaded_info)
        self.warning_info = make_label('', 12, muted=True)
        inner.addWidget(self.warning_info)
        self.file_panel.setVisible(False)
        lay.addWidget(self.file_panel)
        lay.addStretch()
        controls = QHBoxLayout()
        controls.addStretch()
        self.next_button = self.button('Перейти к проверке →', lambda: self.navigate(2), True)
        self.next_button.setEnabled(False)
        controls.addWidget(self.next_button)
        lay.addLayout(controls)

    def build_review(self):
        lay = self.page('Проверка данных', 'Закупных отмечайте галочкой. Каждая строка — одно животное из XLSX.')
        filters = QHBoxLayout()
        self.search = QLineEdit()
        self.search.setPlaceholderText('Поиск по №, федеральному номеру, бирке…')
        self.search.textChanged.connect(self.filter_rows)
        filters.addWidget(self.search, 1)
        self.category = QComboBox()
        self.category.addItem('Все категории')
        self.category.currentTextChanged.connect(self.filter_rows)
        filters.addWidget(self.category)
        self.purchase_filter = QComboBox()
        self.purchase_filter.addItems(('Все животные', 'Только закупные', 'Остальные'))
        self.purchase_filter.currentTextChanged.connect(self.filter_rows)
        filters.addWidget(self.purchase_filter)
        lay.addLayout(filters)

        self.animal_model = AnimalModel(self)
        self.proxy = AnimalFilter(self)
        self.proxy.setSourceModel(self.animal_model)
        self.animal_model.purchase_changed.connect(self.purchase_changed)
        self.table = QTableView()
        self.table.setModel(self.proxy)
        self.purchase_delegate = PurchaseDelegate(self.table)
        self.table.setItemDelegateForColumn(0, self.purchase_delegate)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.AllEditTriggers)
        self.table.setSortingEnabled(True)
        self.table.sortByColumn(1, Qt.SortOrder.AscendingOrder)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.horizontalHeader().setMinimumSectionSize(38)
        head = self.table.horizontalHeader()
        head.setDefaultAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        head.setStretchLastSection(False)
        head.setSectionsMovable(False)
        for i in range(7):
            head.setSectionResizeMode(i, QHeaderView.ResizeMode.Fixed)
        head.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        for col, width in {0: 57, 1: 43, 3: 68, 4: 131, 5: 143, 6: 77}.items():
            self.table.setColumnWidth(col, width)
        self.table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table.setToolTip('Наведите курсор на номер, чтобы увидеть его полностью. Нажмите галочку для отметки закупного.')
        lay.addWidget(self.table, 1)

        metrics = QHBoxLayout()
        metrics.setSpacing(8)
        self.all_card = StatCard('Всего')
        self.bought_card = StatCard('Закупные')
        self.rest_card = StatCard('Остальные')
        for card in (self.all_card, self.bought_card, self.rest_card):
            metrics.addWidget(card, 1)
        lay.addLayout(metrics)
        self.filtered = make_label('Сначала загрузите ведомость.', 11, muted=True)
        lay.addWidget(self.filtered)
        buttons = QHBoxLayout()
        buttons.addWidget(self.button('← К загрузке', lambda: self.navigate(1)))
        self.import_button = self.button('Загрузить разметку…', self.import_markup)
        buttons.addWidget(self.import_button)
        self.save_button = self.button('Сохранить JSONL…', self.save_jsonl)
        buttons.addWidget(self.save_button)
        buttons.addStretch()
        buttons.addWidget(self.button('К отчётам →', lambda: self.navigate(3), True))
        lay.addLayout(buttons)
        self.import_button.setEnabled(False)
        self.save_button.setEnabled(False)

    def build_reports(self):
        lay = self.page('Отчёты', 'Готовая бухгалтерская ведомость по образцу: ТС, ТВ, В, Б и закупные З.')
        stats = QHBoxLayout()
        stats.setSpacing(8)
        self.report_all = StatCard('Всего животных')
        self.report_bought = StatCard('Закупные')
        self.report_rest = StatCard('Остальные')
        for card in (self.report_all, self.report_bought, self.report_rest):
            stats.addWidget(card, 1)
        lay.addLayout(stats)
        self.report_directions = make_label('Откройте ведомость для просмотра итогов.', 12, muted=True)
        lay.addWidget(self.report_directions)

        lay.addWidget(make_label('Разбивка по страницам исходного PDF', 15, True))
        page_row = QHBoxLayout()
        page_row.addWidget(make_label('Страниц:', 12))
        self.page_count = QComboBox()
        self.page_count.addItems(('1', '2', '3'))
        self.page_count.setToolTip('Количество страниц с животными в оригинальном PDF')
        page_row.addWidget(self.page_count)
        page_row.addWidget(make_label('Последние № п/п страниц:', 12))
        self.page_boundaries = QLineEdit()
        self.page_boundaries.setPlaceholderText('Например: 44, 94')
        self.page_boundaries.setToolTip('Для трёх страниц 1–44, 45–94, 95–102 укажите: 44, 94')
        page_row.addWidget(self.page_boundaries, 1)
        lay.addLayout(page_row)
        self.report_pages_hint = make_label(
            'Границы задаются по № п/п животного, а не по номеру строки Excel.', 11, muted=True)
        lay.addWidget(self.report_pages_hint)
        self.page_count.currentTextChanged.connect(self.refresh_export_preview)
        self.page_boundaries.textChanged.connect(self.refresh_export_preview)

        lay.addWidget(make_label('ТС — телята-самцы   ·   ТВ — телята-самки   ·   В — важенки   ·   Б — быки   ·   З — закупные', 11, muted=True))
        self.report_preview = QTabWidget()
        self.report_tables: dict[str, QTableWidget] = {}
        for sheet_name in SHEET_NAMES:
            grid = QTableWidget(6, 5)
            grid.setHorizontalHeaderLabels(('Группа', 'I · голов', 'I · кг', 'II · голов', 'II · кг'))
            grid.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
            grid.verticalHeader().setVisible(False)
            grid.verticalHeader().setDefaultSectionSize(28)
            grid.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
            grid.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
            grid.setAlternatingRowColors(True)
            grid.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
            grid.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
            grid.setColumnWidth(0, 88)
            grid.setFocusPolicy(Qt.FocusPolicy.NoFocus)
            self.report_tables[sheet_name] = grid
            self.report_preview.addTab(grid, sheet_name)
        lay.addWidget(self.report_preview, 1)
        self.export_feedback = make_label('Откройте ведомость, чтобы рассчитать экспорт.', 12, muted=True)
        lay.addWidget(self.export_feedback)
        actions = QHBoxLayout()
        self.export_excel_button = self.button('Сохранить бухгалтерский Excel…', self.save_accounting, True)
        self.export_excel_button.setEnabled(False)
        actions.addWidget(self.export_excel_button)
        actions.addWidget(self.button('Сформировать PDF-реестр…', self.save_pdf))
        actions.addWidget(self.button('Сохранить разметку JSONL…', self.save_jsonl))
        actions.addStretch()
        lay.addLayout(actions)

    def set_accounting_page_defaults(self):
        if self.ledger is None:
            return
        # Эти две границы подтверждены исходным трёхстраничным PDF от 17.11.2026.
        if self.ledger.date == '17.11.2026' and len(self.ledger.animals) == 102:
            count, boundaries = '3', '44, 94'
            hint = '17.11.2026: стр. 1 — №1–44; стр. 2 — №45–94; стр. 3 — №95–102.'
        elif self.ledger.date == '18.11.2026' and len(self.ledger.animals) == 70:
            count, boundaries = '2', ''
            hint = 'Для 18.11.2026 укажите последний № п/п на первой странице исходного PDF.'
        else:
            count, boundaries = '1', ''
            hint = 'По умолчанию одна страница. Если в исходном PDF их больше — укажите число и границы.'
        self.page_count.blockSignals(True)
        self.page_boundaries.blockSignals(True)
        self.page_count.setCurrentText(count)
        self.page_boundaries.setText(boundaries)
        self.page_count.blockSignals(False)
        self.page_boundaries.blockSignals(False)
        self.report_pages_hint.setText(hint)
        self.refresh_export_preview()

    def current_accounting_summary(self):
        if self.ledger is None:
            raise AccountingExportError('Сначала откройте исходный XLSX.')
        boundaries = page_boundaries_from_text(self.page_boundaries.text(), int(self.page_count.currentText()))
        return build_accounting_summary(self.ledger, self.purchased, boundaries)

    def refresh_export_preview(self, *_):
        if not hasattr(self, 'export_feedback'):
            return
        if self.ledger is None:
            self.export_excel_button.setEnabled(False)
            return
        try:
            summary = self.current_accounting_summary()
        except (AccountingExportError, ValueError) as exc:
            self.export_feedback.setText('Нужна проверка границ страниц: ' + str(exc))
            self.export_feedback.setProperty('tone', 'danger')
            self.export_feedback.style().unpolish(self.export_feedback)
            self.export_feedback.style().polish(self.export_feedback)
            self.export_excel_button.setEnabled(False)
            return
        for name, grid in self.report_tables.items():
            values = summary.sheets[name] + [list(summary.totals[name])]
            for row, (group, data) in enumerate(zip((*GROUPS, 'ИТОГО'), values)):
                for col, value in enumerate((group, *data)):
                    text = str(value) if col == 0 else short_weight(float(value))
                    item = QTableWidgetItem(text)
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter if col == 0
                                          else Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                    if row == 5:
                        font = item.font()
                        font.setBold(True)
                        item.setFont(font)
                    grid.setItem(row, col, item)
        self.export_excel_button.setEnabled(True)
        note = (f'Проверено: {summary.count} гол., {short_weight(summary.weight)} кг. '
                f'Закупные: {summary.purchased_count} гол., {short_weight(summary.purchased_weight)} кг. '
                'Строки ИТОГО и суммы страниц сходятся.')
        if (self.ledger.date == '17.11.2026' and len(self.ledger.animals) == 102
                and summary.page_limits == (44, 94)):
            try:
                changes = compare_with_template(summary)
            except (OSError, ValueError) as exc:
                changes = []
                note += ' Не удалось сравнить с образцом: ' + str(exc)
            if changes:
                note += (f' Внимание: {len(changes)} ячеек отличаются от готового образца. '
                         'Разметка закупных НЕ меняется автоматически.')
        self.export_feedback.setText(note)
        self.export_feedback.setProperty('tone', 'muted')
        self.export_feedback.style().unpolish(self.export_feedback)
        self.export_feedback.style().polish(self.export_feedback)

    def save_accounting(self):
        if self.ledger is None:
            QMessageBox.information(self, 'Нет данных', 'Сначала откройте ведомость XLSX.')
            return
        if not self.confirm_empty_markup():
            return
        try:
            summary = self.current_accounting_summary()
            if (self.ledger.date == '17.11.2026' and len(self.ledger.animals) == 102
                    and summary.page_limits == (44, 94)):
                changes = compare_with_template(summary)
                if changes:
                    message = (
                        f'Текущая разметка отличается от готового образца в {len(changes)} ячейках.\n\n'
                        'В образце закупные: 26 голов, 867 кг.\n'
                        f'Сейчас закупные: {summary.purchased_count} голов, '
                        f'{short_weight(summary.purchased_weight)} кг.\n\n'
                        'Экспортировать именно текущие отмеченные галочки? '
                        'Программа не будет подгонять данные под образец.')
                    answer = QMessageBox.question(
                        self, 'Расхождение с бухгалтерским образцом', message,
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                        QMessageBox.StandardButton.No)
                    if answer != QMessageBox.StandardButton.Yes:
                        return
        except (AccountingExportError, ValueError, OSError) as exc:
            QMessageBox.warning(self, 'Проверка экспорта не пройдена', str(exc))
            return
        default = self.ledger.path.parent / (
            self.ledger.date.replace('.', '-') + '_' + self.ledger.path.stem + '_бухгалтерия.xlsx')
        path, _ = QFileDialog.getSaveFileName(
            self, 'Сохранить бухгалтерский Excel', str(default), 'Excel (*.xlsx)')
        if not path:
            return
        destination = Path(path)
        if destination.suffix.lower() != '.xlsx':
            destination = destination.with_suffix('.xlsx')
        try:
            export_accounting_xlsx(summary, destination, source=self.ledger.path)
        except (AccountingExportError, OSError, ValueError) as exc:
            QMessageBox.critical(self, 'Ошибка экспорта Excel', str(exc))
            return
        self.status.setText(f'Бухгалтерская ведомость сохранена: {destination.name}')
        QMessageBox.information(
            self, 'Excel сохранён', f'Готовая бухгалтерская ведомость:\n{destination}\n\n'
            'Все четыре листа сохранены в исходном оформлении. '
            'Исходный XLSX не изменён.\n\n'
            'Для сохранения самих отметок закупных дополнительно используйте JSONL.')

    def build_history(self):
        lay = self.page('История', 'Недавно открытые файлы на этом компьютере.')
        self.history_info = make_label('Пока нет открытых ведомостей.', 13)
        lay.addWidget(self.history_info)
        self.history_actions = QVBoxLayout()
        self.history_actions.setSpacing(6)
        lay.addLayout(self.history_actions)
        lay.addWidget(make_label('Последние пути сохраняются локально. Содержимое ведомостей в историю не копируется.', 12, muted=True))
        lay.addStretch()

    def build_guide(self):
        lay = self.page('Обучение', 'Пошаговая инструкция для бухгалтера. Интернет не требуется.')
        scroller = QScrollArea()
        scroller.setWidgetResizable(True)
        guide_content = QWidget()
        guide_layout = QVBoxLayout(guide_content)
        guide_layout.setContentsMargins(3, 3, 8, 3)
        instructions = (
            ('1. Загрузка ведомости',
             'Откройте раздел «Загрузка ведомости». Перетащите файл Excel .xlsx в окно '
             'или нажмите «Выбрать XLSX…». Если ошиблись файлом, нажмите «Убрать выбранный файл». '
             'Исходный Excel остаётся без изменений.'),
            ('2. Проверка животных',
             'Нажмите «Перейти к проверке». Каждая строка таблицы соответствует одному животному. '
             'Поставьте галочку «Закуп» для закупного оленя. Ищите по № п/п, федеральному номеру или бирке. '
             'Обратите внимание: федеральные номера телят могут повторяться — ориентируйтесь на № п/п.'),
            ('3. Сохранение работы',
             'Отметки автоматически сохраняются в резервную копию на этом компьютере. '
             'Для передачи или архивирования нажмите «Сохранить разметку JSONL…». '
             'При повторном открытии ведомости программа предложит восстановить резервную копию.'),
            ('4. Проверка итогов',
             'В разделе «Отчёты» проверьте количество голов и убойный вес, в том числе закупных. '
             'Сверьте распределение по страницам с бумажной ведомостью. '
             'Не меняйте галочки только ради совпадения с чужим примером.'),
            ('5. Получение готовых документов',
             'Нажмите «Сохранить бухгалтерский Excel…». '
             'Отчёт сохраняется отдельным файлом; исходная ведомость не переписывается. '
             'При необходимости сформируйте PDF-реестр.'),
            ('6. Если возникла ошибка',
             'Проверьте, что выбран файл .xlsx нужной структуры. '
             'Если Excel открыт в другой программе, закройте его и повторите сохранение. '
             'По вопросам работы приложения пишите на ' + EMAIL),
        )
        for heading, explanation in instructions:
            self.card(guide_layout, make_label(heading, 14, True),
                      make_label(explanation, 12, wrap=True))
        guide_layout.addStretch()
        scroller.setWidget(guide_content)
        lay.addWidget(scroller, 1)

    def build_settings(self):
        lay = self.page('Настройки', 'Внешний вид и работа с данными.')
        self.card(lay,
                  make_label('Темы оформления', 16, True),
                  make_label('Светлая и тёмная темы переключаются в правом верхнем углу. Выбор запоминается.'),
                  make_label('В тёмной теме второстепенные надписи имеют повышенный контраст.', 12, muted=True))
        self.card(lay,
                  make_label('Безопасность данных', 16, True),
                  make_label('Исходный Excel не изменяется. Разметка хранится отдельно в JSONL.'),
                  make_label('При импорте JSONL программа сравнивает дату, все записи и веса с текущим XLSX.', 12, muted=True))
        lay.addStretch()

    def build_about(self):
        lay = self.page('О программе', f'Reindeer Ledger {VERSION} — бухгалтерский учёт ведомостей')
        scroller = QScrollArea()
        scroller.setWidgetResizable(True)
        about_content = QWidget()
        about_layout = QVBoxLayout(about_content)
        about_layout.setContentsMargins(3, 3, 8, 3)
        self.card(about_layout,
                  make_label('Reindeer Ledger', 21, True),
                  make_label('Приложение для локальной обработки ведомостей учёта оленей, '
                             'разметки закупных животных и подготовки бухгалтерских отчётов.', 13, wrap=True),
                  make_label('Версия: ' + VERSION + '  |  Платформа: Windows 11 / Linux', 12, muted=True))
        self.card(about_layout,
                  make_label('Разработчик и организация', 16, True),
                  make_label('KV/RaMIS Project Group', 14, True),
                  make_label('Главный разработчик: Кравченко Валерий'),
                  make_label('Официальный сайт: ' + SITE, 12, wrap=True),
                  make_label('Почта для вопросов и предложений: ' + EMAIL, 12, wrap=True))
        self.card(about_layout,
                  make_label('Возможности', 16, True),
                  make_label('Загрузка XLSX; поиск и проверка животных; ручная разметка закупных; '
                             'сводка по группам ТС, ТВ, В, Б и З; бухгалтерский Excel по шаблону; '
                             'PDF-реестр; резервное сохранение разметки; две темы оформления.', 12, wrap=True))
        self.card(about_layout,
                  make_label('Конфиденциальность и ограничения', 16, True),
                  make_label('Данные обрабатываются локально и не отправляются на сервер. '
                             'Исходный XLSX не изменяется. Проверяйте результат по первичным документам. '
                             'Автоматическое распознавание рукописных PDF пока не поддерживается.', 12, wrap=True))
        buttons = QHBoxLayout()
        buttons.addWidget(self.button('Сайт компании ↗', lambda: QDesktopServices.openUrl(QUrl(SITE))))
        buttons.addWidget(self.button('Написать разработчику ↗', lambda: QDesktopServices.openUrl(QUrl('mailto:' + EMAIL))))
        buttons.addStretch()
        about_layout.addLayout(buttons)
        about_layout.addStretch()
        scroller.setWidget(about_content)
        lay.addWidget(scroller, 1)

    def set_theme(self, theme):
        if theme not in PALETTES:
            return
        self.theme = theme
        self.settings.setValue('theme', theme)
        self.setStyleSheet(stylesheet(theme))
        self.table.viewport().update()

    def navigate(self, index):
        if index in (2, 3) and self.ledger is None:
            QMessageBox.information(self, 'Нужна ведомость',
                                    'Для начала откройте Excel-ведомость (.xlsx).')
            index = 1
        self.pages.setCurrentIndex(index)
        for i, btn in enumerate(self.nav_buttons):
            btn.setChecked(i == index)
        if index == 3:
            self.refresh_report()
        if index == 0:
            self.refresh_home()
        if index == 4:
            self.refresh_history()

    def pick_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, 'Выберите ведомость', str(Path.home()), 'Excel (*.xlsx)')
        if path:
            self.load_file(path)

    def confirm_discard(self) -> bool:
        if not self.dirty:
            return True
        answer = QMessageBox.question(
            self, 'Несохранённая разметка',
            'Текущая разметка изменена и не сохранена. Открыть другую ведомость и потерять эти изменения?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No)
        return answer == QMessageBox.StandardButton.Yes

    def backup_path(self, ledger=None) -> Path:
        ledger = ledger or self.ledger
        if ledger is None:
            raise ValueError('Нет открытой ведомости')
        root = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.AppDataLocation)
        folder = Path(root) if root else Path.home() / '.reindeer-ledger'
        digest = hashlib.sha256(str(ledger.path.resolve()).encode('utf-8')).hexdigest()[:20]
        return folder / 'autosave' / (digest + '.jsonl')

    def autosave_markup(self):
        if self.ledger is None or not self.dirty:
            return
        try:
            backup = self.backup_path()
            backup.parent.mkdir(parents=True, exist_ok=True)
            temporary = backup.with_suffix('.jsonl.tmp')
            write_jsonl(self.ledger, temporary, self.purchased)
            os.replace(temporary, backup)
            self.status.setText('Изменения автоматически сохранены в резервную копию')
        except Exception as exc:
            self.status.setText('Не удалось создать резервную копию: ' + str(exc))

    def restore_autosave(self):
        if self.ledger is None:
            return
        backup = self.backup_path()
        if not backup.exists():
            return
        try:
            incoming = read_purchased_jsonl(self.ledger, backup)
        except (OSError, ValueError, MarkupError) as exc:
            QMessageBox.warning(self, 'Резервная копия не подходит',
                                'Автоматическая разметка не применена:\n' + str(exc))
            return
        answer = QMessageBox.question(
            self, 'Восстановить разметку?',
            f'Для этой ведомости найдена резервная копия: {len(incoming)} закупных.\n'
            'Восстановить галочки из неё?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes)
        if answer == QMessageBox.StandardButton.Yes:
            self.purchased.clear()
            self.purchased.update(incoming)
            self.animal_model.refresh_checks()
            self.proxy.refresh_filter()
            self.dirty = True
            self.refresh_summary()
            self.refresh_report()
            self.filtered_info()
            self.status.setText('Разметка восстановлена из резервной копии')

    def remove_file(self):
        if self.ledger is None:
            return
        if not self.confirm_discard():
            return
        self.autosave_timer.stop()
        self.ledger = None
        self.purchased = set()
        self.dirty = False
        self.saved_markup_path = None
        self.animal_model.reset(None, self.purchased)
        self.file_panel.setVisible(False)
        self.loaded_info.setText('Файл не выбран.')
        self.warning_info.clear()
        self.next_button.setEnabled(False)
        self.remove_file_button.setEnabled(False)
        self.import_button.setEnabled(False)
        self.save_button.setEnabled(False)
        self.refresh_summary()
        self.refresh_home()
        self.status.setText('Ведомость убрана. Можно выбрать другой XLSX.')
        self.navigate(1)

    def load_file(self, path):
        file = Path(path)
        if self.ledger is not None and not self.confirm_discard():
            return
        try:
            if file.suffix.lower() != '.xlsx':
                raise ValueError('Поддерживается только формат .xlsx.')
            ledger = read_ledger(file)
        except Exception as exc:
            QMessageBox.critical(self, 'Не удалось открыть ведомость',
                                 'Проверьте структуру XLSX.\n\n' + str(exc))
            return
        self.autosave_timer.stop()
        self.ledger = ledger
        self.purchased = set()
        self.dirty = False
        self.saved_markup_path = None
        self.animal_model.reset(ledger, self.purchased)
        self.category.blockSignals(True)
        self.category.clear()
        self.category.addItem('Все категории')
        self.category.addItems(sorted({a.category for a in ledger.animals}))
        self.category.blockSignals(False)
        self.search.clear()
        self.purchase_filter.setCurrentIndex(0)
        self.filter_rows()
        self.refresh_summary()
        self.set_accounting_page_defaults()
        self.next_button.setEnabled(True)
        self.remove_file_button.setEnabled(True)
        self.import_button.setEnabled(True)
        self.save_button.setEnabled(True)
        total = summarize(ledger.animals)
        self.loaded_info.setText(
            f'✓ {file.name}\n{ledger.date} · {len(ledger.animals)} голов · '
            f'{short_weight(float(total["slaughter_weight"]))} кг убойного веса')
        self.warning_info.setText(
            '\n'.join('Внимание: ' + item for item in ledger.warnings)
            if ledger.warnings else 'Проверка контрольных сумм: замечаний нет.')
        self.file_panel.setVisible(True)
        self.record_recent(file)
        self.refresh_history()
        self.refresh_home()
        self.status.setText(f'Загружено: {file.name} · {len(ledger.animals)} голов')
        self.navigate(1)
        self.restore_autosave()

    def purchase_changed(self):
        self.dirty = True
        self.autosave_timer.start()
        self.proxy.refresh_filter()
        self.refresh_summary()
        self.refresh_report()
        self.filtered_info()
        self.status.setText('Разметка изменена — сохраните JSONL')

    def filter_rows(self, *_):
        if self.ledger is None:
            return
        self.proxy.update_filter(self.search.text(), self.category.currentText(),
                                 self.purchase_filter.currentText())
        self.filtered_info()

    def filtered_info(self):
        if self.ledger is not None:
            self.filtered.setText(
                f'Показано {self.proxy.rowCount()} из {len(self.ledger.animals)} животных. '
                'Отметки относятся к № п/п, а не к федеральному номеру.')

    def refresh_summary(self):
        if self.ledger is None:
            return
        all_animals = self.ledger.animals
        all_s = summarize(all_animals)
        bought = summarize([a for a in all_animals if a.number in self.purchased])
        rest = summarize([a for a in all_animals if a.number not in self.purchased])
        for widget, values in ((self.all_card, all_s),
                               (self.bought_card, bought), (self.rest_card, rest)):
            widget.display(values['count'], values['slaughter_weight'])
        self.refresh_home()

    def refresh_home(self):
        if self.ledger is None:
            self.home_title.setText('Начните с загрузки Excel-ведомости.')
            self.home_description.setText('Выберите файл .xlsx, затем перейдите к проверке данных.')
        else:
            data = summarize(self.ledger.animals)
            self.home_title.setText('Ведомость ' + self.ledger.date)
            self.home_description.setText(
                f'{data["count"]} голов · {short_weight(float(data["slaughter_weight"]))} кг. '
                f'Отмечено закупных: {len(self.purchased)}. '
                + ('Разметка не сохранена.' if self.dirty else 'Изменений нет.'))
        self.home_recent.setText('\n'.join(Path(p).name for p in self.recent_paths[:4])
                                 if self.recent_paths else 'Нет недавно открытых файлов.')

    def refresh_report(self):
        if self.ledger is None:
            return
        animals = self.ledger.animals
        all_s = summarize(animals)
        bought = summarize([a for a in animals if a.number in self.purchased])
        rest = summarize([a for a in animals if a.number not in self.purchased])
        for card, group in ((self.report_all, all_s),
                            (self.report_bought, bought), (self.report_rest, rest)):
            card.display(group['count'], group['slaughter_weight'])
        self.report_directions.setText(
            f'Все животные: реализация — {all_s["free_count"]} гол., '
            f'{short_weight(float(all_s["free_weight"]))} кг; '
            f'переработка — {all_s["industrial_count"]} гол., '
            f'{short_weight(float(all_s["industrial_weight"]))} кг.\n'
            f'Закупные: реализация — {bought["free_count"]} гол., '
            f'{short_weight(float(bought["free_weight"]))} кг; '
            f'переработка — {bought["industrial_count"]} гол., '
            f'{short_weight(float(bought["industrial_weight"]))} кг.')
        self.refresh_export_preview()

    def import_markup(self):
        if self.ledger is None:
            return
        directory = self.ledger.path.parent
        path, _ = QFileDialog.getOpenFileName(
            self, 'Открыть прежнюю разметку', str(directory), 'Разметка JSONL (*.jsonl)')
        if not path:
            return
        try:
            incoming = read_purchased_jsonl(self.ledger, Path(path))
            difference = compare_purchased(self.ledger, self.purchased, incoming)
        except (OSError, ValueError, MarkupError) as exc:
            QMessageBox.warning(self, 'Нельзя применить разметку', str(exc))
            return
        if not self.confirm_import(difference, Path(path).name):
            return
        self.purchased.clear()
        self.purchased.update(incoming)
        self.animal_model.refresh_checks()
        self.proxy.refresh_filter()
        self.dirty = True
        self.refresh_summary()
        self.refresh_report()
        self.filtered_info()
        self.status.setText(f'Загружена разметка: {Path(path).name} · сохраните её при необходимости')
        self.autosave_timer.start()

    def confirm_import(self, diff: dict, filename: str) -> bool:
        dialog = QDialog(self)
        dialog.setWindowTitle('Сравнение разметки закупных')
        dialog.resize(570, 390)
        lay = QVBoxLayout(dialog)
        lay.addWidget(make_label('Сверка с ранее сохранённым JSONL', 17, True))
        lay.addWidget(make_label(filename, 12, muted=True))
        text = (
            f'Текущая разметка: {diff["current_count"]} голов · '
            f'{short_weight(float(diff["current_weight"]))} кг\n'
            f'Разметка из файла: {diff["incoming_count"]} голов · '
            f'{short_weight(float(diff["incoming_weight"]))} кг\n'
            f'Изменение веса: {diff["weight_delta"]:+g} кг\n\n'
            f'Будут добавлены закупные №: '
            f'{", ".join(map(str, diff["added"])) or "нет"}\n\n'
            f'Будут сняты отметки с №: '
            f'{", ".join(map(str, diff["removed"])) or "нет"}\n\n'
            'При применении заменятся все текущие галочки. '
            'Исходный XLSX не изменится.'
        )
        details = QPlainTextEdit()
        details.setReadOnly(True)
        details.setPlainText(text)
        lay.addWidget(details, 1)
        box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        box.button(QDialogButtonBox.StandardButton.Ok).setText('Применить разметку')
        box.button(QDialogButtonBox.StandardButton.Cancel).setText('Отмена')
        box.accepted.connect(dialog.accept)
        box.rejected.connect(dialog.reject)
        lay.addWidget(box)
        return dialog.exec() == QDialog.DialogCode.Accepted

    def confirm_empty_markup(self) -> bool:
        if self.purchased:
            return True
        answer = QMessageBox.question(
            self, 'Закупные не отмечены',
            'Отмечено 0 закупных животных. Это действительно подтверждённое отсутствие закупных, '
            'а не ещё не заполненная разметка?',
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No)
        return answer == QMessageBox.StandardButton.Yes

    def save_jsonl(self):
        if self.ledger is None:
            QMessageBox.information(self, 'Нет данных', 'Сначала откройте ведомость.')
            return
        if not self.confirm_empty_markup():
            return
        default = self.saved_markup_path or (
            self.ledger.path.parent /
            f'{self.ledger.date.replace(".", "-")}_{self.ledger.path.stem}_marked.jsonl')
        path, _ = QFileDialog.getSaveFileName(
            self, 'Сохранить разметку', str(default), 'Разметка JSONL (*.jsonl)')
        if not path:
            return
        destination = Path(path)
        if destination.suffix.lower() != '.jsonl':
            destination = destination.with_suffix('.jsonl')
        try:
            write_jsonl(self.ledger, destination, self.purchased)
        except (OSError, ValueError) as exc:
            QMessageBox.critical(self, 'Ошибка сохранения', str(exc))
            return
        self.saved_markup_path = destination
        self.autosave_markup()
        self.dirty = False
        self.autosave_timer.stop()
        self.status.setText(f'Разметка сохранена: {destination.name}')
        self.refresh_home()
        self.refresh_report()
        QMessageBox.information(self, 'Готово',
                                f'Разметка сохранена в:\n{destination}\n\nОригинальный XLSX не изменён.')

    def save_pdf(self):
        if self.ledger is None:
            QMessageBox.information(self, 'Нет данных', 'Сначала откройте ведомость.')
            return
        if not self.confirm_empty_markup():
            return
        default = self.ledger.path.parent / (
            self.ledger.date.replace('.', '-') + '_' + self.ledger.path.stem + '_report.pdf')
        path, _ = QFileDialog.getSaveFileName(
            self, 'Сформировать PDF', str(default), 'PDF (*.pdf)')
        if not path:
            return
        destination = Path(path)
        if destination.suffix.lower() != '.pdf':
            destination = destination.with_suffix('.pdf')
        try:
            export_pdf(self.ledger, self.purchased, destination)
        except (OSError, RuntimeError, ValueError) as exc:
            QMessageBox.critical(self, 'Ошибка экспорта PDF', str(exc))
            return
        QMessageBox.information(self, 'PDF готов', f'Отчёт сохранён:\n{destination}')

    def load_recent(self):
        recent = self.settings.value('recent_paths', [])
        if isinstance(recent, str):
            recent = [recent] if recent else []
        return [str(p) for p in recent if p]

    def record_recent(self, path):
        current = str(Path(path).resolve())
        self.recent_paths = [current] + [p for p in self.recent_paths if p != current]
        self.recent_paths = self.recent_paths[:8]
        self.settings.setValue('recent_paths', self.recent_paths)

    def refresh_history(self):
        files = [p for p in self.recent_paths if Path(p).is_file()]
        self.history_info.setText('Нажмите на файл, чтобы открыть его повторно.'
                                  if files else 'Нет доступных недавно открытых файлов.')
        while self.history_actions.count():
            item = self.history_actions.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()
        for path in files[:5]:
            button = self.button('Открыть: ' + Path(path).name,
                                 lambda checked=False, selected=path: self.load_file(selected))
            button.setToolTip(path)
            self.history_actions.addWidget(button)

    def closeEvent(self, event):
        if self.dirty:
            self.autosave_timer.stop()
            self.autosave_markup()
        if self.confirm_discard():
            event.accept()
        else:
            event.ignore()


def main():
    app = QApplication(sys.argv)
    app.setApplicationName('Reindeer Ledger')
    app.setOrganizationName(COMPANY)
    base_font = QFont('Segoe UI')
    base_font.setPixelSize(13)
    app.setFont(base_font)
    window = Window()
    window.show()
    if len(sys.argv) > 1:
        window.load_file(sys.argv[1])
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
