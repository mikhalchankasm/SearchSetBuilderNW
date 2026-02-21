#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Navisworks SearchSet XML Generator (Qt/PySide6)
Генератор поисковых наборов Navisworks.

Два режима работы:
1. Excel режим - массовое создание простых наборов из Excel (по свойству "Имя")
2. GUI построитель - визуальный редактор сложных наборов с произвольными свойствами
"""

from __future__ import annotations

import os
import sys
import json
import uuid
import copy
import tempfile
import logging
import xml.etree.ElementTree as ET
import xml.dom.minidom
from typing import List, Optional

import defusedxml.ElementTree as DefusedET
from defusedxml import DefusedXmlException

from PySide6 import QtCore, QtGui, QtWidgets

# Excel
import pandas as pd
import win32com.client
import pythoncom

# Пути к файлам
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Настройка логирования
logging.basicConfig(
    level=logging.WARNING,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(BASE_DIR, 'search_set.log'), encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)
SETTINGS_FILE = os.path.join(BASE_DIR, 'search_set_settings.json')
CATEGORIES_FILE = os.path.join(BASE_DIR, 'categories_history.json')
PROJECTS_DIR = os.path.join(BASE_DIR, 'projects')

# Маппинг типов условий
TEST_MAPPING = {
    "=": "equals",
    "Не равно": "not_equals",
    "Содержит": "contains",
    "Подстановочный знак": "wildcard",
    "Определенный": "is_not_empty",
    "Не определено": "is_empty"
}
TEST_MAPPING_REVERSE = {v: k for k, v in TEST_MAPPING.items()}

DATA_TYPES = ["wstring", "int32", "double", "bool", "datetime"]

MIME_SEARCHSET = 'application/x-navis-searchset-json'


# ==================== Утилиты ====================

def load_settings() -> dict:
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            try:
                os.rename(SETTINGS_FILE, SETTINGS_FILE + '.corrupted')
            except OSError:
                pass
    return {}


def save_settings(settings: dict):
    with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


def load_categories_history() -> dict:
    if os.path.exists(CATEGORIES_FILE):
        try:
            with open(CATEGORIES_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError):
            try:
                os.rename(CATEGORIES_FILE, CATEGORIES_FILE + '.corrupted')
            except OSError:
                pass
    return {
        "categories": [
            {"name": "Элемент", "internal": "LcOaNode"},
            {"name": "AVEVA", "internal": "lcldrvm_props"}
        ],
        "properties": {
            "Элемент": [
                {"name": "Имя", "internal": "LcOaSceneBaseUserName"},
                {"name": "Тип", "internal": "LcOaSceneBaseClassName"}
            ],
            "AVEVA": [
                {"name": "Наименование", "internal": "lcldrvm_prop_Наименование"},
                {"name": "Условный_диаметр", "internal": "lcldrvm_prop_Условный_диаметр"}
            ]
        }
    }


def save_categories_history(data: dict):
    with open(CATEGORIES_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def ensure_projects_dir():
    if not os.path.exists(PROJECTS_DIR):
        os.makedirs(PROJECTS_DIR)


# ==================== XML генерация ====================

def make_condition(value: str, test: str) -> ET.Element:
    cond = ET.Element("condition", test=test, flags="0")
    cat = ET.SubElement(cond, "category")
    ET.SubElement(cat, "name", internal="LcOaNode").text = "Элемент"
    prop = ET.SubElement(cond, "property")
    ET.SubElement(prop, "name", internal="LcOaSceneBaseUserName").text = "Имя"
    val = ET.SubElement(cond, "value")
    ET.SubElement(val, "data", type="wstring").text = value
    return cond


def make_condition_advanced(category: str, category_internal: str, property_name: str,
                            property_internal: str, test: str, value: str,
                            data_type: str = "wstring") -> ET.Element:
    cond = ET.Element("condition", test=test, flags="0")
    cat = ET.SubElement(cond, "category")
    ET.SubElement(cat, "name", internal=category_internal).text = category
    prop = ET.SubElement(cond, "property")
    ET.SubElement(prop, "name", internal=property_internal).text = property_name
    val = ET.SubElement(cond, "value")
    ET.SubElement(val, "data", type=data_type).text = value
    return cond


def generate_xml(df: pd.DataFrame, save_path: str, root_folder: str = ""):
    """
    Generate Navisworks XML from Excel data.

    Expected Excel structure:
    - Column 0: Set name (название набора)
    - Column 'Наборы': Folder/group name (optional)
    - Column 'Тип': Condition type per row (=, Содержит, etc.)
    - Other columns: Search values
    """
    exchange = ET.Element("exchange")
    exchange.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
    exchange.set("xsi:noNamespaceSchemaLocation",
                 "http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd")
    exchange.set("units", "m")
    exchange.set("filename", "")
    exchange.set("filepath", "")

    selectionsets = ET.SubElement(exchange, "selectionsets")

    if root_folder:
        root_viewfolder = ET.SubElement(selectionsets, "viewfolder",
                                        name=root_folder, guid=str(uuid.uuid4()))
        parent_element = root_viewfolder
    else:
        parent_element = selectionsets

    # Get set name column (first column)
    set_name_col = df.columns[0] if len(df.columns) > 0 else None

    # Group by folder name, keeping rows without folder assignment at root level
    if 'Наборы' in df.columns:
        grouped = df.groupby('Наборы', dropna=False)
    else:
        grouped = [(None, df)]

    for group_name, group_df in grouped:
        if group_name and not pd.isna(group_name):
            viewfolder = ET.SubElement(parent_element, "viewfolder",
                                       name=str(group_name), guid=str(uuid.uuid4()))
        else:
            viewfolder = parent_element

        for _, row in group_df.iterrows():
            # Use named column for set name instead of positional indexing
            if set_name_col and pd.notna(row[set_name_col]):
                set_name = str(row[set_name_col])
            else:
                set_name = "Набор"

            selectionset = ET.SubElement(viewfolder, "selectionset",
                                         name=set_name, guid=str(uuid.uuid4()))
            findspec = ET.SubElement(selectionset, "findspec", mode="all", disjoint="0")
            conditions = ET.SubElement(findspec, "conditions")

            # Read condition type from each row individually
            test_type_ru = row.get('Тип', 'Содержит') if 'Тип' in df.columns else 'Содержит'
            test = TEST_MAPPING.get(test_type_ru, "contains")

            # Exclude set name column, 'Наборы', and 'Тип' from value columns
            exclude_cols = ['Наборы', 'Тип']
            if set_name_col:
                exclude_cols.append(set_name_col)
            value_cols = [c for c in df.columns if c not in exclude_cols]

            for col in value_cols:
                val = row.get(col)
                if pd.notna(val) and str(val).strip():
                    cond = make_condition(str(val).strip(), test)
                    conditions.append(cond)

            ET.SubElement(findspec, "locator").text = "/"

    xml_str = ET.tostring(exchange, encoding='utf-8')
    pretty_xml = xml.dom.minidom.parseString(xml_str).toprettyxml(indent="  ", encoding="utf-8")
    with open(save_path, 'wb') as f:
        f.write(pretty_xml)


def generate_xml_from_project(project_data: dict, save_path: str):
    exchange = ET.Element("exchange")
    exchange.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
    exchange.set("xsi:noNamespaceSchemaLocation",
                 "http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd")
    exchange.set("units", "m")
    exchange.set("filename", "")
    exchange.set("filepath", "")

    selectionsets = ET.SubElement(exchange, "selectionsets")

    def add_folder(parent_elem, folder_data):
        vf = ET.SubElement(parent_elem, "viewfolder",
                           name=folder_data['name'], guid=str(uuid.uuid4()))
        for subfolder in folder_data.get('folders', []):
            add_folder(vf, subfolder)
        for s in folder_data.get('sets', []):
            add_set(vf, s)

    def add_set(parent_elem, set_data):
        ss = ET.SubElement(parent_elem, "selectionset",
                           name=set_data['name'], guid=str(uuid.uuid4()))
        mode = set_data.get('mode', 'all')
        findspec = ET.SubElement(ss, "findspec", mode=mode, disjoint="0")
        conditions_elem = ET.SubElement(findspec, "conditions")

        for cond in set_data.get('conditions', []):
            cond_elem = make_condition_advanced(
                category=cond.get("category", "Элемент"),
                category_internal=cond.get("category_internal", "LcOaNode"),
                property_name=cond.get("property", "Имя"),
                property_internal=cond.get("property_internal", "LcOaSceneBaseUserName"),
                test=cond.get("test", "contains"),
                value=cond.get("value", ""),
                data_type=cond.get("data_type", "wstring")
            )
            conditions_elem.append(cond_elem)

        ET.SubElement(findspec, "locator").text = "/"

    for folder in project_data.get('folders', []):
        add_folder(selectionsets, folder)
    for s in project_data.get('sets', []):
        add_set(selectionsets, s)

    xml_str = ET.tostring(exchange, encoding='utf-8')
    pretty_xml = xml.dom.minidom.parseString(xml_str).toprettyxml(indent="  ", encoding="utf-8")
    with open(save_path, 'wb') as f:
        f.write(pretty_xml)


def parse_xml_to_project(xml_path: str) -> dict:
    max_size = 10 * 1024 * 1024
    file_size = os.path.getsize(xml_path)
    if file_size > max_size:
        raise ValueError(f"XML файл слишком большой (>{max_size // 1024 // 1024} МБ)")

    with open(xml_path, 'rb') as f:
        raw_content = f.read()

    if raw_content.startswith(b'\xff\xfe'):
        check_encoding = 'utf-16-le'
        raw_for_check = raw_content[2:]
    elif raw_content.startswith(b'\xfe\xff'):
        check_encoding = 'utf-16-be'
        raw_for_check = raw_content[2:]
    elif raw_content.startswith(b'\xef\xbb\xbf'):
        check_encoding = 'utf-8'
        raw_for_check = raw_content[3:]
    else:
        check_encoding = 'utf-8'
        raw_for_check = raw_content

    if check_encoding.startswith('utf-16'):
        try:
            text_for_check = raw_for_check.decode(check_encoding, errors='replace').lower()
        except Exception:
            text_for_check = ""
        if '<!doctype' in text_for_check or '<!entity' in text_for_check:
            raise ValueError("XML содержит DTD/ENTITY — импорт запрещён")
    else:
        raw_lower = raw_for_check.lower()
        if b'<!doctype' in raw_lower or b'<!entity' in raw_lower:
            raise ValueError("XML содержит DTD/ENTITY — импорт запрещён")

    try:
        root = DefusedET.fromstring(raw_content)
    except DefusedXmlException as e:
        raise ValueError(f"XML заблокирован из соображений безопасности: {e}")
    except ET.ParseError as e:
        raise ValueError(f"Ошибка парсинга XML: {e}")

    project = {"name": os.path.splitext(os.path.basename(xml_path))[0], "folders": [], "sets": []}

    def parse_condition(cond_elem) -> dict:
        test = cond_elem.get("test", "contains")
        cat_elem = cond_elem.find("category/name")
        category = cat_elem.text if cat_elem is not None and cat_elem.text else "Элемент"
        category_internal = cat_elem.get("internal", "LcOaNode") if cat_elem is not None else "LcOaNode"
        prop_elem = cond_elem.find("property/name")
        property_name = prop_elem.text if prop_elem is not None and prop_elem.text else "Имя"
        property_internal = prop_elem.get("internal", "LcOaSceneBaseUserName") if prop_elem is not None else "LcOaSceneBaseUserName"
        val_elem = cond_elem.find("value/data")
        value = val_elem.text if val_elem is not None and val_elem.text else ""
        data_type = val_elem.get("type", "wstring") if val_elem is not None else "wstring"
        return {
            "category": category, "category_internal": category_internal,
            "property": property_name, "property_internal": property_internal,
            "test": test, "value": value, "data_type": data_type
        }

    def parse_selectionset(ss_elem) -> dict:
        name = ss_elem.get("name", "Набор")
        findspec = ss_elem.find("findspec")
        mode = findspec.get("mode", "all") if findspec is not None else "all"
        conditions = []
        if findspec is not None:
            conds_elem = findspec.find("conditions")
            if conds_elem is not None:
                for cond_elem in conds_elem.findall("condition"):
                    conditions.append(parse_condition(cond_elem))
        return {"name": name, "mode": mode, "conditions": conditions}

    def parse_viewfolder(vf_elem) -> dict:
        name = vf_elem.get("name", "Папка")
        folder = {"name": name, "folders": [], "sets": []}
        for child in vf_elem:
            if child.tag == "viewfolder":
                folder["folders"].append(parse_viewfolder(child))
            elif child.tag == "selectionset":
                folder["sets"].append(parse_selectionset(child))
        return folder

    selectionsets = root.find("selectionsets")
    if selectionsets is not None:
        for child in selectionsets:
            if child.tag == "viewfolder":
                project["folders"].append(parse_viewfolder(child))
            elif child.tag == "selectionset":
                project["sets"].append(parse_selectionset(child))

    return project


def create_excel_template() -> str:
    template_path = os.path.join(tempfile.gettempdir(), 'search_set_template.xlsx')
    df = pd.DataFrame({
        'Название': ['Пример 1', 'Пример 2'],
        'Наборы': ['Группа1', 'Группа1'],
        'Тип': ['Содержит', 'Содержит'],
        'Значение1': ['Значение1', 'Значение2'],
        'Значение2': ['', '']
    })
    df.to_excel(template_path, index=False)
    return template_path


# ==================== Модельные классы ====================

class SearchSetItem:
    def __init__(self, name: str, is_folder: bool = False, data: Optional[dict] = None):
        self.name = name
        self.is_folder = is_folder
        self.guid = str(uuid.uuid4())
        self.data = data or {}
        self.children: List['SearchSetItem'] = []
        self.parent: Optional['SearchSetItem'] = None

    def add_child(self, child: 'SearchSetItem'):
        child.parent = self
        self.children.append(child)

    def insert_child_before(self, child: 'SearchSetItem', before: 'SearchSetItem'):
        """Insert child before the specified sibling."""
        child.parent = self
        try:
            idx = self.children.index(before)
            self.children.insert(idx, child)
        except ValueError:
            self.children.append(child)

    def remove_child(self, child: 'SearchSetItem'):
        if child in self.children:
            self.children.remove(child)
            child.parent = None

    def to_dict(self) -> dict:
        if self.is_folder:
            return {
                'name': self.name,
                'folders': [c.to_dict() for c in self.children if c.is_folder],
                'sets': [c.to_dict() for c in self.children if not c.is_folder]
            }
        else:
            return {
                'name': self.name,
                'mode': self.data.get('mode', 'all'),
                'conditions': self.data.get('conditions', [])
            }

    @staticmethod
    def from_dict(d: dict, is_folder: bool = False) -> 'SearchSetItem':
        if is_folder or 'folders' in d or 'sets' in d:
            item = SearchSetItem(d.get('name', 'Папка'), is_folder=True)
            for f in d.get('folders', []):
                item.add_child(SearchSetItem.from_dict(f, is_folder=True))
            for s in d.get('sets', []):
                item.add_child(SearchSetItem.from_dict(s, is_folder=False))
            return item
        else:
            item = SearchSetItem(d.get('name', 'Набор'), is_folder=False)
            item.data = {'mode': d.get('mode', 'all'), 'conditions': d.get('conditions', [])}
            return item

    def is_ancestor_of(self, node: 'SearchSetItem') -> bool:
        cur = node.parent
        while cur is not None:
            if cur is self:
                return True
            cur = cur.parent
        return False


# ==================== Кастомные виджеты ====================

class StructureTree(QtWidgets.QTreeWidget):
    itemMoved = QtCore.Signal(object, object)  # (dragged_item, target_item)
    deleteRequested = QtCore.Signal(list)      # [SearchSetItem, ...]
    contextCollapseRequested = QtCore.Signal()
    contextRenameRequested = QtCore.Signal()
    contextDeleteRequested = QtCore.Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHeaderHidden(True)
        # Разрешаем множественный выбор (Shift/Ctrl)
        self.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)
        self.setAnimated(True)
        self.setIndentation(32)  # Увеличенный отступ для лучшей визуализации иерархии

        self._dragged_item: Optional[SearchSetItem] = None

        self.setStyleSheet("""
            QTreeWidget {
                font-size: 10pt;
                border: 1px solid #ccc;
                border-radius: 4px;
                background-color: #ffffff;
            }
            QTreeWidget::item {
                padding: 4px 4px;
                min-height: 22px;
            }
            QTreeWidget::item:selected {
                background-color: #0078d4;
                color: white;
            }
            QTreeWidget::item:hover:!selected {
                background-color: #e5f3ff;
            }
        """)

    def selected_model_item(self) -> Optional[SearchSetItem]:
        items = self.selectedItems()
        if items:
            return items[0].data(0, QtCore.Qt.UserRole)
        return None

    def selected_model_items(self) -> List[SearchSetItem]:
        """Вернуть все выделенные модельные элементы."""
        result: List[SearchSetItem] = []
        for qitem in self.selectedItems():
            model = qitem.data(0, QtCore.Qt.UserRole)
            if isinstance(model, SearchSetItem):
                result.append(model)
        return result

    def keyPressEvent(self, event: QtGui.QKeyEvent):
        """Обработка Delete для массового удаления без изменения развёрнутости."""
        if event.key() == QtCore.Qt.Key_Delete:
            items = self.selected_model_items()
            if items:
                self.deleteRequested.emit(items)
                return
        super().keyPressEvent(event)

    def contextMenuEvent(self, event: QtGui.QContextMenuEvent):
        """Контекстное меню на элементах дерева."""
        pos = event.pos()
        qitem = self.itemAt(pos)
        if not qitem:
            return

        # Делаем элемент текущим, чтобы MainWindow работал с ним
        self.setCurrentItem(qitem)
        model: SearchSetItem = qitem.data(0, QtCore.Qt.UserRole)

        menu = QtWidgets.QMenu(self)
        collapse_action = None
        if model and model.is_folder:
            collapse_action = menu.addAction("Свернуть все")
            menu.addSeparator()

        rename_action = menu.addAction("Переименовать")
        delete_action = menu.addAction("Удалить")

        chosen = menu.exec(self.mapToGlobal(pos))
        if not chosen:
            return

        if collapse_action and chosen is collapse_action:
            self.contextCollapseRequested.emit()
        elif chosen is rename_action:
            self.contextRenameRequested.emit()
        elif chosen is delete_action:
            self.contextDeleteRequested.emit()

    def startDrag(self, supportedActions):
        item = self.currentItem()
        if not item:
            return
        model: SearchSetItem = item.data(0, QtCore.Qt.UserRole)
        if not model:
            return
        self._dragged_item = model
        mime = QtCore.QMimeData()
        mime.setData(MIME_SEARCHSET, QtCore.QByteArray(model.guid.encode('utf-8')))
        drag = QtGui.QDrag(self)
        drag.setMimeData(mime)
        drag.exec(QtCore.Qt.MoveAction)

    def dragEnterEvent(self, e):
        if e.mimeData().hasFormat(MIME_SEARCHSET):
            e.acceptProposedAction()
        else:
            e.ignore()

    def dragMoveEvent(self, e):
        if e.mimeData().hasFormat(MIME_SEARCHSET):
            target_item = self.itemAt(e.position().toPoint())
            if target_item:
                target_model = target_item.data(0, QtCore.Qt.UserRole)
                # Prevent dropping on self or descendants
                if self._dragged_item and (
                    target_model is self._dragged_item or
                    self._dragged_item.is_ancestor_of(target_model)
                ):
                    e.ignore()
                    return
            e.acceptProposedAction()
        else:
            e.ignore()

    def dropEvent(self, e):
        if not e.mimeData().hasFormat(MIME_SEARCHSET):
            e.ignore()
            return
        if not self._dragged_item:
            e.ignore()
            return

        target_qitem = self.itemAt(e.position().toPoint())
        target_model = None
        if target_qitem:
            target_model = target_qitem.data(0, QtCore.Qt.UserRole)

        # Prevent dropping on self or descendants
        if target_model and (
            target_model is self._dragged_item or
            self._dragged_item.is_ancestor_of(target_model)
        ):
            e.ignore()
            self._dragged_item = None
            return

        e.acceptProposedAction()
        self.itemMoved.emit(self._dragged_item, target_model)
        self._dragged_item = None


class ConditionWidget(QtWidgets.QFrame):
    deleted = QtCore.Signal(object)
    moveUp = QtCore.Signal(object)
    moveDown = QtCore.Signal(object)

    def __init__(self, index: int, categories_history: dict,
                 condition_data: Optional[dict] = None, parent=None):
        super().__init__(parent)
        self.index = index
        self.categories_history = categories_history

        self.setFrameStyle(QtWidgets.QFrame.StyledPanel)
        self.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        self.setStyleSheet("""
            ConditionWidget {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 8px;
            }
        """)

        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(16)

        form_layout = QtWidgets.QGridLayout()
        form_layout.setSpacing(10)
        form_layout.setColumnMinimumWidth(0, 100)

        # Категория
        form_layout.addWidget(QtWidgets.QLabel("Категория:"), 0, 0, QtCore.Qt.AlignRight)
        self.category_combo = QtWidgets.QComboBox()
        self.category_combo.setEditable(True)
        self.category_combo.setMinimumWidth(200)
        cat_names = [c["name"] for c in categories_history.get("categories", [])]
        self.category_combo.addItems(cat_names)
        self.category_combo.currentTextChanged.connect(self._on_category_changed)
        form_layout.addWidget(self.category_combo, 0, 1)

        # Свойство
        form_layout.addWidget(QtWidgets.QLabel("Свойство:"), 1, 0, QtCore.Qt.AlignRight)
        self.property_combo = QtWidgets.QComboBox()
        self.property_combo.setEditable(True)
        self.property_combo.setMinimumWidth(200)
        form_layout.addWidget(self.property_combo, 1, 1)

        # Условие
        form_layout.addWidget(QtWidgets.QLabel("Условие:"), 2, 0, QtCore.Qt.AlignRight)
        self.test_combo = QtWidgets.QComboBox()
        self.test_combo.addItems(list(TEST_MAPPING.keys()))
        self.test_combo.setCurrentText("Содержит")
        form_layout.addWidget(self.test_combo, 2, 1)

        # Значение
        form_layout.addWidget(QtWidgets.QLabel("Значение:"), 3, 0, QtCore.Qt.AlignRight)
        self.value_edit = QtWidgets.QLineEdit()
        self.value_edit.setMinimumWidth(200)
        form_layout.addWidget(self.value_edit, 3, 1)

        # Тип данных
        form_layout.addWidget(QtWidgets.QLabel("Тип:"), 4, 0, QtCore.Qt.AlignRight)
        self.datatype_combo = QtWidgets.QComboBox()
        self.datatype_combo.addItems(DATA_TYPES)
        form_layout.addWidget(self.datatype_combo, 4, 1)

        layout.addLayout(form_layout, 1)

        # Кнопки
        btn_layout = QtWidgets.QVBoxLayout()
        btn_layout.setSpacing(6)

        btn_style = """
            QPushButton {
                font-size: 14pt;
                min-width: 36px;
                min-height: 36px;
                border: 1px solid #ccc;
                border-radius: 4px;
                background-color: #f0f0f0;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
        """

        btn_up = QtWidgets.QPushButton("▲")
        btn_up.setStyleSheet(btn_style)
        btn_up.clicked.connect(lambda: self.moveUp.emit(self))
        btn_layout.addWidget(btn_up)

        btn_down = QtWidgets.QPushButton("▼")
        btn_down.setStyleSheet(btn_style)
        btn_down.clicked.connect(lambda: self.moveDown.emit(self))
        btn_layout.addWidget(btn_down)

        btn_del = QtWidgets.QPushButton("✕")
        btn_del.setStyleSheet(btn_style + "QPushButton { color: #dc3545; font-weight: bold; }")
        btn_del.clicked.connect(lambda: self.deleted.emit(self))
        btn_layout.addWidget(btn_del)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        if condition_data:
            self.load_data(condition_data)
        else:
            self._on_category_changed(self.category_combo.currentText())

    def _on_category_changed(self, cat_name: str):
        self.property_combo.clear()
        props = self.categories_history.get("properties", {}).get(cat_name, [])
        prop_names = [p["name"] for p in props]
        self.property_combo.addItems(prop_names)

    def load_data(self, data: dict):
        self.category_combo.setCurrentText(data.get("category", "Элемент"))
        self._on_category_changed(self.category_combo.currentText())
        self.property_combo.setCurrentText(data.get("property", "Имя"))
        test_ru = TEST_MAPPING_REVERSE.get(data.get("test", "contains"), "Содержит")
        self.test_combo.setCurrentText(test_ru)
        self.value_edit.setText(data.get("value", ""))
        self.datatype_combo.setCurrentText(data.get("data_type", "wstring"))

    def get_data(self) -> dict:
        cat_name = self.category_combo.currentText()
        prop_name = self.property_combo.currentText()

        cat_internal = cat_name
        for c in self.categories_history.get("categories", []):
            if c["name"] == cat_name:
                cat_internal = c.get("internal", cat_name)
                break

        prop_internal = prop_name
        props = self.categories_history.get("properties", {}).get(cat_name, [])
        for p in props:
            if p["name"] == prop_name:
                prop_internal = p.get("internal", prop_name)
                break

        return {
            "category": cat_name, "category_internal": cat_internal,
            "property": prop_name, "property_internal": prop_internal,
            "test": TEST_MAPPING.get(self.test_combo.currentText(), "contains"),
            "value": self.value_edit.text(),
            "data_type": self.datatype_combo.currentText()
        }

    def set_index(self, idx: int):
        self.index = idx


# ==================== Главное окно ====================

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SearchSet Builder")
        self.resize(1400, 900)
        self.setMinimumSize(1000, 600)

        self.settings = load_settings()
        self.categories_history = load_categories_history()
        self.root_item = SearchSetItem("Корень", is_folder=True)
        self.current_item: Optional[SearchSetItem] = None
        self.condition_widgets: List[ConditionWidget] = []

        self.excel = None
        self.workbook = None
        self.worksheet = None
        self._com_initialized = False
        self._excel_created_by_us = False

        # Текущий файл проекта для автосохранения
        self.current_project_path: Optional[str] = None

        self._build_ui()
        self._connect_signals()
        self._refresh_tree()
        self._set_editor_enabled(False)
        self._update_title()

    def _build_ui(self):
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        main_layout = QtWidgets.QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)

        self.tabs = QtWidgets.QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: none;
                background-color: #f5f5f5;
            }
            QTabBar::tab {
                padding: 12px 32px;
                font-size: 12pt;
                font-weight: bold;
                background-color: #e0e0e0;
                border: none;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background-color: #0078d4;
                color: white;
            }
            QTabBar::tab:hover:!selected {
                background-color: #d0d0d0;
            }
        """)
        main_layout.addWidget(self.tabs)

        self.excel_tab = self._create_excel_tab()
        self.tabs.addTab(self.excel_tab, "📊 Excel режим")

        self.gui_tab = self._create_gui_tab()
        self.tabs.addTab(self.gui_tab, "🔧 GUI построитель")

    def _create_excel_tab(self) -> QtWidgets.QWidget:
        tab = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(tab)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(20)

        source_group = QtWidgets.QGroupBox("Источник данных Excel")
        source_group.setStyleSheet("QGroupBox { font-size: 11pt; font-weight: bold; }")
        source_layout = QtWidgets.QVBoxLayout(source_group)
        source_layout.setSpacing(12)

        self.use_xlwings_check = QtWidgets.QCheckBox("Использовать активный Excel")
        self.use_xlwings_check.setChecked(self.settings.get('use_xlwings', True))
        self.use_xlwings_check.toggled.connect(self._toggle_excel_source)
        source_layout.addWidget(self.use_xlwings_check)

        file_layout = QtWidgets.QHBoxLayout()
        self.excel_path_edit = QtWidgets.QLineEdit()
        self.excel_path_edit.setPlaceholderText("Путь к Excel файлу...")
        self.excel_path_edit.setText(self.settings.get('excel_path', ''))
        file_layout.addWidget(self.excel_path_edit, 1)
        self.browse_btn = QtWidgets.QPushButton("Обзор...")
        self.browse_btn.clicked.connect(self._browse_excel)
        file_layout.addWidget(self.browse_btn)
        source_layout.addLayout(file_layout)

        self.excel_status = QtWidgets.QLabel("Не подключен")
        self.excel_status.setStyleSheet("color: #6c757d; font-style: italic; padding: 4px;")
        source_layout.addWidget(self.excel_status)

        btn_layout = QtWidgets.QHBoxLayout()
        self.init_excel_btn = QtWidgets.QPushButton("Подключиться к Excel")
        self.init_excel_btn.clicked.connect(self._initialize_excel)
        btn_layout.addWidget(self.init_excel_btn)
        self.refresh_excel_btn = QtWidgets.QPushButton("Обновить")
        self.refresh_excel_btn.clicked.connect(self._refresh_excel)
        btn_layout.addWidget(self.refresh_excel_btn)
        self.template_btn = QtWidgets.QPushButton("Создать шаблон")
        self.template_btn.clicked.connect(self._create_template)
        btn_layout.addWidget(self.template_btn)
        btn_layout.addStretch()
        source_layout.addLayout(btn_layout)

        layout.addWidget(source_group)

        gen_group = QtWidgets.QGroupBox("Настройки генерации")
        gen_group.setStyleSheet("QGroupBox { font-size: 11pt; font-weight: bold; }")
        gen_layout = QtWidgets.QFormLayout(gen_group)
        gen_layout.setSpacing(12)

        self.save_path_edit = QtWidgets.QLineEdit()
        self.save_path_edit.setText(self.settings.get('save_path', ''))
        save_layout = QtWidgets.QHBoxLayout()
        save_layout.addWidget(self.save_path_edit, 1)
        save_browse = QtWidgets.QPushButton("Обзор...")
        save_browse.clicked.connect(self._browse_save_path)
        save_layout.addWidget(save_browse)
        gen_layout.addRow("Путь сохранения XML:", save_layout)

        layout.addWidget(gen_group)

        self.generate_btn = QtWidgets.QPushButton("Сгенерировать XML")
        self.generate_btn.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                font-size: 14pt;
                font-weight: bold;
                padding: 16px 32px;
                border-radius: 8px;
                border: none;
            }
            QPushButton:hover { background-color: #218838; }
            QPushButton:pressed { background-color: #1e7e34; }
        """)
        self.generate_btn.clicked.connect(self._generate_excel_xml)
        layout.addWidget(self.generate_btn)

        layout.addStretch()
        self._toggle_excel_source(self.use_xlwings_check.isChecked())
        return tab

    def _create_gui_tab(self) -> QtWidgets.QWidget:
        tab = QtWidgets.QWidget()
        layout = QtWidgets.QHBoxLayout(tab)
        layout.setContentsMargins(0, 0, 0, 0)

        splitter = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        layout.addWidget(splitter)

        # Левая панель
        left_widget = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(left_widget)
        left_layout.setContentsMargins(16, 16, 8, 16)
        left_layout.setSpacing(12)

        left_group = QtWidgets.QGroupBox("Структура наборов")
        left_group.setStyleSheet("QGroupBox { font-size: 11pt; font-weight: bold; }")
        left_group_layout = QtWidgets.QVBoxLayout(left_group)

        search_layout = QtWidgets.QHBoxLayout()
        self.tree_search_edit = QtWidgets.QLineEdit()
        self.tree_search_edit.setPlaceholderText("Фильтр по имени...")
        self.tree_search_edit.setClearButtonEnabled(True)
        self.tree_search_edit.textChanged.connect(self._filter_tree)
        search_layout.addWidget(self.tree_search_edit)
        left_group_layout.addLayout(search_layout)

        self.tree = StructureTree()
        left_group_layout.addWidget(self.tree)

        tree_btn_layout = QtWidgets.QHBoxLayout()
        tree_btn_layout.setSpacing(8)

        btn_style = """
            QPushButton {
                padding: 8px 16px;
                font-size: 10pt;
                border: 1px solid #ccc;
                border-radius: 4px;
                background-color: #f8f9fa;
            }
            QPushButton:hover { background-color: #e9ecef; }
        """

        self.add_folder_btn = QtWidgets.QPushButton("+ Папка")
        self.add_folder_btn.setStyleSheet(btn_style)
        self.add_folder_btn.clicked.connect(self._add_folder)
        tree_btn_layout.addWidget(self.add_folder_btn)

        self.add_set_btn = QtWidgets.QPushButton("+ Набор")
        self.add_set_btn.setStyleSheet(btn_style)
        self.add_set_btn.clicked.connect(self._add_set)
        tree_btn_layout.addWidget(self.add_set_btn)

        self.copy_btn = QtWidgets.QPushButton("Копировать")
        self.copy_btn.setStyleSheet(btn_style)
        self.copy_btn.clicked.connect(self._copy_item)
        tree_btn_layout.addWidget(self.copy_btn)

        self.rename_btn = QtWidgets.QPushButton("Переименовать")
        self.rename_btn.setStyleSheet(btn_style)
        self.rename_btn.clicked.connect(self._rename_item)
        tree_btn_layout.addWidget(self.rename_btn)

        self.delete_btn = QtWidgets.QPushButton("Удалить")
        self.delete_btn.setStyleSheet(btn_style + "QPushButton { color: #dc3545; }")
        self.delete_btn.clicked.connect(self._delete_item)
        tree_btn_layout.addWidget(self.delete_btn)

        left_group_layout.addLayout(tree_btn_layout)
        left_layout.addWidget(left_group)

        # Кнопки проекта
        project_btn_layout = QtWidgets.QHBoxLayout()
        project_btn_layout.setSpacing(8)

        self.new_project_btn = QtWidgets.QPushButton("Новый")
        self.new_project_btn.setStyleSheet(btn_style)
        self.new_project_btn.clicked.connect(self._new_project)
        project_btn_layout.addWidget(self.new_project_btn)

        self.load_json_btn = QtWidgets.QPushButton("Открыть проект")
        self.load_json_btn.setStyleSheet(btn_style)
        self.load_json_btn.clicked.connect(self._load_json)
        project_btn_layout.addWidget(self.load_json_btn)

        self.save_json_btn = QtWidgets.QPushButton("Сохранить проект")
        self.save_json_btn.setStyleSheet(btn_style)
        self.save_json_btn.clicked.connect(self._save_json)
        project_btn_layout.addWidget(self.save_json_btn)

        self.save_as_btn = QtWidgets.QPushButton("Сохранить как...")
        self.save_as_btn.setStyleSheet(btn_style)
        self.save_as_btn.clicked.connect(self._save_json_as)
        project_btn_layout.addWidget(self.save_as_btn)

        left_layout.addLayout(project_btn_layout)

        # Метка с путём к текущему файлу проекта
        self.project_path_label = QtWidgets.QLabel()
        self.project_path_label.setStyleSheet("""
            QLabel {
                color: #666;
                font-size: 9pt;
                padding: 2px 4px;
            }
        """)
        self.project_path_label.setWordWrap(True)
        left_layout.addWidget(self.project_path_label)

        import_export_layout = QtWidgets.QHBoxLayout()
        import_export_layout.setSpacing(8)

        self.import_xml_btn = QtWidgets.QPushButton("Импорт из Navisworks")
        self.import_xml_btn.setStyleSheet(btn_style)
        self.import_xml_btn.clicked.connect(self._import_xml)
        import_export_layout.addWidget(self.import_xml_btn)

        self.export_xml_btn = QtWidgets.QPushButton("Экспорт в Navisworks")
        self.export_xml_btn.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                font-weight: bold;
                padding: 8px 16px;
                font-size: 10pt;
                border-radius: 4px;
                border: none;
            }
            QPushButton:hover { background-color: #218838; }
        """)
        self.export_xml_btn.clicked.connect(self._export_xml)
        import_export_layout.addWidget(self.export_xml_btn)

        left_layout.addLayout(import_export_layout)
        splitter.addWidget(left_widget)

        # Правая панель
        right_widget = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(right_widget)
        right_layout.setContentsMargins(8, 16, 16, 16)

        right_group = QtWidgets.QGroupBox("Редактор условий")
        right_group.setStyleSheet("QGroupBox { font-size: 11pt; font-weight: bold; }")
        right_group_layout = QtWidgets.QVBoxLayout(right_group)
        right_group_layout.setSpacing(12)

        self.editor_hint = QtWidgets.QLabel("Выберите набор слева, чтобы редактировать условия.")
        self.editor_hint.setWordWrap(True)
        self.editor_hint.setAlignment(QtCore.Qt.AlignCenter)
        self.editor_hint.setStyleSheet("color: #6c757d; font-style: italic; padding: 6px;")
        right_group_layout.addWidget(self.editor_hint)

        name_layout = QtWidgets.QHBoxLayout()
        name_layout.addWidget(QtWidgets.QLabel("Имя набора:"))
        self.name_edit = QtWidgets.QLineEdit()
        self.name_edit.setPlaceholderText("Введите имя набора...")
        self.name_edit.setStyleSheet("padding: 8px; font-size: 11pt;")
        self.name_edit.editingFinished.connect(self._on_name_changed)
        name_layout.addWidget(self.name_edit, 1)
        right_group_layout.addLayout(name_layout)

        logic_layout = QtWidgets.QHBoxLayout()
        logic_layout.addWidget(QtWidgets.QLabel("Логика:"))
        self.mode_group = QtWidgets.QButtonGroup()
        self.mode_all = QtWidgets.QRadioButton("AND (все условия)")
        self.mode_all.setChecked(True)
        self.mode_any = QtWidgets.QRadioButton("OR (любое условие)")
        self.mode_group.addButton(self.mode_all, 0)
        self.mode_group.addButton(self.mode_any, 1)
        logic_layout.addWidget(self.mode_all)
        logic_layout.addWidget(self.mode_any)
        logic_layout.addStretch()
        right_group_layout.addLayout(logic_layout)

        scroll = QtWidgets.QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet("QScrollArea { border: none; background-color: transparent; }")

        self.conditions_widget = QtWidgets.QWidget()
        self.conditions_layout = QtWidgets.QVBoxLayout(self.conditions_widget)
        self.conditions_layout.setAlignment(QtCore.Qt.AlignTop)
        self.conditions_layout.setSpacing(12)
        self.conditions_layout.setContentsMargins(0, 0, 0, 0)
        scroll.setWidget(self.conditions_widget)

        right_group_layout.addWidget(scroll, 1)

        self.add_condition_btn = QtWidgets.QPushButton("+ Добавить условие")
        self.add_condition_btn.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                font-size: 11pt;
                border: 2px dashed #0078d4;
                border-radius: 6px;
                background-color: #f0f7ff;
                color: #0078d4;
            }
            QPushButton:hover {
                background-color: #e0efff;
            }
        """)
        self.add_condition_btn.clicked.connect(lambda: self._add_condition())
        right_group_layout.addWidget(self.add_condition_btn)

        right_layout.addWidget(right_group)
        splitter.addWidget(right_widget)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        splitter.setSizes([400, 800])

        return tab

    def _connect_signals(self):
        self.tree.itemSelectionChanged.connect(self._on_tree_select)
        self.tree.itemMoved.connect(self._on_item_moved)
        self.tree.deleteRequested.connect(self._on_tree_delete_requested)
        self.tree.contextCollapseRequested.connect(self._on_tree_context_collapse)
        self.tree.contextRenameRequested.connect(self._rename_item)
        self.tree.contextDeleteRequested.connect(self._delete_item)

    # ==================== Excel режим ====================

    def _toggle_excel_source(self, use_xlwings: bool):
        self.excel_path_edit.setEnabled(not use_xlwings)
        self.browse_btn.setEnabled(not use_xlwings)
        self.init_excel_btn.setEnabled(use_xlwings)
        self.refresh_excel_btn.setEnabled(use_xlwings)

    def _browse_excel(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Выберите Excel файл", "", "Excel Files (*.xlsx)")
        if path:
            self.excel_path_edit.setText(path)

    def _browse_save_path(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Сохранить XML", "", "XML Files (*.xml)")
        if path:
            self.save_path_edit.setText(path)

    def _initialize_excel(self):
        try:
            if not self._com_initialized:
                pythoncom.CoInitialize()
                self._com_initialized = True
            self.excel_status.setText("Подключение...")
            try:
                self.excel = win32com.client.GetActiveObject("Excel.Application")
                self._excel_created_by_us = False
                self.excel_status.setText("✓ Подключен к Excel")
            except Exception:
                self.excel = win32com.client.Dispatch("Excel.Application")
                self._excel_created_by_us = True
                self.excel_status.setText("✓ Создан новый Excel")

            if self.excel.Workbooks.Count > 0:
                self.workbook = self.excel.ActiveWorkbook
                self.worksheet = self.workbook.ActiveSheet
                self.excel_status.setText(f"✓ {self.workbook.Name} / {self.worksheet.Name}")
                self.excel_status.setStyleSheet("color: #28a745; font-weight: bold;")
            else:
                self.workbook = None
                self.worksheet = None
                self.excel_status.setText("⚠ Нет открытых книг")
                self.excel_status.setStyleSheet("color: #ffc107;")
        except Exception as e:
            self.workbook = None
            self.worksheet = None
            self.excel_status.setText(f"✕ Ошибка: {e}")
            self.excel_status.setStyleSheet("color: #dc3545;")

    def _refresh_excel(self):
        if not self.excel:
            self._initialize_excel()
            return
        try:
            if self.excel.Workbooks.Count > 0:
                self.workbook = self.excel.ActiveWorkbook
                self.worksheet = self.workbook.ActiveSheet
                self.excel_status.setText(f"✓ {self.workbook.Name} / {self.worksheet.Name}")
                self.excel_status.setStyleSheet("color: #28a745; font-weight: bold;")
            else:
                # No workbooks open - clear stale references
                self.workbook = None
                self.worksheet = None
                self.excel_status.setText("⚠ Нет открытых книг")
                self.excel_status.setStyleSheet("color: #ffc107;")
        except Exception as e:
            self.workbook = None
            self.worksheet = None
            self.excel_status.setText(f"✕ Ошибка: {e}")
            self.excel_status.setStyleSheet("color: #dc3545;")

    def _create_template(self):
        try:
            path = create_excel_template()
            os.startfile(path)
            QtWidgets.QMessageBox.information(self, "Готово", f"Шаблон создан:\n{path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Ошибка", str(e))

    def _generate_excel_xml(self):
        save_path = self.save_path_edit.text()
        if not save_path:
            QtWidgets.QMessageBox.warning(self, "Ошибка", "Укажите путь сохранения XML")
            return

        try:
            if self.use_xlwings_check.isChecked():
                if not self.worksheet:
                    QtWidgets.QMessageBox.warning(self, "Ошибка", "Подключитесь к Excel")
                    return
                used_range = self.worksheet.UsedRange
                data = used_range.Value
                if not data:
                    QtWidgets.QMessageBox.warning(self, "Ошибка", "Нет данных в листе")
                    return
                # Handle scalar (single cell) or single row
                if not isinstance(data, (list, tuple)):
                    QtWidgets.QMessageBox.warning(self, "Ошибка",
                        "Данные должны содержать заголовок и минимум одну строку")
                    return
                if len(data) < 2:
                    QtWidgets.QMessageBox.warning(self, "Ошибка",
                        "Данные должны содержать заголовок и минимум одну строку данных")
                    return
                # Check first row is also a tuple/list with multiple columns
                if not isinstance(data[0], (list, tuple)) or len(data[0]) < 2:
                    QtWidgets.QMessageBox.warning(self, "Ошибка",
                        "Требуется минимум 2 столбца (название набора + значения)")
                    return
                df = pd.DataFrame(data[1:], columns=data[0])
            else:
                excel_path = self.excel_path_edit.text()
                if not excel_path or not os.path.exists(excel_path):
                    QtWidgets.QMessageBox.warning(self, "Ошибка", "Укажите Excel файл")
                    return
                if excel_path.lower().endswith(".xls"):
                    QtWidgets.QMessageBox.warning(
                        self, "Ошибка",
                        "Формат .xls не поддерживается. Сохраните файл как .xlsx."
                    )
                    return
                df = pd.read_excel(excel_path)

            if df.empty:
                QtWidgets.QMessageBox.warning(self, "Ошибка", "Таблица не содержит данных")
                return

            generate_xml(df, save_path)

            self.settings['use_xlwings'] = self.use_xlwings_check.isChecked()
            self.settings['excel_path'] = self.excel_path_edit.text()
            self.settings['save_path'] = save_path
            save_settings(self.settings)

            QtWidgets.QMessageBox.information(self, "Готово", f"XML сохранен:\n{save_path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Ошибка", str(e))

    # ==================== GUI построитель ====================

    def _get_expanded_guids(self) -> set:
        """Собрать GUID всех развёрнутых элементов дерева."""
        expanded: set[str] = set()

        def walk(qitem: QtWidgets.QTreeWidgetItem):
            model = qitem.data(0, QtCore.Qt.UserRole)
            if qitem.isExpanded() and isinstance(model, SearchSetItem):
                expanded.add(model.guid)
            for i in range(qitem.childCount()):
                walk(qitem.child(i))

        for i in range(self.tree.topLevelItemCount()):
            walk(self.tree.topLevelItem(i))
        return expanded

    def _refresh_tree(self, preserve_state: bool = True):
        """Перестроить дерево, опционально сохранив развёрнутость и выделение."""
        expanded_guids: set[str] = set()
        selected_guid: Optional[str] = None

        if preserve_state and self.tree.topLevelItemCount() > 0:
            expanded_guids = self._get_expanded_guids()
            sel_item = self.tree.selected_model_item()
            if sel_item:
                selected_guid = sel_item.guid

        self.tree.clear()
        self.current_item = None

        def add_node(parent_qitem, node: SearchSetItem):
            icon = "📁" if node.is_folder else "📄"
            qitem = QtWidgets.QTreeWidgetItem([f"{icon} {node.name}"])
            qitem.setData(0, QtCore.Qt.UserRole, node)
            if parent_qitem is None:
                self.tree.addTopLevelItem(qitem)
            else:
                parent_qitem.addChild(qitem)
            # Восстанавливаем развёрнутость, если есть сохранённое состояние
            if not preserve_state or not expanded_guids:
                qitem.setExpanded(True)
            else:
                qitem.setExpanded(node.guid in expanded_guids)

            # Восстанавливаем текущее выделение
            if selected_guid and node.guid == selected_guid:
                self.tree.setCurrentItem(qitem)

            for c in node.children:
                add_node(qitem, c)

        for c in self.root_item.children:
            add_node(None, c)

        # Re-apply filter after rebuild, if any
        if hasattr(self, "tree_search_edit"):
            search_text = self.tree_search_edit.text()
            if search_text:
                self._filter_tree(search_text)

    def _filter_tree(self, text: str):
        needle = text.strip().lower()
        if not needle:
            # Unhide all items instead of calling _refresh_tree to avoid
            # infinite recursion: _refresh_tree re-checks text() which
            # may still be non-empty whitespace and call _filter_tree again.
            def unhide_all(qitem: QtWidgets.QTreeWidgetItem):
                qitem.setHidden(False)
                for i in range(qitem.childCount()):
                    unhide_all(qitem.child(i))

            for i in range(self.tree.topLevelItemCount()):
                unhide_all(self.tree.topLevelItem(i))
            return

        def match_item(qitem: QtWidgets.QTreeWidgetItem) -> bool:
            model = qitem.data(0, QtCore.Qt.UserRole)
            name = model.name if isinstance(model, SearchSetItem) else qitem.text(0)
            name = name.lower()
            matched = needle in name
            child_match = False
            for i in range(qitem.childCount()):
                if match_item(qitem.child(i)):
                    child_match = True
            visible = matched or child_match
            qitem.setHidden(not visible)
            if child_match:
                qitem.setExpanded(True)
            return visible

        for i in range(self.tree.topLevelItemCount()):
            match_item(self.tree.topLevelItem(i))

    def _update_title(self):
        """Update window title with current project file name."""
        base_title = "SearchSet Builder"
        if self.current_project_path:
            file_name = os.path.basename(self.current_project_path)
            self.setWindowTitle(f"{file_name} [автосохранение] — {base_title}")
        else:
            self.setWindowTitle(f"Новый проект (не сохранён!) — {base_title}")
        self._update_project_path_label()

    def _update_project_path_label(self):
        """Update the project path label."""
        if self.current_project_path:
            self.project_path_label.setText(f"📁 {self.current_project_path}")
        else:
            self.project_path_label.setText("📁 Проект не сохранён")

    def _auto_save(self):
        """Auto-save current project to file if path is set.

        Uses atomic write (temp file + rename) to prevent corruption.
        Logs errors instead of silently ignoring them.
        """
        if not self.current_project_path:
            return

        try:
            self._save_current_item()
            data = {
                'name': os.path.splitext(os.path.basename(self.current_project_path))[0],
                'folders': [c.to_dict() for c in self.root_item.children if c.is_folder],
                'sets': [c.to_dict() for c in self.root_item.children if not c.is_folder]
            }

            # Атомарная запись: сначала записываем во временный файл
            temp_fd, temp_path = tempfile.mkstemp(
                dir=os.path.dirname(self.current_project_path),
                prefix='.tmp_',
                suffix='.json',
                text=True
            )

            try:
                with os.fdopen(temp_fd, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)

                # Заменяем целевой файл только если запись прошла успешно
                if os.path.exists(self.current_project_path):
                    os.replace(temp_path, self.current_project_path)
                else:
                    os.rename(temp_path, self.current_project_path)
            except Exception:
                # Удаляем временный файл в случае ошибки
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                raise

        except Exception as e:
            logger.error(f"Ошибка автосохранения проекта {self.current_project_path}: {e}",
                        exc_info=True)
            # Показываем предупреждение пользователю только при критической ошибке
            QtWidgets.QMessageBox.warning(
                self,
                "Ошибка автосохранения",
                f"Не удалось автоматически сохранить проект:\n{str(e)}\n\n"
                f"Рекомендуется вручную сохранить проект через меню 'Сохранить как'."
            )

    def _on_tree_select(self):
        self._save_current_item()
        item = self.tree.selected_model_item()
        if item and not item.is_folder:
            self.current_item = item
            self._load_set_to_editor(item)
        else:
            self.current_item = None
            self._clear_editor()

    def _on_item_moved(self, dragged: SearchSetItem, target: Optional[SearchSetItem]):
        """Move dragged item to target position."""
        if not dragged:
            return

        # Remove from current parent
        if dragged.parent:
            dragged.parent.remove_child(dragged)
        else:
            self.root_item.remove_child(dragged)

        # Add to new position
        if target is None:
            # Dropped on empty space -> move to end of root
            self.root_item.add_child(dragged)
        elif target.is_folder:
            # Dropped on folder -> move INTO folder (at end)
            target.add_child(dragged)
        else:
            # Dropped on set/item -> insert BEFORE that item
            if target.parent:
                target.parent.insert_child_before(dragged, target)
            else:
                self.root_item.insert_child_before(dragged, target)

        self._refresh_tree()
        self._auto_save()

    def _clear_editor(self):
        self.name_edit.clear()
        self.mode_all.setChecked(True)
        for w in self.condition_widgets:
            self.conditions_layout.removeWidget(w)
            w.deleteLater()
        self.condition_widgets.clear()
        self._set_editor_enabled(False)

    def _load_set_to_editor(self, item: SearchSetItem):
        self._clear_editor()
        self._set_editor_enabled(True)
        self.name_edit.setText(item.name)
        mode = item.data.get('mode', 'all')
        self.mode_any.setChecked(mode == 'any')
        self.mode_all.setChecked(mode != 'any')

        for cond in item.data.get('conditions', []):
            self._add_condition(cond)

    def _set_editor_enabled(self, enabled: bool):
        if hasattr(self, "editor_hint"):
            self.editor_hint.setVisible(not enabled)
        self.name_edit.setEnabled(enabled)
        self.mode_all.setEnabled(enabled)
        self.mode_any.setEnabled(enabled)
        self.add_condition_btn.setEnabled(enabled)
        self.conditions_widget.setEnabled(enabled)

    def _save_current_item(self):
        if not self.current_item:
            return
        self.current_item.name = self.name_edit.text()
        self.current_item.data['mode'] = 'any' if self.mode_any.isChecked() else 'all'
        conditions = [w.get_data() for w in self.condition_widgets]
        self.current_item.data['conditions'] = conditions
        self._update_categories_history_from_conditions(conditions)

    def _update_categories_history_from_conditions(self, conditions: List[dict]):
        """Persist new categories/properties entered by the user."""
        if not conditions:
            return

        changed = False
        categories = self.categories_history.setdefault("categories", [])
        properties = self.categories_history.setdefault("properties", {})

        for cond in conditions:
            cat = cond.get("category")
            cat_int = cond.get("category_internal") or cat
            prop = cond.get("property")
            prop_int = cond.get("property_internal") or prop

            if cat:
                if not any(c.get("name") == cat for c in categories):
                    categories.append({"name": cat, "internal": cat_int or cat})
                    changed = True

                if prop:
                    props = properties.setdefault(cat, [])
                    if not any(p.get("name") == prop for p in props):
                        props.append({"name": prop, "internal": prop_int or prop})
                        changed = True

        if changed:
            save_categories_history(self.categories_history)

    def _on_name_changed(self):
        self._save_current_item()
        self._refresh_tree()
        self._auto_save()

    def _add_condition(self, data: Optional[dict] = None):
        idx = len(self.condition_widgets)
        widget = ConditionWidget(idx, self.categories_history, data)
        widget.deleted.connect(self._on_condition_deleted)
        widget.moveUp.connect(self._on_condition_move_up)
        widget.moveDown.connect(self._on_condition_move_down)
        self.conditions_layout.addWidget(widget)
        self.condition_widgets.append(widget)

    def _on_condition_deleted(self, widget: ConditionWidget):
        self.condition_widgets.remove(widget)
        self.conditions_layout.removeWidget(widget)
        widget.deleteLater()
        self._auto_save()

    def _on_condition_move_up(self, widget: ConditionWidget):
        idx = self.condition_widgets.index(widget)
        if idx > 0:
            self.condition_widgets[idx], self.condition_widgets[idx - 1] = \
                self.condition_widgets[idx - 1], self.condition_widgets[idx]
            self._rebuild_conditions_ui()

    def _on_condition_move_down(self, widget: ConditionWidget):
        idx = self.condition_widgets.index(widget)
        if idx < len(self.condition_widgets) - 1:
            self.condition_widgets[idx], self.condition_widgets[idx + 1] = \
                self.condition_widgets[idx + 1], self.condition_widgets[idx]
            self._rebuild_conditions_ui()

    def _rebuild_conditions_ui(self):
        data = [w.get_data() for w in self.condition_widgets]
        for w in self.condition_widgets:
            self.conditions_layout.removeWidget(w)
            w.deleteLater()
        self.condition_widgets.clear()
        for d in data:
            self._add_condition(d)
        self._auto_save()

    def _add_folder(self):
        name, ok = QtWidgets.QInputDialog.getText(self, "Новая папка", "Имя папки:")
        if not ok or not name:
            return
        new_folder = SearchSetItem(name, is_folder=True)
        selected = self.tree.selected_model_item()
        if selected and selected.is_folder:
            selected.add_child(new_folder)
        else:
            self.root_item.add_child(new_folder)
        self._refresh_tree()
        self._auto_save()

    def _add_set(self):
        name, ok = QtWidgets.QInputDialog.getText(self, "Новый набор", "Имя набора:")
        if not ok or not name:
            return
        new_set = SearchSetItem(name, is_folder=False, data={'mode': 'all', 'conditions': []})
        selected = self.tree.selected_model_item()
        if selected and selected.is_folder:
            selected.add_child(new_set)
        else:
            self.root_item.add_child(new_set)
        self._refresh_tree()
        self._auto_save()

    def _copy_item(self):
        selected = self.tree.selected_model_item()
        if not selected:
            QtWidgets.QMessageBox.warning(self, "Копирование", "Выберите элемент")
            return

        def deep_copy(item: SearchSetItem) -> SearchSetItem:
            new_item = SearchSetItem(item.name, is_folder=item.is_folder)
            new_item.data = copy.deepcopy(item.data)
            for c in item.children:
                new_item.add_child(deep_copy(c))
            return new_item

        new_item = deep_copy(selected)
        new_item.name += " (копия)"

        if selected.parent:
            selected.parent.add_child(new_item)
        else:
            self.root_item.add_child(new_item)
        self._refresh_tree()
        self._auto_save()

    def _rename_item(self):
        selected = self.tree.selected_model_item()
        if not selected:
            QtWidgets.QMessageBox.warning(self, "Переименование", "Выберите элемент")
            return

        item_type = "папки" if selected.is_folder else "набора"
        new_name, ok = QtWidgets.QInputDialog.getText(
            self, "Переименование",
            f"Новое имя {item_type}:",
            text=selected.name
        )
        if ok and new_name and new_name != selected.name:
            selected.name = new_name
            # Update editor if this is the current item
            if self.current_item is selected:
                self.name_edit.setText(new_name)
            self._refresh_tree()
            self._auto_save()

    def _filter_items_for_delete(self, items: List[SearchSetItem]) -> List[SearchSetItem]:
        """Убрать элементы, у которых предок тоже выбран (удаляем только верхние узлы)."""
        result: List[SearchSetItem] = []
        for item in items:
            skip = False
            for other in items:
                if other is item:
                    continue
                if other.is_ancestor_of(item):
                    skip = True
                    break
            if not skip:
                result.append(item)
        return result

    def _delete_items(self, items: List[SearchSetItem]):
        """Удалить один или несколько элементов, сохранив структуру дерева."""
        if not items:
            return

        # Убираем дубликаты и вложенные элементы
        unique_items: List[SearchSetItem] = []
        seen = set()
        for it in items:
            if it not in seen:
                seen.add(it)
                unique_items.append(it)
        unique_items = self._filter_items_for_delete(unique_items)

        if not unique_items:
            return

        if len(unique_items) == 1:
            question = "Удалить выбранный элемент?"
        else:
            question = f"Удалить выделенные элементы ({len(unique_items)})?"

        if QtWidgets.QMessageBox.question(self, "Удаление", question) != QtWidgets.QMessageBox.Yes:
            return

        for item in unique_items:
            if item.parent:
                item.parent.remove_child(item)
            else:
                self.root_item.remove_child(item)

        self._clear_editor()
        # Сохраняем развёрнутость/структуру после удаления
        self._refresh_tree()
        self._auto_save()

    def _delete_item(self):
        """Удаление через кнопку — поддерживает множественный выбор."""
        items = self.tree.selected_model_items()
        if not items:
            return
        self._delete_items(items)

    def _on_tree_delete_requested(self, items: List[SearchSetItem]):
        """Удаление по клавише Delete в дереве."""
        self._delete_items(items)

    def _on_tree_context_collapse(self):
        """Свернуть выбранную папку и всё, что под ней, из контекстного меню."""
        qitem = self.tree.currentItem()
        if not qitem:
            return
        model = qitem.data(0, QtCore.Qt.UserRole)
        if not isinstance(model, SearchSetItem) or not model.is_folder:
            return

        def collapse_rec(item: QtWidgets.QTreeWidgetItem):
            for i in range(item.childCount()):
                collapse_rec(item.child(i))
            item.setExpanded(False)

        collapse_rec(qitem)

    def _new_project(self):
        if QtWidgets.QMessageBox.question(
            self, "Новый проект", "Создать новый проект?\nТекущий проект будет закрыт."
        ) != QtWidgets.QMessageBox.Yes:
            return
        self.root_item = SearchSetItem("Корень", is_folder=True)
        self.current_project_path = None
        self._clear_editor()
        self._refresh_tree()
        self._update_title()

    def _load_json(self):
        ensure_projects_dir()
        start_dir = os.path.dirname(self.current_project_path) if self.current_project_path else PROJECTS_DIR
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Открыть проект", start_dir, "JSON Files (*.json)")
        if not path:
            return
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self.root_item = SearchSetItem("Корень", is_folder=True)
            for f_data in data.get('folders', []):
                self.root_item.add_child(SearchSetItem.from_dict(f_data, is_folder=True))
            for s_data in data.get('sets', []):
                self.root_item.add_child(SearchSetItem.from_dict(s_data, is_folder=False))
            self.current_project_path = path
            self._clear_editor()
            self._refresh_tree()
            self._update_title()
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Ошибка", str(e))

    def _save_json(self):
        """Save project. If no current file, show Save As dialog."""
        ensure_projects_dir()
        if self.current_project_path:
            # Уже есть файл - сохраняем молча
            self._auto_save()
            return

        # Нет файла - показываем диалог выбора
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Сохранить проект", PROJECTS_DIR, "JSON Files (*.json)")
        if not path:
            return
        try:
            self._save_current_item()
            data = {
                'name': os.path.splitext(os.path.basename(path))[0],
                'folders': [c.to_dict() for c in self.root_item.children if c.is_folder],
                'sets': [c.to_dict() for c in self.root_item.children if not c.is_folder]
            }
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self.current_project_path = path
            self._update_title()
            QtWidgets.QMessageBox.information(self, "Сохранено", f"Проект сохранен:\n{path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Ошибка", str(e))

    def _save_json_as(self):
        """Save project to a new file (always shows dialog)."""
        ensure_projects_dir()
        start_dir = os.path.dirname(self.current_project_path) if self.current_project_path else PROJECTS_DIR
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Сохранить проект как...", start_dir, "JSON Files (*.json)")
        if not path:
            return
        try:
            self._save_current_item()
            data = {
                'name': os.path.splitext(os.path.basename(path))[0],
                'folders': [c.to_dict() for c in self.root_item.children if c.is_folder],
                'sets': [c.to_dict() for c in self.root_item.children if not c.is_folder]
            }
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self.current_project_path = path
            self._update_title()
            QtWidgets.QMessageBox.information(self, "Сохранено", f"Проект сохранен:\n{path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Ошибка", str(e))

    def _import_xml(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Импорт из Navisworks", "", "XML Files (*.xml)")
        if not path:
            return
        try:
            project = parse_xml_to_project(path)
            self.root_item = SearchSetItem("Корень", is_folder=True)
            for f_data in project.get('folders', []):
                self.root_item.add_child(SearchSetItem.from_dict(f_data, is_folder=True))
            for s_data in project.get('sets', []):
                self.root_item.add_child(SearchSetItem.from_dict(s_data, is_folder=False))
            self._update_categories_from_project(project)
            # Импорт XML не устанавливает путь проекта (разные форматы)
            self.current_project_path = None
            self._clear_editor()
            self._refresh_tree()
            self._update_title()
            QtWidgets.QMessageBox.information(self, "Импорт", "XML импортирован")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Ошибка", str(e))

    def _update_categories_from_project(self, project: dict):
        def process_conditions(conditions):
            for cond in conditions:
                cat = cond.get('category')
                cat_int = cond.get('category_internal')
                prop = cond.get('property')
                prop_int = cond.get('property_internal')

                if cat and cat_int:
                    exists = any(c['name'] == cat for c in self.categories_history.get('categories', []))
                    if not exists:
                        self.categories_history.setdefault('categories', []).append(
                            {'name': cat, 'internal': cat_int})

                if cat and prop and prop_int:
                    props = self.categories_history.setdefault('properties', {}).setdefault(cat, [])
                    exists = any(p['name'] == prop for p in props)
                    if not exists:
                        props.append({'name': prop, 'internal': prop_int})

        def process_folder(folder):
            for s in folder.get('sets', []):
                process_conditions(s.get('conditions', []))
            for f in folder.get('folders', []):
                process_folder(f)

        for f in project.get('folders', []):
            process_folder(f)
        for s in project.get('sets', []):
            process_conditions(s.get('conditions', []))

        save_categories_history(self.categories_history)

    def _export_xml(self):
        self._save_current_item()
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Экспорт в Navisworks", "", "XML Files (*.xml)")
        if not path:
            return
        try:
            data = {
                'folders': [c.to_dict() for c in self.root_item.children if c.is_folder],
                'sets': [c.to_dict() for c in self.root_item.children if not c.is_folder]
            }
            generate_xml_from_project(data, path)
            QtWidgets.QMessageBox.information(self, "Экспорт", f"XML сохранен:\n{path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Ошибка", str(e))

    def closeEvent(self, event):
        # Save current editor state before closing
        self._save_current_item()
        self._auto_save()

        # Quit Excel if we created it ourselves
        if self._excel_created_by_us and self.excel:
            try:
                self.excel.Quit()
            except Exception:
                pass
            finally:
                self.excel = None
                self.workbook = None
                self.worksheet = None

        if self._com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        event.accept()


def main():
    app = QtWidgets.QApplication(sys.argv)

    font = app.font()
    font.setPointSize(font.pointSize() + 1)
    app.setFont(font)

    app.setStyle('Fusion')

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
