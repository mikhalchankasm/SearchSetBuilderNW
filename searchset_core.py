# -*- coding: utf-8 -*-
"""Pure domain helpers for search sets: condition presets, batch creation, validation.

No Qt and no Excel here — this module is safe to unit-test headless and is the
single source of truth for *correct* Navisworks internal names (a wrong
``*_internal`` is the usual cause of a set that matches zero objects).
"""
from __future__ import annotations

from typing import Dict, List

# --- Vetted category/property internal names -------------------------------
# Keep the display name and its Navisworks internal id together so the UI never
# has to guess an internal from a Russian label.
CONDITION_PRESETS: List[Dict[str, str]] = [
    {"label": "Имя (Элемент)",
     "category": "Элемент", "category_internal": "LcOaNode",
     "property": "Имя", "property_internal": "LcOaSceneBaseUserName",
     "data_type": "wstring"},
    {"label": "Файл источника (Элемент)",
     "category": "Элемент", "category_internal": "LcOaNode",
     "property": "Файл источника", "property_internal": "LcOaNodeSourceFile",
     "data_type": "wstring"},
    {"label": "AVEVA · Наименование",
     "category": "AVEVA", "category_internal": "lcldrvm_props",
     "property": "Наименование", "property_internal": "lcldrvm_prop_Наименование",
     "data_type": "wstring"},
]

# --- Batch-create strategies (preset + default comparison test) ------------
MATCH_STRATEGIES: List[Dict[str, str]] = [
    {"key": "name_equals", "preset": "Имя (Элемент)", "test": "equals",
     "label": "Имя = строка (точное совпадение, 1 узел на файл)"},
    {"key": "source_equals", "preset": "Файл источника (Элемент)", "test": "equals",
     "label": "Файл источника = строка (вся геометрия из файла)"},
    {"key": "name_wildcard", "preset": "Имя (Элемент)", "test": "wildcard",
     "label": "Имя — подстановочный знак (* ?)"},
    {"key": "name_contains", "preset": "Имя (Элемент)", "test": "contains",
     "label": "Имя содержит строку"},
    {"key": "aveva_contains", "preset": "AVEVA · Наименование", "test": "contains",
     "label": "AVEVA · Наименование содержит строку"},
]

_PRESET_BY_LABEL = {p["label"]: p for p in CONDITION_PRESETS}


def preset_condition(preset_label: str, test: str = "equals", value: str = "") -> dict:
    """Build a condition dict from a preset label (with correct internal names)."""
    p = _PRESET_BY_LABEL[preset_label]
    return {
        "category": p["category"], "category_internal": p["category_internal"],
        "property": p["property"], "property_internal": p["property_internal"],
        "test": test, "value": value, "data_type": p["data_type"],
    }


def strategy_by_key(key: str) -> dict:
    for s in MATCH_STRATEGIES:
        if s["key"] == key:
            return s
    raise KeyError(key)


def build_sets_from_names(names, strategy_key: str,
                          name_template: str = "{name}") -> List[dict]:
    """Turn raw name lines into selection-set dicts (one condition each).

    Blank lines are skipped. ``name_template`` may reference ``{name}``.
    """
    strat = strategy_by_key(strategy_key)
    sets: List[dict] = []
    for raw in names:
        name = (raw or "").strip()
        if not name:
            continue
        try:
            set_name = name_template.format(name=name)
        except (KeyError, IndexError, ValueError):
            set_name = name
        cond = preset_condition(strat["preset"], test=strat["test"], value=name)
        sets.append({"name": set_name, "mode": "all", "conditions": [cond]})
    return sets


_TESTS_NEED_VALUE = {"equals", "not_equals", "contains", "wildcard"}


def validate_project(project: dict) -> List[str]:
    """Return human-readable issues for a project dict (empty list = all good)."""
    issues: List[str] = []
    dup: Dict[tuple, int] = {}

    def check_set(s: dict, folder_path: str):
        name = (s.get("name") or "").strip()
        where = f"{folder_path or 'корень'} / {name or '<без имени>'}"
        if not name:
            issues.append(f"Набор без имени (папка «{folder_path or 'корень'}»)")
        else:
            dup[(folder_path, name)] = dup.get((folder_path, name), 0) + 1
        conds = s.get("conditions", [])
        if not conds:
            issues.append(f"Набор без условий: {where}")
        for i, c in enumerate(conds, 1):
            test = c.get("test", "")
            value = (c.get("value") or "").strip()
            if test in _TESTS_NEED_VALUE and not value:
                issues.append(f"Пустое значение (условие {i}, тест «{test}»): {where}")

    def walk(folder: dict, parent_path: str):
        fname = folder.get("name", "")
        fpath = f"{parent_path} / {fname}" if parent_path else fname
        for s in folder.get("sets", []):
            check_set(s, fpath)
        for f in folder.get("folders", []):
            walk(f, fpath)

    for f in project.get("folders", []):
        walk(f, "")
    for s in project.get("sets", []):
        check_set(s, "")

    for (folder_path, name), cnt in dup.items():
        if cnt > 1:
            issues.append(
                f"Повтор имени «{name}» в папке «{folder_path or 'корень'}» ({cnt}×)")
    return issues
