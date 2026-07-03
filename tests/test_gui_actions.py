# -*- coding: utf-8 -*-
"""Offscreen GUI tests for the new UI actions (mass-create, undo/redo, presets)."""
import pytest

pytest.importorskip("PySide6")

import search_set as ss              # noqa: E402
from searchset_core import build_sets_from_names  # noqa: E402


@pytest.fixture
def win(qapp):
    w = ss.MainWindow()
    yield w
    w.close()


def test_mass_create_inserts_sets_with_correct_internal(win):
    sets = build_sets_from_names(["A.rvm", "B.rvm"], "source_equals")
    n = win._add_sets_to_target(sets, "RVM", None)
    assert n == 2
    rvm = next(c for c in win.root_item.children if c.is_folder and c.name == "RVM")
    assert len(rvm.children) == 2
    cond = rvm.children[0].data["conditions"][0]
    assert cond["property_internal"] == "LcOaNodeSourceFile"
    assert cond["test"] == "equals"


def test_undo_redo_structural(win):
    win._add_sets_to_target(build_sets_from_names(["X.rvm"], "name_equals"), "F", None)
    assert any(c.name == "F" for c in win.root_item.children)
    win._undo()
    assert not any(c.name == "F" for c in win.root_item.children)
    win._redo()
    assert any(c.name == "F" for c in win.root_item.children)


def test_preset_condition_internal_survives_get_data(win):
    win._add_sets_to_target(build_sets_from_names(["X.rvm"], "name_equals"), "F", None)
    s = next(c for c in win.root_item.children if c.name == "F").children[0]
    win.current_item = s
    win._load_set_to_editor(s)
    win._add_preset_condition("Файл источника (Элемент)")
    data = win.condition_widgets[-1].get_data()
    assert data["property_internal"] == "LcOaNodeSourceFile"


def test_clipboard_copy_paste(win):
    win._add_sets_to_target(build_sets_from_names(["X.rvm"], "name_equals"), "F", None)
    folder = next(c for c in win.root_item.children if c.name == "F")
    win._clipboard_item = folder.to_dict()
    before = len(win.root_item.children)
    win._paste_from_clipboard()
    assert len(win.root_item.children) == before + 1
    assert any(c.name.endswith("(вставка)") for c in win.root_item.children)
