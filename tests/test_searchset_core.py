# -*- coding: utf-8 -*-
"""Tests for the pure search-set domain layer (presets, batch create, validation)."""
import searchset_core as core


def test_preset_condition_has_correct_internals():
    c = core.preset_condition("Файл источника (Элемент)", test="equals", value="X.rvm")
    assert c["category_internal"] == "LcOaNode"
    assert c["property_internal"] == "LcOaNodeSourceFile"
    assert c["test"] == "equals" and c["value"] == "X.rvm"
    assert c["data_type"] == "wstring"


def test_build_sets_name_equals_strategy():
    names = ["240000-АС1.rvm", "  ", "240100-ТК2.rvm", ""]
    sets = core.build_sets_from_names(names, "name_equals")
    assert len(sets) == 2  # blanks skipped
    assert sets[0]["name"] == "240000-АС1.rvm"
    cond = sets[0]["conditions"][0]
    assert cond["property_internal"] == "LcOaSceneBaseUserName"
    assert cond["test"] == "equals"
    assert cond["value"] == "240000-АС1.rvm"


def test_build_sets_source_strategy_and_template():
    sets = core.build_sets_from_names(
        ["240148-ТСГ1.rvm"], "source_equals", name_template="RVM · {name}")
    assert sets[0]["name"] == "RVM · 240148-ТСГ1.rvm"
    assert sets[0]["conditions"][0]["property_internal"] == "LcOaNodeSourceFile"


def test_build_sets_bad_template_falls_back_to_name():
    sets = core.build_sets_from_names(["X.rvm"], "name_equals", name_template="{oops}")
    assert sets[0]["name"] == "X.rvm"


def test_validate_clean_project_has_no_issues():
    proj = {"folders": [{"name": "СО", "folders": [], "sets": [
        {"name": "A", "mode": "all", "conditions": [
            core.preset_condition("Имя (Элемент)", "equals", "A.rvm")]}]}], "sets": []}
    assert core.validate_project(proj) == []


def test_validate_catches_empty_value_missing_conditions_and_dupes():
    proj = {"folders": [{"name": "F", "folders": [], "sets": [
        {"name": "no-cond", "mode": "all", "conditions": []},
        {"name": "empty-val", "mode": "all", "conditions": [
            core.preset_condition("Имя (Элемент)", "equals", "")]},
        {"name": "dup", "mode": "all", "conditions": [
            core.preset_condition("Имя (Элемент)", "equals", "x")]},
        {"name": "dup", "mode": "all", "conditions": [
            core.preset_condition("Имя (Элемент)", "equals", "y")]},
    ]}], "sets": []}
    issues = core.validate_project(proj)
    joined = " | ".join(issues)
    assert "без условий" in joined
    assert "Пустое значение" in joined
    assert "Повтор имени" in joined and "dup" in joined
