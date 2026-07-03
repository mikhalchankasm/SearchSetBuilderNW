# -*- coding: utf-8 -*-
"""End-to-end round-trip: project dict -> XML -> project dict.

Exercises ``parse_xml_to_project`` (lives in ``search_set``), so it needs the
GUI deps; skipped cleanly where PySide6 is unavailable (e.g. a headless CI).
"""
import pytest

pytest.importorskip("PySide6")

import navis_exchange as nx  # noqa: E402
import search_set as ss      # noqa: E402


def _project():
    return {
        "folders": [{
            "name": "СО",
            "folders": [{
                "name": "240101-НК1", "folders": [], "sets": [
                    {"name": "any-set", "mode": "any", "conditions": [
                        {"category": "AVEVA", "category_internal": "lcldrvm_props",
                         "property": "Наименование",
                         "property_internal": "lcldrvm_prop_Наименование",
                         "test": "contains", "value": "Отвод", "data_type": "wstring"},
                        {"category": "Элемент", "category_internal": "LcOaNode",
                         "property": "Файл источника",
                         "property_internal": "LcOaNodeSourceFile",
                         "test": "equals", "value": "240101-НК1.rvm",
                         "data_type": "wstring"}]},
                    {"name": "all-set", "mode": "all", "conditions": [
                        {"category": "Элемент", "category_internal": "LcOaNode",
                         "property": "Имя",
                         "property_internal": "LcOaSceneBaseUserName",
                         "test": "equals", "value": "X.rvm",
                         "data_type": "wstring"}]},
                ]}],
            "sets": [],
        }],
        "sets": [],
    }


def test_roundtrip_preserves_modes_cyrillic_and_internals(tmp_path):
    out = tmp_path / "rt.xml"
    nx.generate_xml_from_project(_project(), str(out))
    back = ss.parse_xml_to_project(str(out))

    folder = back["folders"][0]
    assert folder["name"] == "СО"
    sub = folder["folders"][0]
    assert sub["name"] == "240101-НК1"

    by_name = {s["name"]: s for s in sub["sets"]}
    assert by_name["any-set"]["mode"] == "any"   # OR chain inferred back from flags
    assert by_name["all-set"]["mode"] == "all"

    c0, c1 = by_name["any-set"]["conditions"]
    assert c0["property"] == "Наименование"
    assert c0["property_internal"] == "lcldrvm_prop_Наименование"
    assert c0["value"] == "Отвод"
    assert c1["property_internal"] == "LcOaNodeSourceFile"
    assert c1["test"] == "equals"


def test_roundtrip_single_condition_and_wildcard(tmp_path):
    proj = {"folders": [], "sets": [
        {"name": "one", "mode": "all", "conditions": [
            {"category": "Элемент", "category_internal": "LcOaNode",
             "property": "Имя", "property_internal": "LcOaSceneBaseUserName",
             "test": "wildcard", "value": "240*", "data_type": "wstring"}]}]}
    out = tmp_path / "one.xml"
    nx.generate_xml_from_project(proj, str(out))
    back = ss.parse_xml_to_project(str(out))

    s = back["sets"][0]
    assert s["mode"] == "all"
    assert s["conditions"][0]["test"] == "wildcard"
    assert s["conditions"][0]["value"] == "240*"
