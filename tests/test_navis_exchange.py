# -*- coding: utf-8 -*-
"""Unit tests for the pure Navisworks exchange-XML layer (no Qt / no Excel)."""
import xml.etree.ElementTree as ET

import navis_exchange as nx


def test_findspec_mode_to_xml():
    # 'any' is encoded as findspec mode="all" + disjoint="1"
    assert nx.findspec_mode_to_xml("any") == ("all", "1")
    assert nx.findspec_mode_to_xml("all") == ("all", "0")
    assert nx.findspec_mode_to_xml("selected") == ("selected", "0")
    assert nx.findspec_mode_to_xml("below") == ("below", "0")
    # unknown falls back to AND
    assert nx.findspec_mode_to_xml("garbage") == ("all", "0")


def test_condition_flags_and_chain():
    assert nx.condition_flags_to_xml("all", 0) == "0"
    assert nx.condition_flags_to_xml("all", 4) == "0"


def test_condition_flags_or_chain():
    # first condition of an OR group is 26, the rest are 90
    assert nx.condition_flags_to_xml("any", 0) == "26"
    assert nx.condition_flags_to_xml("any", 1) == "90"
    assert nx.condition_flags_to_xml("any", 7) == "90"


def test_infer_mode_from_flags():
    assert nx.infer_condition_mode_from_flags([]) == "all"
    assert nx.infer_condition_mode_from_flags([{"flags": "0"}]) == "all"
    assert nx.infer_condition_mode_from_flags(
        [{"flags": "26"}, {"flags": "90"}, {"flags": "90"}]) == "any"
    assert nx.infer_condition_mode_from_flags(
        [{"flags": "0"}, {"flags": "0"}]) == "all"
    # malformed flags must not raise
    assert nx.infer_condition_mode_from_flags(
        [{"flags": "x"}, {"flags": "y"}]) == "all"


def test_make_condition_advanced_structure_and_cyrillic():
    el = nx.make_condition_advanced(
        "AVEVA", "lcldrvm_props", "Наименование", "lcldrvm_prop_Наименование",
        "contains", "Отвод", "wstring", flags="26")
    assert el.tag == "condition"
    assert el.get("test") == "contains"
    assert el.get("flags") == "26"
    assert el.find("category/name").text == "AVEVA"
    assert el.find("category/name").get("internal") == "lcldrvm_props"
    assert el.find("property/name").text == "Наименование"
    assert el.find("property/name").get("internal") == "lcldrvm_prop_Наименование"
    data = el.find("value/data")
    assert data.get("type") == "wstring" and data.text == "Отвод"


def _sample_project():
    return {
        "folders": [{
            "name": "СО", "folders": [], "sets": [{
                "name": "Отвод", "mode": "any", "conditions": [
                    {"category": "AVEVA", "category_internal": "lcldrvm_props",
                     "property": "Наименование",
                     "property_internal": "lcldrvm_prop_Наименование",
                     "test": "contains", "value": "Отвод", "data_type": "wstring"},
                    {"category": "Элемент", "category_internal": "LcOaNode",
                     "property": "Файл источника",
                     "property_internal": "LcOaNodeSourceFile",
                     "test": "equals", "value": "240101-НК1.rvm",
                     "data_type": "wstring"},
                ]}]}],
        "sets": [],
    }


def test_generate_xml_is_valid_and_or_encoded(tmp_path):
    out = tmp_path / "sets.xml"
    nx.generate_xml_from_project(_sample_project(), str(out))

    text = out.read_text(encoding="utf-8")
    assert "nw-exchange-12.0.xsd" in text
    assert "Отвод" in text and "Файл источника" in text  # cyrillic survives

    root = ET.parse(str(out)).getroot()
    assert root.tag == "exchange"
    fs = root.find(".//selectionset/findspec")
    assert fs.get("disjoint") == "1"          # 'any' -> disjoint
    conds = root.findall(".//conditions/condition")
    assert [c.get("flags") for c in conds] == ["26", "90"]   # OR chain
    assert root.find(".//locator").text == "/"
