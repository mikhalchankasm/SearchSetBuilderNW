# -*- coding: utf-8 -*-
"""Navisworks nw-exchange 12.0 XML helpers (stdlib only, no Qt)."""

from __future__ import annotations

import uuid
import xml.dom.minidom
import xml.etree.ElementTree as ET
from typing import List

_DEFAULT_CATEGORY = "\u042d\u043b\u0435\u043c\u0435\u043d\u0442"
_DEFAULT_PROP_NAME = "\u0418\u043c\u044f"

NAVIS_FINDSPEC_MODES = frozenset({"all", "selected", "below"})
NAVIS_COND_FLAG_GROUP_START = 26
NAVIS_COND_FLAG_OR_NEXT = 90


def findspec_mode_to_xml(mode: str) -> tuple[str, str]:
    """(mode_attr, disjoint) for findspec in exchange XML."""
    if mode == "any":
        return "all", "1"
    if mode in NAVIS_FINDSPEC_MODES:
        return mode, "0"
    return "all", "0"


def condition_flags_to_xml(set_mode: str, condition_index: int) -> str:
    """condition/@flags for OR chains (26, 90, ...) vs AND (0)."""
    if set_mode == "any":
        return str(NAVIS_COND_FLAG_OR_NEXT if condition_index > 0 else NAVIS_COND_FLAG_GROUP_START)
    return "0"


def infer_condition_mode_from_flags(conditions: List[dict]) -> str:
    """Infer mode 'any' from Navisworks flags chain (26 + 90...)."""
    if len(conditions) < 2:
        return "all"
    try:
        nums = [int(str(c.get("flags", "0"))) for c in conditions]
    except ValueError:
        return "all"
    first, rest = nums[0], nums[1:]
    if first == NAVIS_COND_FLAG_OR_NEXT:
        return "all"
    if all(r == NAVIS_COND_FLAG_OR_NEXT for r in rest):
        return "any"
    return "all"


def make_condition_advanced(
    category: str,
    category_internal: str,
    property_name: str,
    property_internal: str,
    test: str,
    value: str,
    data_type: str = "wstring",
    flags: str = "0",
) -> ET.Element:
    cond = ET.Element("condition", test=test, flags=flags)
    cat = ET.SubElement(cond, "category")
    ET.SubElement(cat, "name", internal=category_internal).text = category
    prop = ET.SubElement(cond, "property")
    ET.SubElement(prop, "name", internal=property_internal).text = property_name
    val = ET.SubElement(cond, "value")
    ET.SubElement(val, "data", type=data_type).text = value
    return cond


def generate_xml_from_project(project_data: dict, save_path: str) -> None:
    exchange = ET.Element("exchange")
    exchange.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
    exchange.set(
        "xsi:noNamespaceSchemaLocation",
        "http://download.autodesk.com/us/navisworks/schemas/nw-exchange-12.0.xsd",
    )
    exchange.set("units", "m")
    exchange.set("filename", "")
    exchange.set("filepath", "")

    selectionsets = ET.SubElement(exchange, "selectionsets")

    def add_folder(parent_elem: ET.Element, folder_data: dict) -> None:
        vf = ET.SubElement(parent_elem, "viewfolder", name=folder_data["name"], guid=str(uuid.uuid4()))
        for subfolder in folder_data.get("folders", []):
            add_folder(vf, subfolder)
        for s in folder_data.get("sets", []):
            add_set(vf, s)

    def add_set(parent_elem: ET.Element, set_data: dict) -> None:
        ss = ET.SubElement(parent_elem, "selectionset", name=set_data["name"], guid=str(uuid.uuid4()))
        mode = set_data.get("mode", "all")
        fs_mode, fs_disjoint = findspec_mode_to_xml(mode)
        findspec = ET.SubElement(ss, "findspec", mode=fs_mode, disjoint=fs_disjoint)
        conditions_elem = ET.SubElement(findspec, "conditions")

        for i, cond in enumerate(set_data.get("conditions", [])):
            cond_elem = make_condition_advanced(
                category=cond.get("category", _DEFAULT_CATEGORY),
                category_internal=cond.get("category_internal", "LcOaNode"),
                property_name=cond.get("property", _DEFAULT_PROP_NAME),
                property_internal=cond.get("property_internal", "LcOaSceneBaseUserName"),
                test=cond.get("test", "contains"),
                value=cond.get("value", ""),
                data_type=cond.get("data_type", "wstring"),
                flags=condition_flags_to_xml(mode, i),
            )
            conditions_elem.append(cond_elem)

        ET.SubElement(findspec, "locator").text = "/"

    for folder in project_data.get("folders", []):
        add_folder(selectionsets, folder)
    for s in project_data.get("sets", []):
        add_set(selectionsets, s)

    xml_str = ET.tostring(exchange, encoding="utf-8")
    pretty_xml = xml.dom.minidom.parseString(xml_str).toprettyxml(indent="  ", encoding="utf-8")
    with open(save_path, "wb") as f:
        f.write(pretty_xml)
