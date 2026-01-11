"""
Heuristic mapping for converting shapes without an LLM.
"""

from typing import Dict, Any, List


def _normalize_name(name: str) -> str:
    return "".join(ch.lower() for ch in name if ch.isalnum())

def _normalize_text(text: str) -> str:
    return "".join(ch.lower() for ch in text.strip() if ch.isalnum())


# Role labels that should be mapped to Process (not Swimlane)
ROLE_LABELS = {"部门负责人", "分管领导"}
ROLE_LABELS_NORMALIZED = {_normalize_text(label) for label in ROLE_LABELS}

# Swimlane header labels - department/organization names
SWIMLANE_HEADERS = {"经办人", "财务部", "法务部", "业务部", "管理层", "董事长", "需求部门", "采购部"}
SWIMLANE_HEADERS_NORMALIZED = {_normalize_text(label) for label in SWIMLANE_HEADERS}


class AutoMapper:
    """Map shapes to target masters using name-based heuristics."""

    _CONNECTOR_LABELS = {
        "通过",
        "未通过",
        "合同金额超过xx元",
        "合同金额不超过xx元",
    }
    _CONNECTOR_LABELS_NORMALIZED = {_normalize_text(label) for label in _CONNECTOR_LABELS}

    _SYNONYMS = {
        "rectangle": ["process", "rounded rectangle"],
        "roundedrectangle": ["rounded rectangle", "process"],
        "diamond": ["decision"],
        "decision": ["decision"],
        "startend": ["start/end"],
        "terminator": ["start/end"],
        "dynamicconnector": ["dynamic connector", "connector"],
        "connector": ["dynamic connector"],
        "simplearrow": ["dynamic connector"],
        "frame": ["cff container", "swimlane", "swimlane (vertical)", "swimlane list"],
        "decagon": ["decision"],
    }

    def create_mapping(
        self,
        diagram_data: Dict[str, Any],
        target_masters: List[Dict[str, str]],
    ) -> Dict[str, str]:
        target_names = [m.get("name", "") for m in target_masters if m.get("name")]
        target_by_norm = {_normalize_name(name): name for name in target_names}

        mapping: Dict[str, str] = {}

        for page in diagram_data.get("pages", []):
            for shape in page.get("shapes", []):
                self._map_shape(shape, target_by_norm, mapping, is_connector=False)
            for connector in page.get("connectors", []):
                self._map_shape(connector, target_by_norm, mapping, is_connector=True)

        return mapping

    def _map_shape(
        self,
        shape: Dict[str, Any],
        target_by_norm: Dict[str, str],
        mapping: Dict[str, str],
        is_connector: bool,
    ) -> None:
        shape_id = shape.get("id")
        if shape_id is None:
            return
        shape_id = str(shape_id)

        master_name = shape.get("master_name") or ""
        master_norm = _normalize_name(master_name)
        shape_text = (shape.get("text") or "").strip()
        text_norm = _normalize_text(shape_text) if shape_text else ""
        size = shape.get("size") or {}
        width = size.get("width")
        height = size.get("height")

        if is_connector:
            mapped = self._match_targets(["dynamic connector", "connector"], target_by_norm)
            if mapped:
                mapping[shape_id] = mapped
            return

        if shape_text and text_norm in self._CONNECTOR_LABELS_NORMALIZED:
            if width is None or height is None or (width <= 0.8 and height <= 0.4) or (width == 1.0 and height == 1.0):
                mapped = self._match_targets(["dynamic connector", "connector"], target_by_norm)
                if mapped:
                    mapping[shape_id] = mapped
                    return

        # Role labels (部门负责人, 分管领导) should map to Process, not Swimlane
        if text_norm in ROLE_LABELS_NORMALIZED:
            mapped = self._match_targets(["process", "rounded rectangle"], target_by_norm)
            if mapped:
                mapping[shape_id] = mapped
                return

        if not master_norm and shape_text:
            text_len = len(text_norm)

            if text_norm in self._CONNECTOR_LABELS_NORMALIZED:
                mapped = self._match_targets(["dynamic connector", "connector"], target_by_norm)
                if mapped:
                    mapping[shape_id] = mapped
                    return

            if width is not None and height is not None:
                if width <= 0.8 and height <= 0.4:
                    mapped = self._match_targets(["dynamic connector", "connector"], target_by_norm)
                    if mapped:
                        mapping[shape_id] = mapped
                        return
                if width <= 2.5 and height <= 1.1:
                    mapped = self._match_targets(["process", "rounded rectangle"], target_by_norm)
                    if mapped:
                        mapping[shape_id] = mapped
                        return
            if 0 < text_len <= 6 and "流程" not in shape_text:
                mapped = self._match_targets(["process", "rounded rectangle"], target_by_norm)
                if mapped:
                    mapping[shape_id] = mapped
                    return

        # Direct match
        if master_norm in target_by_norm:
            mapping[shape_id] = target_by_norm[master_norm]
            return

        # Synonyms
        for candidate in self._SYNONYMS.get(master_norm, []):
            mapped = self._match_targets([candidate], target_by_norm)
            if mapped:
                mapping[shape_id] = mapped
                return

        # Keyword heuristics
        keywords = []
        if "rect" in master_norm:
            keywords.extend(["process", "rounded rectangle"])
        if "diamond" in master_norm or "decision" in master_norm:
            keywords.append("decision")
        if "start" in master_norm or "end" in master_norm:
            keywords.append("start/end")
        if "connector" in master_norm or "arrow" in master_norm:
            keywords.append("dynamic connector")

        mapped = self._match_targets(keywords, target_by_norm)
        if mapped:
            mapping[shape_id] = mapped

    @staticmethod
    def _match_targets(candidates: List[str], target_by_norm: Dict[str, str]) -> str:
        for candidate in candidates:
            norm = _normalize_name(candidate)
            if norm in target_by_norm:
                return target_by_norm[norm]
            # Fuzzy containment match
            for target_norm, target_name in target_by_norm.items():
                if norm in target_norm or target_norm in norm:
                    return target_name
        return ""
