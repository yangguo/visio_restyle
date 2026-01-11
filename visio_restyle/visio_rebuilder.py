"""
Rebuild Visio diagram with new masters while preserving layout and connections.
"""

from typing import Dict, Optional, Set, Tuple
from pathlib import Path
import copy
import os
import re
import shutil
import tempfile
import zipfile
import uuid
import xml.etree.ElementTree as ET


NS = {
    "v": "http://schemas.microsoft.com/office/visio/2012/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
ET.register_namespace("", NS["v"])
ET.register_namespace("r", NS["r"])


def _write_xml_with_namespaces(tree: ET.ElementTree, path: Path) -> None:
    """Write XML file ensuring proper namespace declarations are preserved."""
    import io
    output = io.BytesIO()
    tree.write(output, encoding="utf-8", xml_declaration=True)
    xml_content = output.getvalue().decode("utf-8")
    
    # Ensure the r namespace is declared on the root element if not present
    # This is needed because ElementTree drops unused namespace declarations
    r_ns = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    
    # Find the root element and check if it needs the r namespace
    root_match = re.match(r'(<\?xml[^>]+\?>\s*<\w+)([^>]*)(>)', xml_content, re.DOTALL)
    if root_match:
        root_start = root_match.group(1)
        root_attrs = root_match.group(2)
        root_end = root_match.group(3)
        
        # Add xmlns:r if not present and if the file might need it
        if 'xmlns:r=' not in root_attrs and 'r:' in xml_content:
            # Insert before the closing >
            xml_content = root_start + root_attrs + ' ' + r_ns + root_end + xml_content[root_match.end():]
        elif 'xmlns:r=' not in root_attrs:
            # Still add it for Visio compatibility even if not strictly needed
            xml_content = root_start + root_attrs + ' ' + r_ns + root_end + xml_content[root_match.end():]
    
    path.write_text(xml_content, encoding="utf-8")

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
MASTER_REL_TYPE = "http://schemas.microsoft.com/visio/2010/relationships/master"

# Register the relationships namespace as default for .rels files
# This needs to be done separately when writing relationship files


def _write_relationships_xml(rel_root: ET.Element, path: Path) -> None:
    """Write a relationships XML file with proper namespace handling."""
    import io
    
    # Temporarily register the REL_NS as default namespace
    # Save and restore the original namespace registrations
    output = io.BytesIO()
    tree = ET.ElementTree(rel_root)
    tree.write(output, encoding="UTF-8", xml_declaration=True)
    xml_content = output.getvalue().decode("utf-8")
    
    # Replace ns0: prefix with nothing (use default namespace)
    # The proper way is to ensure the root element uses default namespace
    xml_content = xml_content.replace('ns0:', '')
    xml_content = xml_content.replace(':ns0', '')
    xml_content = xml_content.replace('xmlns:ns0=', 'xmlns=')
    
    path.write_text(xml_content, encoding="utf-8")

STYLE_SECTIONS_TO_REMOVE = {
    "Fill",
    "Line",
    "QuickStyle",
    "Image",
    "Geometry",  # Remove geometry so shape inherits from new master
}

STYLE_CELLS_TO_REMOVE = {
    "LineColor",
    "LineColorTrans",
    "LinePattern",
    "LineWeight",
    "LineCap",
    "LineBeginArrow",
    "LineEndArrow",
    "LineBeginArrowSize",
    "LineEndArrowSize",
    "LineRounding",
    "FillForegnd",
    "FillForegndTrans",
    "FillBkgnd",
    "FillBkgndTrans",
    "FillPattern",
    "ShdwForegnd",
    "ShdwForegndTrans",
    "ShdwBkgnd",
    "ShdwBkgndTrans",
    "QuickStyleLineColor",
    "QuickStyleFillColor",
    "QuickStyleShadowColor",
    "QuickStyleFontColor",
    "QuickStyleLineMatrix",
    "QuickStyleFillMatrix",
    "QuickStyleEffectsMatrix",
    "QuickStyleFontMatrix",
}

STYLE_CELLS_TO_COPY = {
    "LineColor",
    "LineColorTrans",
    "LinePattern",
    "LineWeight",
    "LineCap",
    "LineBeginArrow",
    "LineEndArrow",
    "LineBeginArrowSize",
    "LineEndArrowSize",
    "LineRounding",
    "FillForegnd",
    "FillForegndTrans",
    "FillBkgnd",
    "FillBkgndTrans",
    "FillPattern",
    "ShdwForegnd",
    "ShdwForegndTrans",
    "ShdwBkgnd",
    "ShdwBkgndTrans",
    "QuickStyleLineColor",
    "QuickStyleFillColor",
    "QuickStyleShadowColor",
    "QuickStyleFontColor",
    "QuickStyleLineMatrix",
    "QuickStyleFillMatrix",
    "QuickStyleEffectsMatrix",
    "QuickStyleFontMatrix",
}


def _normalize_name(name: str) -> str:
    return "".join(ch.lower() for ch in name if ch.isalnum())


def _normalize_label_text(text: str) -> str:
    return "".join(ch.lower() for ch in text.strip() if ch.isalnum())


def _safe_float(value: Optional[str]) -> Optional[float]:
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _get_shape_text(shape: ET.Element) -> str:
    text_elem = shape.find("v:Text", NS)
    if text_elem is None:
        return ""
    return "".join(text_elem.itertext()).strip()


def _get_cell_value(shape: ET.Element, name: str) -> Optional[str]:
    cell = shape.find(f"v:Cell[@N='{name}']", NS)
    return cell.get("V") if cell is not None else None


def _set_cell_value(shape: ET.Element, name: str, value: str) -> None:
    cell = shape.find(f"v:Cell[@N='{name}']", NS)
    if cell is None:
        cell = ET.Element("{%s}Cell" % NS["v"])
        cell.set("N", name)
        shape.append(cell)
    cell.set("V", str(value))


def _set_shape_text(shape: ET.Element, text: str) -> None:
    text_elem = shape.find("v:Text", NS)
    if text_elem is None:
        text_elem = ET.SubElement(shape, "{%s}Text" % NS["v"])
    for child in list(text_elem):
        text_elem.remove(child)
    text_elem.text = text


def _remove_shape_text(shape: ET.Element) -> None:
    text_elem = shape.find("v:Text", NS)
    if text_elem is not None:
        shape.remove(text_elem)


def _shape_bounds(shape: ET.Element) -> Optional[Tuple[float, float, float, float]]:
    width = _safe_float(_get_cell_value(shape, "Width"))
    height = _safe_float(_get_cell_value(shape, "Height"))
    pinx = _safe_float(_get_cell_value(shape, "PinX"))
    piny = _safe_float(_get_cell_value(shape, "PinY"))
    if width is None or height is None or pinx is None or piny is None:
        return None
    return (pinx - width / 2, pinx + width / 2, piny - height / 2, piny + height / 2)


def _median(values: list) -> Optional[float]:
    values = [value for value in values if value is not None]
    if not values:
        return None
    values.sort()
    mid = len(values) // 2
    if len(values) % 2 == 1:
        return values[mid]
    return (values[mid - 1] + values[mid]) / 2


def _mode(values: list, precision: int = 3) -> Optional[float]:
    rounded = {}
    for value in values:
        if value is None:
            continue
        key = round(value, precision)
        bucket = rounded.setdefault(key, [])
        bucket.append(value)
    if not rounded:
        return None
    best_key = max(rounded.items(), key=lambda item: len(item[1]))[0]
    bucket = rounded[best_key]
    return sum(bucket) / len(bucket)


def _cluster_row_centers(values: list, tolerance: float) -> list:
    values = sorted([value for value in values if value is not None], reverse=True)
    if not values:
        return []
    clusters = [[values[0]]]
    for value in values[1:]:
        current = clusters[-1]
        center = sum(current) / len(current)
        if abs(value - center) <= tolerance:
            current.append(value)
        else:
            clusters.append([value])
    return [sum(cluster) / len(cluster) for cluster in clusters]


X_SCALE_CELLS = {
    "PinX",
    "Width",
    "LocPinX",
    "TxtPinX",
    "TxtLocPinX",
    "BeginX",
    "EndX",
}

Y_SCALE_CELLS = {
    "PinY",
    "Height",
    "LocPinY",
    "TxtPinY",
    "TxtLocPinY",
    "BeginY",
    "EndY",
}


class VisioRebuilder:
    """Rebuild a Visio diagram with new shape masters."""

    def __init__(
        self,
        source_path: str,
        target_template_path: str,
        mapping: Dict[str, str],
    ):
        """Initialize the rebuilder.

        Args:
            source_path: Path to source .vsdx file
            target_template_path: Path to target .vsdx file with desired masters
            mapping: Dictionary mapping shape IDs to new master names
        """
        self.source_path = Path(source_path)
        self.target_template_path = Path(target_template_path)
        self.mapping = mapping

    def rebuild(self, output_path: str) -> None:
        """Rebuild the diagram with new masters.

        Args:
            output_path: Path to save the rebuilt .vsdx file
        """
        from .master_injector import MasterInjector

        print(f"Injecting masters from {self.target_template_path}...")
        injector = MasterInjector(self.source_path, self.target_template_path)
        master_name_map = injector.inject(output_path)
        print(f"Injected {len(master_name_map)} masters.")

        self._apply_mappings_and_style(Path(output_path), master_name_map)

    def _apply_mappings_and_style(
        self,
        output_path: Path,
        master_name_map: Dict[str, str],
    ) -> None:
        ET.register_namespace("", NS["v"])
        ET.register_namespace("r", NS["r"])
        work_dir = Path(tempfile.mkdtemp())
        try:
            with zipfile.ZipFile(output_path, "r") as z:
                z.extractall(work_dir)

            self._apply_template_theme(work_dir)
            self._apply_template_pagesheet(work_dir)
            self._apply_template_document_styles(work_dir)
            template_styles = self._load_template_style_overrides()
            template_text_prototypes = self._load_template_text_prototypes()
            scale_x, scale_y = self._get_page_scale_factors()
            self._apply_mappings(
                work_dir,
                master_name_map,
                template_styles,
                template_text_prototypes,
                scale_x,
                scale_y,
            )
            self._update_page_master_relationships(work_dir)

            self._write_vsdx(work_dir, output_path)
        finally:
            shutil.rmtree(work_dir, ignore_errors=True)

    def _load_source_master_dimensions(self, work_dir: Path) -> Dict[str, Dict[str, str]]:
        """Load Width/Height from source masters for dimension preservation."""
        masters_dir = work_dir / "visio/masters"
        dimensions: Dict[str, Dict[str, str]] = {}
        
        if not masters_dir.exists():
            return dimensions
        
        # Load master ID to file mapping from relationships
        masters_rels = masters_dir / "_rels/masters.xml.rels"
        master_files = {}
        if masters_rels.exists():
            rels_tree = ET.parse(masters_rels)
            for rel in rels_tree.getroot():
                rid = rel.get("Id")
                target = rel.get("Target")
                if rid and target:
                    master_files[rid] = target
        
        # Load master names and their relationship IDs
        masters_xml = masters_dir / "masters.xml"
        if not masters_xml.exists():
            return dimensions
        
        masters_tree = ET.parse(masters_xml)
        for master in masters_tree.getroot().findall(".//v:Master", NS):
            master_id = master.get("ID")
            rel_elem = master.find(".//v:Rel", NS)
            if rel_elem is None:
                continue
            rid = rel_elem.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            if not rid or rid not in master_files:
                continue
            
            master_file = masters_dir / master_files[rid]
            if not master_file.exists():
                continue
            
            try:
                master_tree = ET.parse(master_file)
                shape = master_tree.getroot().find(".//v:Shape", NS)
                if shape:
                    width_cell = shape.find('v:Cell[@N="Width"]', NS)
                    height_cell = shape.find('v:Cell[@N="Height"]', NS)
                    dimensions[master_id] = {
                        "Width": width_cell.get("V") if width_cell is not None else None,
                        "Height": height_cell.get("V") if height_cell is not None else None,
                    }
            except ET.ParseError:
                continue
        
        return dimensions

    def _preserve_shape_dimensions(
        self,
        shape: ET.Element,
        source_master_dims: Dict[str, Dict[str, str]],
    ) -> None:
        """Preserve Width/Height if they're inherited from the old master."""
        old_master_id = shape.get("Master")
        if not old_master_id:
            return
        
        master_dims = source_master_dims.get(old_master_id, {})
        
        # Check Width - add explicit value if inherited
        width_cell = shape.find('v:Cell[@N="Width"]', NS)
        if width_cell is None and master_dims.get("Width"):
            new_cell = ET.Element("{%s}Cell" % NS["v"])
            new_cell.set("N", "Width")
            new_cell.set("V", master_dims["Width"])
            shape.insert(0, new_cell)
        
        # Check Height - add explicit value if inherited
        height_cell = shape.find('v:Cell[@N="Height"]', NS)
        if height_cell is None and master_dims.get("Height"):
            new_cell = ET.Element("{%s}Cell" % NS["v"])
            new_cell.set("N", "Height")
            new_cell.set("V", master_dims["Height"])
            shape.insert(0, new_cell)

    def _apply_mappings(
        self,
        work_dir: Path,
        master_name_map: Dict[str, str],
        template_styles: Dict[str, Dict[str, ET.Element]],
        template_text_prototypes: Dict[str, ET.Element],
        scale_x: float,
        scale_y: float,
    ) -> None:
        pages_dir = work_dir / "visio/pages"
        if not pages_dir.exists():
            return

        mapping_by_id = {str(k): v for k, v in self.mapping.items()}

        master_lookup = {}
        for name, master_id in master_name_map.items():
            master_lookup[name] = master_id
            master_lookup[_normalize_name(name)] = master_id

        # Load source master dimensions for preservation
        source_master_dims = self._load_source_master_dimensions(work_dir)
        template_swimlane_metrics = self._load_template_swimlane_metrics()
        source_master_names = self._load_master_name_map(work_dir)

        for page_path in sorted(pages_dir.glob("page*.xml")):
            if page_path.name == "pages.xml":
                continue

            tree = ET.parse(page_path)
            root = tree.getroot()

            if scale_x != 1.0 or scale_y != 1.0:
                self._scale_page_shapes(root, scale_x, scale_y)

            swimlane_spec = self._build_swimlane_spec(root, template_text_prototypes)
            if swimlane_spec:
                swimlane_spec["flow_rows"] = self._extract_flow_rows(root, source_master_names)
                swimlane_spec["master_names"] = source_master_names
                lane_master_name = self._select_master_name(
                    master_lookup,
                    ("Swimlane (vertical)", "Swimlane"),
                )
                if lane_master_name:
                    container_master_name = self._select_master_name(
                        master_lookup,
                        ("CFF Container", "Swimlane List"),
                    )
                    swimlane_spec["lane_master"] = lane_master_name
                    swimlane_spec["container_master"] = container_master_name
                    if container_master_name is None:
                        swimlane_spec["container_shape"] = None
                    swimlane_spec["heading_proto"] = template_text_prototypes.get("heading")
                    swimlane_spec["title_proto"] = template_text_prototypes.get("title")
                    self._apply_swimlane_layout(root, swimlane_spec, template_swimlane_metrics)
                    lane_ids = []
                    for shape in swimlane_spec["lane_shapes"]:
                        shape_id = shape.get("ID")
                        if shape_id:
                            mapping_by_id[shape_id] = lane_master_name
                            lane_ids.append(shape_id)
                    container_shape = swimlane_spec.get("container_shape")
                    if container_shape is not None and container_master_name:
                        container_id = container_shape.get("ID")
                        if container_id:
                            mapping_by_id[container_id] = container_master_name
                            self._reorder_shapes(root, [container_id] + lane_ids)
                    else:
                        self._reorder_shapes(root, lane_ids)

            for shape in root.findall(".//v:Shape", NS):
                shape_id = shape.get("ID")
                if not shape_id:
                    continue

                target_master = mapping_by_id.get(shape_id)
                if not target_master:
                    continue

                master_id = self._resolve_master_id(target_master, master_lookup)
                if not master_id:
                    print(f"Warning: Master '{target_master}' not found in injected masters")
                    continue

                # Preserve dimensions before changing master
                self._preserve_shape_dimensions(shape, source_master_dims)
                
                shape.set("Master", str(master_id))
                self._strip_style_overrides(shape)
                self._apply_template_style(shape, target_master, template_styles)

            self._transfer_connector_labels(root)

            _write_xml_with_namespaces(tree, page_path)

    def _scale_page_shapes(
        self,
        root: ET.Element,
        scale_x: float,
        scale_y: float,
    ) -> None:
        for shape in root.findall(".//v:Shape", NS):
            for cell in shape.findall("v:Cell", NS):
                name = cell.get("N")
                if not name:
                    continue
                if name in X_SCALE_CELLS:
                    value = _safe_float(cell.get("V"))
                    if value is None:
                        continue
                    cell.set("V", f"{value * scale_x}")
                elif name in Y_SCALE_CELLS:
                    value = _safe_float(cell.get("V"))
                    if value is None:
                        continue
                    cell.set("V", f"{value * scale_y}")

    def _select_master_name(
        self,
        master_lookup: Dict[str, str],
        candidates: Tuple[str, ...],
    ) -> Optional[str]:
        for name in candidates:
            if name in master_lookup or _normalize_name(name) in master_lookup:
                return name
        return None

    def _build_swimlane_spec(
        self,
        root: ET.Element,
        template_text_prototypes: Dict[str, ET.Element],
    ) -> Optional[Dict[str, object]]:
        header_shapes = self._find_header_shapes(root)
        if len(header_shapes) < 2:
            return None

        header_shapes.sort(key=lambda s: _safe_float(_get_cell_value(s, "PinX")) or 0.0)
        header_heights = [_safe_float(_get_cell_value(s, "Height")) for s in header_shapes]
        header_height = _median(header_heights)
        heading_proto = template_text_prototypes.get("heading")
        if heading_proto is not None:
            proto_height = _safe_float(_get_cell_value(heading_proto, "Height"))
            if proto_height:
                header_height = proto_height
        if header_height is None:
            return None

        header_ids = {shape.get("ID") for shape in header_shapes if shape.get("ID")}
        title_shape = self._find_title_shape(root, header_ids)
        title_text = _get_shape_text(title_shape) if title_shape is not None else ""
        title_height = _safe_float(_get_cell_value(title_shape, "Height")) if title_shape is not None else None
        title_proto = template_text_prototypes.get("title")
        if title_proto is not None:
            proto_height = _safe_float(_get_cell_value(title_proto, "Height"))
            if proto_height:
                title_height = proto_height

        all_shapes = [
            shape
            for shape in root.findall(".//v:Shape", NS)
            if not self._is_connector_shape(shape)
            and shape not in header_shapes
        ]
        content_shapes = [
            shape
            for shape in all_shapes
            if title_shape is None or shape is not title_shape
        ]
        bounds_all = self._compute_bounds(all_shapes)
        content_bounds = self._compute_bounds(content_shapes)
        header_bounds = self._compute_bounds(header_shapes)
        if not bounds_all:
            return None

        if header_bounds:
            minx, maxx = header_bounds[0], header_bounds[1]
        else:
            minx, maxx = bounds_all[0], bounds_all[1]
        miny = bounds_all[2]
        maxy = bounds_all[3]

        if title_shape is not None:
            title_bounds = _shape_bounds(title_shape)
            if title_bounds:
                maxy = max(maxy, title_bounds[3])

        if maxx <= minx or maxy <= miny:
            return None

        lane_labels = [_get_shape_text(shape) for shape in header_shapes]

        return {
            "lane_shapes": header_shapes,
            "container_shape": title_shape,
            "bounds": (minx, maxx, miny, maxy),
            "content_bounds": content_bounds,
            "header_height": header_height,
            "title_height": title_height,
            "title_text": title_text,
            "lane_guid": f"{{{uuid.uuid4()}}}",
            "lane_labels": lane_labels,
        }

    def _find_header_shapes(self, root: ET.Element) -> list:
        candidates = []
        for shape in root.findall(".//v:Shape", NS):
            if self._is_connector_shape(shape):
                continue
            if not _get_shape_text(shape):
                continue
            if (
                _get_cell_value(shape, "QuickStyleLineColor") == "101"
                and _get_cell_value(shape, "QuickStyleFillColor") == "101"
            ):
                candidates.append(shape)

        if not candidates:
            return []

        pinys = [
            _safe_float(_get_cell_value(shape, "PinY"))
            for shape in candidates
            if _get_cell_value(shape, "PinY") is not None
        ]
        if not pinys:
            return candidates
        max_piny = max(pinys)
        header_height = _median(
            [
                _safe_float(_get_cell_value(shape, "Height"))
                for shape in candidates
                if _get_cell_value(shape, "Height") is not None
            ]
        )
        if header_height:
            threshold = max_piny - header_height * 0.6
            candidates = [
                shape
                for shape in candidates
                if (_safe_float(_get_cell_value(shape, "PinY")) or 0.0) >= threshold
            ]
        return candidates

    def _find_title_shape(
        self,
        root: ET.Element,
        header_ids: Set[str],
    ) -> Optional[ET.Element]:
        candidates = []
        pinys = []
        for shape in root.findall(".//v:Shape", NS):
            if self._is_connector_shape(shape):
                continue
            shape_id = shape.get("ID")
            if shape_id in header_ids:
                continue
            text = _get_shape_text(shape)
            if not text:
                continue
            width = _safe_float(_get_cell_value(shape, "Width"))
            piny = _safe_float(_get_cell_value(shape, "PinY"))
            if width is None or piny is None:
                continue
            candidates.append((width, piny, shape))
            pinys.append(piny)

        if not candidates:
            return None

        max_piny = max(pinys)
        min_piny = min(pinys)
        if max_piny == min_piny:
            threshold = max_piny
        else:
            threshold = max_piny - (max_piny - min_piny) * 0.2
        top_candidates = [item for item in candidates if item[1] >= threshold]
        if not top_candidates:
            top_candidates = candidates
        return max(top_candidates, key=lambda item: item[0])[2]

    def _compute_bounds(
        self,
        shapes: list,
    ) -> Optional[Tuple[float, float, float, float]]:
        bounds = [None, None, None, None]
        for shape in shapes:
            shape_bounds = _shape_bounds(shape)
            if not shape_bounds:
                continue
            minx, maxx, miny, maxy = shape_bounds
            if bounds[0] is None or minx < bounds[0]:
                bounds[0] = minx
            if bounds[1] is None or maxx > bounds[1]:
                bounds[1] = maxx
            if bounds[2] is None or miny < bounds[2]:
                bounds[2] = miny
            if bounds[3] is None or maxy > bounds[3]:
                bounds[3] = maxy
        if bounds[0] is None:
            return None
        return (bounds[0], bounds[1], bounds[2], bounds[3])

    def _apply_swimlane_layout(
        self,
        root: ET.Element,
        spec: Dict[str, object],
        template_metrics: Optional[Dict[str, object]] = None,
    ) -> None:
        lane_shapes = spec.get("lane_shapes") or []
        if not lane_shapes:
            return

        lane_shapes.sort(key=lambda s: _safe_float(_get_cell_value(s, "PinX")) or 0.0)
        lane_labels = spec.get("lane_labels") or []
        lane_guid = spec.get("lane_guid")

        input_lane_widths = []
        lane_bounds_in = []
        for shape in lane_shapes:
            width = _safe_float(_get_cell_value(shape, "Width"))
            pinx = _safe_float(_get_cell_value(shape, "PinX"))
            if width is None or pinx is None:
                input_lane_widths.append(None)
                lane_bounds_in.append((0.0, 0.0))
                continue
            input_lane_widths.append(width)
            lane_bounds_in.append((pinx - width / 2, pinx + width / 2))

        if template_metrics and template_metrics.get("lane_widths"):
            template_lane_widths = template_metrics.get("lane_widths") or []
            lane_widths = []
            for index in range(len(lane_shapes)):
                if index < len(template_lane_widths):
                    lane_widths.append(template_lane_widths[index])
                else:
                    fallback_width = input_lane_widths[index] if index < len(input_lane_widths) else None
                    if fallback_width is None and template_lane_widths:
                        fallback_width = template_lane_widths[-1]
                    lane_widths.append(fallback_width or 1.0)

            heading_height = template_metrics.get("heading_height") or spec.get("header_height")
            lane_height = template_metrics.get("lane_height") or spec.get("header_height")
            if heading_height is None or lane_height is None:
                return

            container_left = template_metrics.get("container_left")
            container_bottom = template_metrics.get("container_bottom")
            lane_left_offset = template_metrics.get("lane_left_offset") or 0.0
            if container_left is None or container_bottom is None:
                return

            lane_left = container_left + lane_left_offset
            lane_bounds_out = []
            for width in lane_widths:
                lane_bounds_out.append((lane_left, lane_left + width))
                lane_left += width

            container_right = lane_left
            container_width = container_right - container_left
            container_height = lane_height + heading_height
            container_center_x = (container_left + container_right) / 2
            container_center_y = container_bottom + container_height / 2
            container_top = container_bottom + container_height

            lane_center_y = container_bottom + lane_height / 2

            for index, (shape, lane_width) in enumerate(zip(lane_shapes, lane_widths)):
                left, right = lane_bounds_out[index]
                pinx = (left + right) / 2
                _set_cell_value(shape, "PinX", pinx)
                _set_cell_value(shape, "Width", lane_width)
                _set_cell_value(shape, "PinY", lane_center_y)
                _set_cell_value(shape, "Height", lane_height)

                text = lane_labels[index] if index < len(lane_labels) else _get_shape_text(shape)
                if text:
                    self._set_user_cell(shape, "visHeadingText", text)
                    _remove_shape_text(shape)
                if lane_guid:
                    self._set_user_cell(shape, "SwimlaneListGUID", lane_guid)

            container_shape = spec.get("container_shape")
            title_text = spec.get("title_text") or ""
            title_height = spec.get("title_height") or heading_height
            if container_shape is not None:
                _set_cell_value(container_shape, "PinX", container_center_x)
                _set_cell_value(container_shape, "PinY", container_center_y)
                _set_cell_value(container_shape, "Width", container_width)
                _set_cell_value(container_shape, "Height", container_height)
                self._set_user_cell(container_shape, "numLanes", str(len(lane_shapes)))
                self._set_user_cell(container_shape, "CFFVertical", "1")
                self._set_user_cell(container_shape, "msvSDContainerLocked", "1")
                self._set_user_cell(container_shape, "visHeadingHeight", str(heading_height))
                self._set_user_cell(container_shape, "visShowTitle", "1")
                if title_text:
                    _remove_shape_text(container_shape)

            content_bounds = spec.get("content_bounds")
            skip_ids = {
                shape.get("ID")
                for shape in lane_shapes
                if shape.get("ID")
            }
            if container_shape is not None and container_shape.get("ID"):
                skip_ids.add(container_shape.get("ID"))
            if content_bounds:
                row_centers = None
                row_map = None
                decision_target = None
                flow_shape_ids: Optional[Set[str]] = None
                decision_shape_ids: Optional[Set[str]] = None
                master_names = spec.get("master_names") or {}
                if template_metrics and master_names:
                    row_info = self._build_flow_row_map(root, master_names, template_metrics)
                    if row_info:
                        row_centers, row_map, _, decision_target = row_info
                        flow_shape_ids, decision_shape_ids = self._collect_flow_shape_ids(
                            root,
                            master_names,
                            template_metrics.get("process_height"),
                        )

                y_transform = None
                if row_map is None:
                    source_rows = spec.get("flow_rows") or {}
                    source_top = source_rows.get("flow_top")
                    source_decision = source_rows.get("decision_row")
                    template_top = template_metrics.get("flow_top") if template_metrics else None
                    template_decision = template_metrics.get("decision_row") if template_metrics else None
                    if source_top is not None and source_decision is not None and template_top is not None and template_decision is not None:
                        if source_top != source_decision:
                            scale = (template_decision - template_top) / (source_decision - source_top)
                            if scale > 0:
                                offset = template_top - (source_top * scale)
                                y_transform = (scale, offset)
                self._remap_swimlane_content(
                    root,
                    content_bounds,
                    lane_bounds_in,
                    lane_bounds_out,
                    (container_bottom, container_bottom + lane_height),
                    skip_ids,
                    row_centers=row_centers,
                    row_map=row_map,
                    flow_shape_ids=flow_shape_ids,
                    decision_shape_ids=decision_shape_ids,
                    decision_target=decision_target,
                    y_transform=y_transform,
                )

            heading_proto = spec.get("heading_proto")
            if heading_proto is not None and lane_labels:
                header_center_y = container_bottom + lane_height + (heading_height / 2)
                self._append_heading_shapes(
                    root,
                    heading_proto,
                    lane_labels,
                    lane_widths,
                    lane_bounds_out[0][0],
                    header_center_y,
                    heading_height,
                )

            title_proto = spec.get("title_proto")
            if title_proto is not None and title_text:
                title_center_y = container_top - (title_height / 2)
                self._append_title_shape(
                    root,
                    title_proto,
                    title_text,
                    container_center_x,
                    title_center_y,
                    container_width,
                    title_height,
                )

            return

        bounds = spec.get("bounds")
        header_height = spec.get("header_height")
        if not bounds or not header_height:
            return

        minx, maxx, miny, maxy = bounds
        container_width = maxx - minx
        container_height = maxy - miny
        if container_width <= 0 or container_height <= 0:
            return

        title_height = spec.get("title_height") or header_height
        title_text = spec.get("title_text") or ""

        container_center_x = (minx + maxx) / 2
        container_center_y = (miny + maxy) / 2
        container_top = maxy
        container_bottom = miny

        heading_height = header_height
        heading_proto = spec.get("heading_proto")
        if heading_proto is not None:
            proto_height = _safe_float(_get_cell_value(heading_proto, "Height"))
            if proto_height:
                heading_height = proto_height

        lane_height = max(container_height - title_height, heading_height)
        lane_center_y = container_bottom + (lane_height / 2)

        lane_widths = []
        for shape in lane_shapes:
            width = _safe_float(_get_cell_value(shape, "Width"))
            lane_widths.append(width)
        if all(width is None for width in lane_widths):
            lane_widths = [container_width / len(lane_shapes)] * len(lane_shapes)
        else:
            default_width = container_width / len(lane_shapes)
            lane_widths = [
                width if width is not None else default_width for width in lane_widths
            ]
            total_width = sum(lane_widths)
            if total_width > 0:
                scale = container_width / total_width
                lane_widths = [width * scale for width in lane_widths]

        current_x = minx
        for index, (shape, lane_width) in enumerate(zip(lane_shapes, lane_widths)):
            pinx = current_x + lane_width / 2
            _set_cell_value(shape, "PinX", pinx)
            _set_cell_value(shape, "Width", lane_width)
            _set_cell_value(shape, "PinY", lane_center_y)
            _set_cell_value(shape, "Height", lane_height)

            text = lane_labels[index] if index < len(lane_labels) else _get_shape_text(shape)
            if text:
                self._set_user_cell(shape, "visHeadingText", text)
                _remove_shape_text(shape)
            if lane_guid:
                self._set_user_cell(shape, "SwimlaneListGUID", lane_guid)

            current_x += lane_width

        container_shape = spec.get("container_shape")
        if container_shape is not None:
            _set_cell_value(container_shape, "PinX", container_center_x)
            _set_cell_value(container_shape, "PinY", container_center_y)
            _set_cell_value(container_shape, "Width", container_width)
            _set_cell_value(container_shape, "Height", container_height)
            self._set_user_cell(container_shape, "numLanes", str(len(lane_shapes)))
            self._set_user_cell(container_shape, "CFFVertical", "1")
            self._set_user_cell(container_shape, "msvSDContainerLocked", "1")
            self._set_user_cell(container_shape, "visHeadingHeight", str(title_height))
            self._set_user_cell(container_shape, "visShowTitle", "1")
            if title_text:
                _remove_shape_text(container_shape)

        heading_proto = spec.get("heading_proto")
        if heading_proto is not None and lane_labels:
            header_center_y = container_top - title_height - (heading_height / 2)
            self._append_heading_shapes(
                root,
                heading_proto,
                lane_labels,
                lane_widths,
                minx,
                header_center_y,
                heading_height,
            )

        title_proto = spec.get("title_proto")
        if title_proto is not None and title_text:
            title_center_y = container_top - (title_height / 2)
            self._append_title_shape(
                root,
                title_proto,
                title_text,
                container_center_x,
                title_center_y,
                container_width,
                title_height,
            )

    def _find_lane_index(
        self,
        x: Optional[float],
        lane_bounds: list,
    ) -> Optional[int]:
        if x is None:
            return None
        for index, (left, right) in enumerate(lane_bounds):
            if left <= x <= right:
                return index
        return None

    def _remap_swimlane_content(
        self,
        root: ET.Element,
        source_bounds: Tuple[float, float, float, float],
        lane_bounds_in: list,
        lane_bounds_out: list,
        target_y_bounds: Optional[Tuple[float, float]],
        skip_shape_ids: Set[str],
        row_centers: Optional[list] = None,
        row_map: Optional[Dict[float, float]] = None,
        flow_shape_ids: Optional[Set[str]] = None,
        decision_shape_ids: Optional[Set[str]] = None,
        decision_target: Optional[float] = None,
        y_transform: Optional[Tuple[float, float]] = None,
    ) -> None:
        if not source_bounds or not lane_bounds_in or not lane_bounds_out:
            return

        src_minx, src_maxx, src_miny, src_maxy = source_bounds
        if src_maxx <= src_minx or src_maxy <= src_miny:
            return
        if target_y_bounds is not None:
            tgt_min_y, tgt_max_y = target_y_bounds
            if tgt_max_y <= tgt_min_y:
                return

        global_src_minx = lane_bounds_in[0][0]
        global_src_maxx = lane_bounds_in[-1][1]
        global_tgt_minx = lane_bounds_out[0][0]
        global_tgt_maxx = lane_bounds_out[-1][1]

        x_scale_global = 1.0
        x_offset_global = 0.0
        if global_src_maxx > global_src_minx:
            x_scale_global = (global_tgt_maxx - global_tgt_minx) / (global_src_maxx - global_src_minx)
            x_offset_global = global_tgt_minx - global_src_minx * x_scale_global

        def map_x(value: Optional[float]) -> Optional[float]:
            if value is None:
                return None
            lane_index = self._find_lane_index(value, lane_bounds_in)
            if lane_index is None:
                return value * x_scale_global + x_offset_global
            in_left, in_right = lane_bounds_in[lane_index]
            out_left, out_right = lane_bounds_out[lane_index]
            if in_right <= in_left:
                return value * x_scale_global + x_offset_global
            ratio = (value - in_left) / (in_right - in_left)
            return out_left + ratio * (out_right - out_left)

        def map_y(value: Optional[float], shape_id: Optional[str] = None) -> Optional[float]:
            if value is None:
                return None
            if row_centers and row_map:
                if decision_target is not None and decision_shape_ids and shape_id in decision_shape_ids:
                    return decision_target
                if flow_shape_ids and shape_id in flow_shape_ids:
                    nearest = min(row_centers, key=lambda center: abs(value - center))
                    target_center = row_map.get(nearest)
                    if target_center is not None:
                        return target_center
                nearest = min(row_centers, key=lambda center: abs(value - center))
                target_center = row_map.get(nearest)
                if target_center is not None:
                    return target_center + (value - nearest)
            if y_transform is not None:
                scale, offset = y_transform
                return value * scale + offset
            if target_y_bounds is None:
                return value
            y_scale = (tgt_max_y - tgt_min_y) / (src_maxy - src_miny)
            y_offset = tgt_min_y - src_miny * y_scale
            return value * y_scale + y_offset

        for shape in root.findall(".//v:Shape", NS):
            shape_id = shape.get("ID")
            if shape_id in skip_shape_ids:
                continue

            if self._is_connector_shape(shape):
                begin_x = map_x(_safe_float(_get_cell_value(shape, "BeginX")))
                begin_y = map_y(_safe_float(_get_cell_value(shape, "BeginY")), shape_id)
                end_x = map_x(_safe_float(_get_cell_value(shape, "EndX")))
                end_y = map_y(_safe_float(_get_cell_value(shape, "EndY")), shape_id)
                if begin_x is not None:
                    _set_cell_value(shape, "BeginX", begin_x)
                if begin_y is not None:
                    _set_cell_value(shape, "BeginY", begin_y)
                if end_x is not None:
                    _set_cell_value(shape, "EndX", end_x)
                if end_y is not None:
                    _set_cell_value(shape, "EndY", end_y)
                if begin_x is not None and end_x is not None:
                    _set_cell_value(shape, "PinX", (begin_x + end_x) / 2)
                if begin_y is not None and end_y is not None:
                    _set_cell_value(shape, "PinY", (begin_y + end_y) / 2)
                continue

            pinx = map_x(_safe_float(_get_cell_value(shape, "PinX")))
            piny = map_y(_safe_float(_get_cell_value(shape, "PinY")), shape_id)
            if pinx is not None:
                _set_cell_value(shape, "PinX", pinx)
            if piny is not None:
                _set_cell_value(shape, "PinY", piny)

    def _transfer_connector_labels(self, root: ET.Element) -> None:
        label_candidates = []
        connectors = []
        shapes_root = root.find("v:Shapes", NS)
        if shapes_root is None:
            return

        connector_labels = {
            _normalize_label_text("通过"),
            _normalize_label_text("未通过"),
            _normalize_label_text("合同金额超过xx元"),
            _normalize_label_text("合同金额不超过xx元"),
        }

        for shape in list(shapes_root.findall("v:Shape", NS)):
            if self._is_connector_shape(shape):
                connectors.append(shape)
                continue

            text = _get_shape_text(shape)
            if not text:
                continue

            normalized = _normalize_label_text(text)
            if normalized not in connector_labels:
                continue

            width = _safe_float(_get_cell_value(shape, "Width")) or 0.0
            height = _safe_float(_get_cell_value(shape, "Height")) or 0.0
            if width > 1.2 or height > 0.8:
                continue

            pinx = _safe_float(_get_cell_value(shape, "PinX"))
            piny = _safe_float(_get_cell_value(shape, "PinY"))
            if pinx is None or piny is None:
                continue
            label_candidates.append((shape, text, pinx, piny))

        if not label_candidates or not connectors:
            return

        available_connectors = []
        for connector in connectors:
            if _get_shape_text(connector):
                continue
            begin_x = _safe_float(_get_cell_value(connector, "BeginX"))
            begin_y = _safe_float(_get_cell_value(connector, "BeginY"))
            end_x = _safe_float(_get_cell_value(connector, "EndX"))
            end_y = _safe_float(_get_cell_value(connector, "EndY"))
            if begin_x is None or begin_y is None or end_x is None or end_y is None:
                continue
            mid_x = (begin_x + end_x) / 2
            mid_y = (begin_y + end_y) / 2
            available_connectors.append((connector, mid_x, mid_y))

        if not available_connectors:
            return

        for label_shape, label_text, label_x, label_y in label_candidates:
            best_connector = None
            best_distance = None
            for connector, mid_x, mid_y in available_connectors:
                dx = label_x - mid_x
                dy = label_y - mid_y
                distance = (dx * dx + dy * dy) ** 0.5
                if best_distance is None or distance < best_distance:
                    best_distance = distance
                    best_connector = connector
            if best_connector is None or best_distance is None or best_distance > 2.0:
                continue
            _set_shape_text(best_connector, label_text)
            try:
                shapes_root.remove(label_shape)
            except ValueError:
                continue

    def _set_user_cell(self, shape: ET.Element, row_name: str, value: str) -> None:
        user_section = shape.find("v:Section[@N='User']", NS)
        if user_section is None:
            user_section = ET.SubElement(shape, "{%s}Section" % NS["v"])
            user_section.set("N", "User")

        row = user_section.find(f"v:Row[@N='{row_name}']", NS)
        if row is None:
            row = ET.SubElement(user_section, "{%s}Row" % NS["v"])
            row.set("N", row_name)

        cell = row.find("v:Cell[@N='Value']", NS)
        if cell is None:
            cell = ET.SubElement(row, "{%s}Cell" % NS["v"])
        cell.set("N", "Value")
        cell.set("V", str(value))

    def _append_heading_shapes(
        self,
        root: ET.Element,
        prototype: ET.Element,
        labels: list,
        lane_widths: list,
        start_x: float,
        center_y: float,
        height: float,
    ) -> None:
        shapes_root = root.find("v:Shapes", NS)
        if shapes_root is None:
            return

        next_id = self._next_shape_id(shapes_root)
        current_x = start_x
        for label, lane_width in zip(labels, lane_widths):
            new_shape = copy.deepcopy(prototype)
            new_shape.set("ID", str(next_id))
            next_id += 1

            self._remove_section(new_shape, "User")

            _set_shape_text(new_shape, label)
            _set_cell_value(new_shape, "PinX", current_x + lane_width / 2)
            _set_cell_value(new_shape, "PinY", center_y)
            _set_cell_value(new_shape, "Width", lane_width)
            _set_cell_value(new_shape, "Height", height)
            shapes_root.append(new_shape)
            current_x += lane_width

    def _next_shape_id(self, shapes_root: ET.Element) -> int:
        max_id = 0
        for shape in shapes_root.findall("v:Shape", NS):
            shape_id = shape.get("ID")
            if not shape_id:
                continue
            try:
                max_id = max(max_id, int(shape_id))
            except ValueError:
                continue
        return max_id + 1

    def _append_title_shape(
        self,
        root: ET.Element,
        prototype: ET.Element,
        title_text: str,
        center_x: float,
        center_y: float,
        width: float,
        height: float,
    ) -> None:
        shapes_root = root.find("v:Shapes", NS)
        if shapes_root is None:
            return

        new_shape = copy.deepcopy(prototype)
        new_shape.set("ID", str(self._next_shape_id(shapes_root)))
        _set_shape_text(new_shape, title_text)
        _set_cell_value(new_shape, "PinX", center_x)
        _set_cell_value(new_shape, "PinY", center_y)
        _set_cell_value(new_shape, "Width", width)
        _set_cell_value(new_shape, "Height", height)
        shapes_root.append(new_shape)

    def _remove_section(self, shape: ET.Element, section_name: str) -> None:
        for section in list(shape.findall("v:Section", NS)):
            if section.get("N") == section_name:
                shape.remove(section)

    def _reorder_shapes(self, root: ET.Element, shape_ids: list) -> None:
        if not shape_ids:
            return
        shapes_root = root.find("v:Shapes", NS)
        if shapes_root is None:
            return

        shape_map = {}
        for shape in list(shapes_root.findall("v:Shape", NS)):
            shape_id = shape.get("ID")
            if shape_id in shape_ids:
                shape_map[shape_id] = shape
                shapes_root.remove(shape)

        insert_at = 0
        for shape_id in shape_ids:
            shape = shape_map.get(shape_id)
            if shape is None:
                continue
            shapes_root.insert(insert_at, shape)
            insert_at += 1

    def _resolve_master_id(
        self,
        target_master: str,
        master_lookup: Dict[str, str],
    ) -> Optional[str]:
        if target_master in master_lookup:
            return master_lookup[target_master]
        normalized = _normalize_name(target_master)
        return master_lookup.get(normalized)

    def _strip_style_overrides(self, shape: ET.Element) -> None:
        for child in list(shape):
            if child.tag.endswith("Section"):
                if child.get("N") in STYLE_SECTIONS_TO_REMOVE:
                    shape.remove(child)

        for cell in list(shape.findall("v:Cell", NS)):
            if cell.get("N") in STYLE_CELLS_TO_REMOVE:
                shape.remove(cell)

    def _apply_template_style(
        self,
        shape: ET.Element,
        target_master: str,
        template_styles: Dict[str, Dict[str, ET.Element]],
    ) -> None:
        style_cells = template_styles.get(target_master)
        if not style_cells:
            return

        # Remove any existing style cells we plan to replace.
        for cell in list(shape.findall("v:Cell", NS)):
            if cell.get("N") in STYLE_CELLS_TO_COPY:
                shape.remove(cell)

        for cell_name, cell in style_cells.items():
            shape.append(copy.deepcopy(cell))

    def _load_template_style_overrides(self) -> Dict[str, Dict[str, ET.Element]]:
        with zipfile.ZipFile(self.target_template_path, "r") as z:
            try:
                page_xml = z.read("visio/pages/page1.xml")
                masters_xml = z.read("visio/masters/masters.xml")
            except KeyError:
                return {}

        page_root = ET.fromstring(page_xml)
        masters_root = ET.fromstring(masters_xml)

        master_names = {}
        for master in masters_root.findall(".//v:Master", NS):
            master_id = master.get("ID")
            name = master.get("NameU") or master.get("Name")
            if master_id and name:
                master_names[master_id] = name

        styles_by_master: Dict[str, Dict[str, ET.Element]] = {}

        for shape in page_root.findall(".//v:Shape", NS):
            master_id = shape.get("Master")
            if not master_id:
                continue
            master_name = master_names.get(master_id)
            if not master_name or master_name in styles_by_master:
                continue

            style_cells: Dict[str, ET.Element] = {}
            for cell in shape.findall("v:Cell", NS):
                name = cell.get("N")
                if name in STYLE_CELLS_TO_COPY:
                    style_cells[name] = cell

            if style_cells:
                styles_by_master[master_name] = style_cells

        return styles_by_master

    def _load_template_text_prototypes(self) -> Dict[str, ET.Element]:
        with zipfile.ZipFile(self.target_template_path, "r") as z:
            try:
                page_xml = z.read("visio/pages/page1.xml")
            except KeyError:
                return {}

        page_root = ET.fromstring(page_xml)

        text_shapes = []
        for shape in page_root.findall(".//v:Shape", NS):
            if shape.get("Master"):
                continue
            text = _get_shape_text(shape)
            if not text:
                continue
            fill = _get_cell_value(shape, "FillForegnd")
            height = _safe_float(_get_cell_value(shape, "Height"))
            width = _safe_float(_get_cell_value(shape, "Width")) or 0.0
            if fill == "#f2f2f2" and height:
                text_shapes.append((height, width, shape))

        if not text_shapes:
            return {}

        text_shapes.sort(key=lambda item: item[0])
        heading_shape = text_shapes[0][2]
        title_shape = max(text_shapes, key=lambda item: item[1])[2]

        return {
            "heading": copy.deepcopy(heading_shape),
            "title": copy.deepcopy(title_shape),
        }

    def _get_user_cell_value(self, shape: ET.Element, row_name: str) -> Optional[str]:
        user_section = shape.find("v:Section[@N='User']", NS)
        if user_section is None:
            return None
        row = user_section.find(f"v:Row[@N='{row_name}']", NS)
        if row is None:
            return None
        cell = row.find("v:Cell[@N='Value']", NS)
        return cell.get("V") if cell is not None else None

    def _load_template_swimlane_metrics(self) -> Optional[Dict[str, object]]:
        with zipfile.ZipFile(self.target_template_path, "r") as z:
            try:
                page_xml = z.read("visio/pages/page1.xml")
                masters_xml = z.read("visio/masters/masters.xml")
            except KeyError:
                return None

        page_root = ET.fromstring(page_xml)
        masters_root = ET.fromstring(masters_xml)

        master_names = {}
        for master in masters_root.findall(".//v:Master", NS):
            master_id = master.get("ID")
            name = master.get("NameU") or master.get("Name")
            if master_id and name:
                master_names[master_id] = name

        container_shape = None
        swimlane_list_shape = None
        lane_shapes = []

        for shape in page_root.findall(".//v:Shape", NS):
            master_name = master_names.get(shape.get("Master"))
            if master_name == "CFF Container":
                container_shape = shape
            elif master_name == "Swimlane List":
                swimlane_list_shape = shape
            elif master_name == "Swimlane (vertical)":
                lane_shapes.append(shape)

        if container_shape is None or not lane_shapes:
            return None

        container_width = _safe_float(_get_cell_value(container_shape, "Width"))
        container_height = _safe_float(_get_cell_value(container_shape, "Height"))
        container_pinx = _safe_float(_get_cell_value(container_shape, "PinX"))
        container_piny = _safe_float(_get_cell_value(container_shape, "PinY"))
        if container_width is None or container_height is None or container_pinx is None or container_piny is None:
            return None

        container_left = container_pinx - container_width / 2
        container_bottom = container_piny - container_height / 2

        lane_shapes.sort(key=lambda s: _safe_float(_get_cell_value(s, "PinX")) or 0.0)
        lane_widths = []
        lane_height = None
        lane_left_offset = None
        for shape in lane_shapes:
            width = _safe_float(_get_cell_value(shape, "Width"))
            height = _safe_float(_get_cell_value(shape, "Height"))
            pinx = _safe_float(_get_cell_value(shape, "PinX"))
            if width is None or pinx is None:
                continue
            lane_widths.append(width)
            if lane_height is None and height is not None:
                lane_height = height
            if lane_left_offset is None:
                lane_left_offset = (pinx - width / 2) - container_left

        heading_height = None
        if swimlane_list_shape is not None:
            heading_height = _safe_float(self._get_user_cell_value(swimlane_list_shape, "visHeadingHeight"))

        master_dims = self._load_master_dimensions_from_vsdx(self.target_template_path)
        process_height = None
        start_end_height = None
        for master_id, name in master_names.items():
            if name == "Process" and master_id in master_dims:
                process_height = _safe_float(master_dims[master_id].get("Height"))
            elif name == "Start/End" and master_id in master_dims:
                start_end_height = _safe_float(master_dims[master_id].get("Height"))

        flow_y = []
        decision_y = []
        for shape in page_root.findall(".//v:Shape", NS):
            master_name = master_names.get(shape.get("Master"), "")
            if master_name not in {"Process", "Decision", "Start/End"}:
                continue
            pin_y = _safe_float(_get_cell_value(shape, "PinY"))
            if pin_y is None:
                continue
            flow_y.append(pin_y)
            if master_name == "Decision":
                decision_y.append(pin_y)

        tolerance = (process_height or 0.5) * 0.6
        flow_rows = _cluster_row_centers(flow_y, tolerance)

        return {
            "container_left": container_left,
            "container_bottom": container_bottom,
            "lane_left_offset": lane_left_offset or 0.0,
            "lane_widths": lane_widths,
            "lane_height": lane_height,
            "heading_height": heading_height,
            "flow_top": max(flow_y) if flow_y else None,
            "decision_row": _mode(decision_y),
            "flow_rows": flow_rows,
            "process_height": process_height,
            "start_end_height": start_end_height,
        }

    def _load_master_name_map(self, work_dir: Path) -> Dict[str, str]:
        masters_xml_path = work_dir / "visio/masters/masters.xml"
        if not masters_xml_path.exists():
            return {}
        root = ET.parse(masters_xml_path).getroot()
        master_names = {}
        for master in root.findall(".//v:Master", NS):
            master_id = master.get("ID")
            name = master.get("NameU") or master.get("Name")
            if master_id and name:
                master_names[master_id] = name
        return master_names

    def _load_master_dimensions_from_vsdx(self, vsdx_path: Path) -> Dict[str, Dict[str, str]]:
        dimensions: Dict[str, Dict[str, str]] = {}
        try:
            with zipfile.ZipFile(vsdx_path, "r") as z:
                masters_xml = z.read("visio/masters/masters.xml")
                rels_xml = z.read("visio/masters/_rels/masters.xml.rels")
        except (KeyError, FileNotFoundError, zipfile.BadZipFile):
            return dimensions

        masters_root = ET.fromstring(masters_xml)
        rels_root = ET.fromstring(rels_xml)
        rels_map = {rel.get("Id"): rel.get("Target") for rel in rels_root}

        master_targets = {}
        for master in masters_root.findall(".//v:Master", NS):
            master_id = master.get("ID")
            rel = master.find("v:Rel", NS)
            if master_id is None or rel is None:
                continue
            rel_id = rel.get(f"{{{NS['r']}}}id")
            target = rels_map.get(rel_id)
            if target:
                master_targets[master_id] = target

        with zipfile.ZipFile(vsdx_path, "r") as z:
            for master_id, target in master_targets.items():
                try:
                    master_xml = z.read(f"visio/masters/{target}")
                except KeyError:
                    continue
                master_root = ET.fromstring(master_xml)
                shape = master_root.find(".//v:Shape", NS)
                if shape is None:
                    continue
                width_cell = shape.find('v:Cell[@N="Width"]', NS)
                height_cell = shape.find('v:Cell[@N="Height"]', NS)
                dimensions[master_id] = {
                    "Width": width_cell.get("V") if width_cell is not None else None,
                    "Height": height_cell.get("V") if height_cell is not None else None,
                }

        return dimensions

    def _extract_flow_rows(
        self,
        root: ET.Element,
        master_names: Dict[str, str],
    ) -> Dict[str, Optional[float]]:
        flow_y = []
        decision_y = []
        flow_master_names = {
            "process",
            "roundedrectangle",
            "rectangle",
            "decision",
            "diamond",
            "decagon",
            "startend",
            "terminator",
            "start/end",
        }
        decision_master_names = {"decision", "diamond", "decagon"}
        connector_labels = {
            _normalize_label_text("通过"),
            _normalize_label_text("未通过"),
            _normalize_label_text("合同金额超过xx元"),
            _normalize_label_text("合同金额不超过xx元"),
        }

        for shape in root.findall(".//v:Shape", NS):
            if self._is_connector_shape(shape):
                continue
            if (
                _get_cell_value(shape, "QuickStyleLineColor") == "101"
                and _get_cell_value(shape, "QuickStyleFillColor") == "101"
            ):
                continue
            text = _get_shape_text(shape)
            if text and _normalize_label_text(text) in connector_labels:
                continue
            master_name = master_names.get(shape.get("Master"), "")
            master_norm = _normalize_name(master_name)
            if not master_norm:
                if not text:
                    continue
                width = _safe_float(_get_cell_value(shape, "Width"))
                height = _safe_float(_get_cell_value(shape, "Height"))
                if width is not None and height is not None:
                    if width > 2.5 or height > 1.2:
                        continue
                else:
                    if len(text.strip()) > 6:
                        continue
                master_norm = "process"
            if master_norm not in flow_master_names:
                continue
            pin_y = _safe_float(_get_cell_value(shape, "PinY"))
            if pin_y is None:
                continue
            flow_y.append(pin_y)
            if master_norm in decision_master_names:
                decision_y.append(pin_y)

        return {
            "flow_top": max(flow_y) if flow_y else None,
            "decision_row": _mode(decision_y),
        }

    def _build_flow_row_map(
        self,
        root: ET.Element,
        master_names: Dict[str, str],
        template_metrics: Dict[str, object],
    ) -> Optional[Tuple[list, Dict[float, float], float, float]]:
        process_height = template_metrics.get("process_height") or 0.5
        tolerance = process_height * 0.6
        max_flow_height = process_height * 2.5
        template_rows = template_metrics.get("flow_rows") or []
        if not template_rows:
            return None

        flow_values = []
        decision_values = []
        row_members = []
        decision_master_names = {"decision", "diamond", "decagon"}
        flow_master_names = {
            "process",
            "roundedrectangle",
            "rectangle",
            "decision",
            "diamond",
            "decagon",
            "startend",
            "terminator",
            "start/end",
        }
        connector_labels = {
            _normalize_label_text("通过"),
            _normalize_label_text("未通过"),
            _normalize_label_text("合同金额超过xx元"),
            _normalize_label_text("合同金额不超过xx元"),
        }

        for shape in root.findall(".//v:Shape", NS):
            if self._is_connector_shape(shape):
                continue
            if (
                _get_cell_value(shape, "QuickStyleLineColor") == "101"
                and _get_cell_value(shape, "QuickStyleFillColor") == "101"
            ):
                continue
            height = _safe_float(_get_cell_value(shape, "Height"))
            if height is not None and height > max_flow_height:
                continue
            text = _get_shape_text(shape)
            if text and _normalize_label_text(text) in connector_labels:
                continue
            master_name = master_names.get(shape.get("Master"), "")
            master_norm = _normalize_name(master_name)
            if not master_norm:
                if not text:
                    continue
                width = _safe_float(_get_cell_value(shape, "Width"))
                height = _safe_float(_get_cell_value(shape, "Height"))
                if width is not None and height is not None:
                    if width > 2.5 or (height is not None and height > max_flow_height):
                        continue
                else:
                    if len(text.strip()) > 6:
                        continue
                master_norm = "process"
            if master_norm not in flow_master_names:
                continue
            pin_y = _safe_float(_get_cell_value(shape, "PinY"))
            if pin_y is None:
                continue
            flow_values.append(pin_y)
            row_members.append((pin_y, master_norm))
            if master_norm in decision_master_names:
                decision_values.append(pin_y)

        if not flow_values:
            return None

        row_centers = _cluster_row_centers(flow_values, tolerance)
        if not row_centers:
            return None

        decision_center = None
        if decision_values:
            decision_assignments = {}
            for value in decision_values:
                nearest = min(row_centers, key=lambda center: abs(value - center))
                decision_assignments[nearest] = decision_assignments.get(nearest, 0) + 1
            decision_center = max(decision_assignments, key=decision_assignments.get)

        if decision_center is None:
            return None

        row_counts = {center: 0 for center in row_centers}
        row_decision_counts = {center: 0 for center in row_centers}
        for value, master_norm in row_members:
            nearest = min(row_centers, key=lambda center: abs(value - center))
            row_counts[nearest] += 1
            if master_norm in decision_master_names:
                row_decision_counts[nearest] += 1

        if len(decision_values) >= 2:
            outlier_centers = [
                center
                for center in row_centers
                if center > decision_center
                and row_counts.get(center, 0) <= 1
                and row_decision_counts.get(center, 0) == row_counts.get(center, 0)
            ]
            if outlier_centers:
                row_centers = [center for center in row_centers if center not in outlier_centers]

        decision_index = row_centers.index(decision_center)
        target_decision_index = min(decision_index, len(template_rows) - 1)
        target_decision = template_rows[target_decision_index]

        row_map: Dict[float, float] = {}

        for idx in range(decision_index):
            target_row = template_rows[idx] if idx < len(template_rows) else template_rows[-1]
            row_map[row_centers[idx]] = target_row

        row_map[decision_center] = target_decision

        rows_below = len(row_centers) - decision_index - 1
        if rows_below > 0:
            bottom_index = min(
                len(template_rows) - 1,
                target_decision_index + max(2, rows_below + 1),
            )
            bottom_anchor = template_rows[bottom_index]

            bottom_rows = row_centers[decision_index + 1 :]
            bottom_contains_start_end = False
            for value, master_norm in row_members:
                if master_norm in {"startend", "terminator", "start/end"}:
                    nearest = min(bottom_rows, key=lambda center: abs(value - center))
                    if nearest == bottom_rows[-1]:
                        bottom_contains_start_end = True
                        break

            start_end_height = template_metrics.get("start_end_height")
            if bottom_contains_start_end and start_end_height:
                bottom_anchor -= start_end_height / 2

            source_decision = decision_center
            source_bottom = row_centers[-1]
            for src_y in bottom_rows:
                if source_decision == source_bottom:
                    t = 1.0
                else:
                    t = (source_decision - src_y) / (source_decision - source_bottom)
                row_map[src_y] = target_decision - t * (target_decision - bottom_anchor)

        decision_target = row_map.get(decision_center)
        if decision_target is None:
            decision_target = target_decision

        return row_centers, row_map, decision_center, decision_target

    def _collect_flow_shape_ids(
        self,
        root: ET.Element,
        master_names: Dict[str, str],
        process_height: Optional[float],
    ) -> Tuple[Set[str], Set[str]]:
        max_flow_height = (process_height or 0.5) * 2.5
        connector_labels = {
            _normalize_label_text("通过"),
            _normalize_label_text("未通过"),
            _normalize_label_text("合同金额超过xx元"),
            _normalize_label_text("合同金额不超过xx元"),
        }
        flow_master_names = {
            "process",
            "roundedrectangle",
            "rectangle",
            "decision",
            "diamond",
            "decagon",
            "startend",
            "terminator",
            "start/end",
        }
        decision_master_names = {"decision", "diamond", "decagon"}

        flow_shape_ids: Set[str] = set()
        decision_shape_ids: Set[str] = set()

        for shape in root.findall(".//v:Shape", NS):
            shape_id = shape.get("ID")
            if not shape_id:
                continue
            if self._is_connector_shape(shape):
                continue
            if (
                _get_cell_value(shape, "QuickStyleLineColor") == "101"
                and _get_cell_value(shape, "QuickStyleFillColor") == "101"
            ):
                continue
            height = _safe_float(_get_cell_value(shape, "Height"))
            if height is not None and height > max_flow_height:
                continue

            text = _get_shape_text(shape)
            if text and _normalize_label_text(text) in connector_labels:
                continue

            master_name = master_names.get(shape.get("Master"), "")
            master_norm = _normalize_name(master_name)
            if not master_norm:
                if not text:
                    continue
                width = _safe_float(_get_cell_value(shape, "Width"))
                height = _safe_float(_get_cell_value(shape, "Height"))
                if width is not None and height is not None:
                    if width > 2.5 or (height is not None and height > max_flow_height):
                        continue
                else:
                    if len(text.strip()) > 6:
                        continue
                master_norm = "process"
            if master_norm not in flow_master_names:
                continue

            flow_shape_ids.add(shape_id)
            if master_norm in decision_master_names:
                decision_shape_ids.add(shape_id)

        return flow_shape_ids, decision_shape_ids

    def _update_page_master_relationships(self, work_dir: Path) -> None:
        master_targets = self._load_master_targets(work_dir)
        if not master_targets:
            return

        pages_dir = work_dir / "visio/pages"
        rels_dir = pages_dir / "_rels"
        if not pages_dir.exists():
            return
        rels_dir.mkdir(parents=True, exist_ok=True)

        for page_path in sorted(pages_dir.glob("page*.xml")):
            if page_path.name == "pages.xml":
                continue

            used_master_ids = self._collect_page_master_ids(page_path)
            if not used_master_ids:
                continue

            rel_path = rels_dir / f"{page_path.name}.rels"
            rel_root, rel_ids, rel_targets = self._load_relationships(rel_path)

            next_id = self._next_rel_id(rel_ids)
            for master_id in used_master_ids:
                target = master_targets.get(master_id)
                if not target:
                    continue
                rel_target = f"../masters/{target}"
                if rel_target in rel_targets:
                    continue
                rel_id = f"rId{next_id}"
                next_id += 1
                rel = ET.SubElement(rel_root, "Relationship")
                rel.set("Id", rel_id)
                rel.set("Type", MASTER_REL_TYPE)
                rel.set("Target", rel_target)
                rel_targets.add(rel_target)

            self._write_relationships(rel_path, rel_root)

    def _load_master_targets(self, work_dir: Path) -> Dict[str, str]:
        masters_xml_path = work_dir / "visio/masters/masters.xml"
        rels_xml_path = work_dir / "visio/masters/_rels/masters.xml.rels"
        if not masters_xml_path.exists() or not rels_xml_path.exists():
            return {}

        masters_root = ET.parse(masters_xml_path).getroot()
        rels_root = ET.parse(rels_xml_path).getroot()
        rels_map = {rel.get("Id"): rel.get("Target") for rel in rels_root}

        master_targets = {}
        for master in masters_root.findall(".//v:Master", NS):
            master_id = master.get("ID")
            rel = master.find("v:Rel", NS)
            if master_id is None or rel is None:
                continue
            rel_id = rel.get(f"{{{NS['r']}}}id")
            target = rels_map.get(rel_id)
            if target:
                master_targets[master_id] = target

        return master_targets

    def _collect_page_master_ids(self, page_path: Path) -> Set[str]:
        root = ET.parse(page_path).getroot()
        master_ids = set()
        for shape in root.findall(".//v:Shape", NS):
            master_id = shape.get("Master")
            if master_id:
                master_ids.add(master_id)
        return master_ids

    def _load_relationships(self, rel_path: Path):
        if rel_path.exists():
            root = ET.parse(rel_path).getroot()
        else:
            root = ET.Element("Relationships", {"xmlns": REL_NS})

        rel_ids = set()
        rel_targets = set()
        for rel in root:
            rel_id = rel.get("Id")
            rel_target = rel.get("Target")
            if rel_id:
                rel_ids.add(rel_id)
            if rel_target:
                rel_targets.add(rel_target)
        return root, rel_ids, rel_targets

    def _next_rel_id(self, existing_ids: Set[str]) -> int:
        max_id = 0
        for rel_id in existing_ids:
            if rel_id.startswith("rId"):
                try:
                    max_id = max(max_id, int(rel_id[3:]))
                except ValueError:
                    continue
        return max_id + 1

    def _write_relationships(self, rel_path: Path, rel_root: ET.Element) -> None:
        _write_relationships_xml(rel_root, rel_path)

    def _is_connector_shape(self, shape: ET.Element) -> bool:
        for cell in shape.findall("v:Cell", NS):
            if cell.get("N") in {"BeginX", "EndX"}:
                return True
        return False

    def _apply_template_theme(self, work_dir: Path) -> None:
        output_theme_target = self._get_theme_target_from_dir(work_dir)
        template_theme_target = self._get_theme_target_from_template()

        if not output_theme_target or not template_theme_target:
            return

        with zipfile.ZipFile(self.target_template_path, "r") as z:
            template_theme_path = str(Path("visio") / template_theme_target)
            try:
                theme_data = z.read(template_theme_path)
            except KeyError:
                return

        output_theme_path = work_dir / "visio" / output_theme_target
        output_theme_path.parent.mkdir(parents=True, exist_ok=True)
        output_theme_path.write_bytes(theme_data)

    def _get_theme_target_from_dir(self, work_dir: Path) -> Optional[str]:
        rels_path = work_dir / "visio/_rels/document.xml.rels"
        if not rels_path.exists():
            return None

        rels_root = ET.parse(rels_path).getroot()
        for rel in rels_root:
            rel_type = rel.get("Type", "")
            if rel_type.endswith("/theme"):
                return rel.get("Target")
        return None

    def _get_theme_target_from_template(self) -> Optional[str]:
        with zipfile.ZipFile(self.target_template_path, "r") as z:
            try:
                rels_data = z.read("visio/_rels/document.xml.rels")
            except KeyError:
                return None

        rels_root = ET.fromstring(rels_data)
        for rel in rels_root:
            rel_type = rel.get("Type", "")
            if rel_type.endswith("/theme"):
                return rel.get("Target")
        return None

    def _apply_template_pagesheet(self, work_dir: Path) -> None:
        output_pages_path = work_dir / "visio/pages/pages.xml"
        if not output_pages_path.exists():
            return

        with zipfile.ZipFile(self.target_template_path, "r") as z:
            try:
                template_pages_data = z.read("visio/pages/pages.xml")
            except KeyError:
                return

        output_root = ET.parse(output_pages_path).getroot()
        template_root = ET.fromstring(template_pages_data)

        output_page = output_root.find(".//v:Page", NS)
        template_page = template_root.find(".//v:Page", NS)
        if output_page is None or template_page is None:
            return

        for attr in ("ViewScale", "ViewCenterX", "ViewCenterY"):
            if template_page.get(attr) is not None:
                output_page.set(attr, template_page.get(attr))

        template_pagesheet = template_page.find("v:PageSheet", NS)
        if template_pagesheet is None:
            return

        output_pagesheet = output_page.find("v:PageSheet", NS)
        if output_pagesheet is not None:
            output_page.remove(output_pagesheet)
        output_page.append(copy.deepcopy(template_pagesheet))

        _write_xml_with_namespaces(ET.ElementTree(output_root), output_pages_path)

    def _get_page_scale_factors(self) -> Tuple[float, float]:
        source_dims = self._get_page_dimensions(self.source_path)
        template_dims = self._get_page_dimensions(self.target_template_path)
        if not source_dims or not template_dims:
            return 1.0, 1.0

        source_width, source_height = source_dims
        template_width, template_height = template_dims
        if not source_width or not source_height:
            return 1.0, 1.0

        scale_x = template_width / source_width if source_width else 1.0
        scale_y = template_height / source_height if source_height else 1.0
        return scale_x, scale_y

    def _get_page_dimensions(self, vsdx_path: Path) -> Optional[Tuple[float, float]]:
        try:
            with zipfile.ZipFile(vsdx_path, "r") as z:
                pages_xml = z.read("visio/pages/pages.xml")
        except (KeyError, FileNotFoundError, zipfile.BadZipFile):
            return None

        root = ET.fromstring(pages_xml)
        page = root.find(".//v:Page", NS)
        if page is None:
            return None
        pagesheet = page.find("v:PageSheet", NS)
        if pagesheet is None:
            return None

        width = height = None
        for cell in pagesheet.findall(".//v:Cell", NS):
            if cell.get("N") == "PageWidth":
                width = _safe_float(cell.get("V"))
            elif cell.get("N") == "PageHeight":
                height = _safe_float(cell.get("V"))

        if width is None or height is None:
            return None
        return width, height

    def _apply_template_document_styles(self, work_dir: Path) -> None:
        output_doc_path = work_dir / "visio/document.xml"
        if not output_doc_path.exists():
            return

        with zipfile.ZipFile(self.target_template_path, "r") as z:
            try:
                template_doc_data = z.read("visio/document.xml")
            except KeyError:
                return

        output_root = ET.parse(output_doc_path).getroot()
        template_root = ET.fromstring(template_doc_data)

        for tag in ("Colors", "FaceNames", "StyleSheets"):
            output_node = output_root.find(f"v:{tag}", NS)
            template_node = template_root.find(f"v:{tag}", NS)
            if template_node is None:
                continue
            if output_node is not None:
                output_root.remove(output_node)
            output_root.append(copy.deepcopy(template_node))

        _write_xml_with_namespaces(ET.ElementTree(output_root), output_doc_path)

    def _write_vsdx(self, work_dir: Path, output_path: Path) -> None:
        fd, tmp_path = tempfile.mkstemp(suffix=".vsdx")
        os.close(fd)
        Path(tmp_path).unlink(missing_ok=True)

        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as z:
            for path in work_dir.rglob("*"):
                if path.is_file():
                    z.write(path, path.relative_to(work_dir))

        shutil.move(tmp_path, output_path)
