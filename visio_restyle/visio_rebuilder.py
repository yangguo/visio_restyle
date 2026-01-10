"""
Rebuild Visio diagram with new masters while preserving layout and connections.
"""

from typing import Dict, Optional, Set
from pathlib import Path
import copy
import os
import re
import shutil
import tempfile
import zipfile
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
            self._apply_mappings(work_dir, master_name_map, template_styles)
            self._update_page_master_relationships(work_dir)

            self._write_vsdx(work_dir, output_path)
        finally:
            shutil.rmtree(work_dir, ignore_errors=True)

    def _apply_mappings(
        self,
        work_dir: Path,
        master_name_map: Dict[str, str],
        template_styles: Dict[str, Dict[str, ET.Element]],
    ) -> None:
        pages_dir = work_dir / "visio/pages"
        if not pages_dir.exists():
            return

        mapping_by_id = {str(k): v for k, v in self.mapping.items()}

        master_lookup = {}
        for name, master_id in master_name_map.items():
            master_lookup[name] = master_id
            master_lookup[_normalize_name(name)] = master_id

        for page_path in sorted(pages_dir.glob("page*.xml")):
            if page_path.name == "pages.xml":
                continue

            tree = ET.parse(page_path)
            root = tree.getroot()

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

                shape.set("Master", str(master_id))
                self._strip_style_overrides(shape)
                self._apply_template_style(shape, target_master, template_styles)

            _write_xml_with_namespaces(tree, page_path)

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
