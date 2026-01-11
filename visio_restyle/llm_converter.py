"""
LLM-centric Visio diagram converter.
Maximizes LLM usage for analysis, transformation planning, and output generation.
"""

import json
import os
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Any, Optional, Tuple
from pathlib import Path
from dataclasses import dataclass, asdict
from openai import OpenAI


NS = {"v": "http://schemas.microsoft.com/office/visio/2012/main"}


@dataclass
class ShapeInfo:
    """Detailed shape information for LLM analysis."""
    id: str
    master_name: str
    text: str
    x: float
    y: float
    width: float
    height: float
    is_connector: bool
    begin_x: Optional[float] = None
    begin_y: Optional[float] = None
    end_x: Optional[float] = None
    end_y: Optional[float] = None


@dataclass
class TransformSpec:
    """LLM-generated transformation specification for a shape."""
    shape_id: str
    target_master: str
    new_x: float
    new_y: float
    new_width: Optional[float]
    new_height: Optional[float]
    preserve_text: bool
    style_notes: str


class LLMConverter:
    """
    LLM-centric Visio converter that uses the LLM for:
    1. Input diagram analysis
    2. Template structure analysis
    3. Transformation planning
    4. Output specification generation
    """

    def __init__(self, api_key: Optional[str] = None, model: Optional[str] = None):
        self.api_key = api_key or os.environ.get("OPENAI_API_KEY")
        if not self.api_key:
            raise ValueError("OpenAI API key required for LLM-centric conversion")
        
        self.model = model or os.environ.get("LLM_MODEL", "gpt-4")
        
        client_kwargs = {"api_key": self.api_key}
        if os.environ.get("OPENAI_API_BASE"):
            client_kwargs["base_url"] = os.environ.get("OPENAI_API_BASE")
        if os.environ.get("OPENAI_TIMEOUT"):
            try:
                client_kwargs["timeout"] = float(os.environ.get("OPENAI_TIMEOUT"))
            except ValueError:
                pass
        
        self.client = OpenAI(**client_kwargs)
        print(f"LLM Converter initialized with model: {self.model}")

    def convert(
        self,
        source_path: str,
        template_path: str,
        output_path: str,
    ) -> Dict[str, Any]:
        """
        Full LLM-driven conversion workflow.
        
        Returns:
            Dictionary with conversion details and LLM analysis results
        """
        results = {}
        
        # Step 1: Extract raw data from both files
        print("\n[1/5] Extracting diagram data...")
        source_data = self._extract_diagram_data(source_path)
        template_data = self._extract_diagram_data(template_path)
        template_masters = self._extract_masters(template_path)
        
        results["source_shapes"] = len(source_data["shapes"])
        results["template_masters"] = len(template_masters)
        
        # Step 2: LLM analyzes input diagram
        print("[2/5] LLM analyzing input diagram structure...")
        input_analysis = self._llm_analyze_input(source_data)
        results["input_analysis"] = input_analysis
        print(f"      Identified: {input_analysis.get('diagram_type', 'unknown')} diagram")
        print(f"      Swimlanes: {input_analysis.get('swimlane_count', 0)}")
        print(f"      Flow steps: {input_analysis.get('flow_step_count', 0)}")
        
        # Step 3: LLM analyzes template capabilities
        print("[3/5] LLM analyzing template capabilities...")
        template_analysis = self._llm_analyze_template(template_data, template_masters)
        results["template_analysis"] = template_analysis
        print(f"      Style: {template_analysis.get('style_description', 'unknown')}")
        
        # Step 4: LLM generates transformation plan
        print("[4/5] LLM generating transformation plan...")
        transform_plan = self._llm_generate_transform_plan(
            source_data, input_analysis, template_analysis, template_masters
        )
        results["transform_plan"] = transform_plan
        print(f"      Generated {len(transform_plan.get('shape_transforms', []))} shape transforms")
        
        # Step 5: Apply transformations and generate output
        print("[5/5] Applying LLM-guided transformations...")
        self._apply_transformations(
            source_path, template_path, output_path, transform_plan, template_masters
        )
        results["output_path"] = output_path
        
        print(f"\n✓ LLM-driven conversion complete: {output_path}")
        return results

    def _extract_diagram_data(self, vsdx_path: str) -> Dict[str, Any]:
        """Extract detailed shape data from a .vsdx file."""
        shapes = []
        page_width = 11.0
        page_height = 8.5
        
        with zipfile.ZipFile(vsdx_path, "r") as z:
            # Get page dimensions
            try:
                pages_xml = z.read("visio/pages/pages.xml")
                pages_root = ET.fromstring(pages_xml)
                page = pages_root.find(".//v:Page", NS)
                if page is not None:
                    sheet = page.find("v:PageSheet", NS)
                    if sheet is not None:
                        for cell in sheet.findall(".//v:Cell", NS):
                            if cell.get("N") == "PageWidth":
                                page_width = float(cell.get("V", 11.0))
                            elif cell.get("N") == "PageHeight":
                                page_height = float(cell.get("V", 8.5))
            except (KeyError, ET.ParseError):
                pass
            
            # Get master names
            master_names = {}
            try:
                masters_xml = z.read("visio/masters/masters.xml")
                masters_root = ET.fromstring(masters_xml)
                for master in masters_root.findall(".//v:Master", NS):
                    mid = master.get("ID")
                    name = master.get("NameU") or master.get("Name")
                    if mid and name:
                        master_names[mid] = name
            except (KeyError, ET.ParseError):
                pass
            
            # Extract shapes from page1
            try:
                page_xml = z.read("visio/pages/page1.xml")
                page_root = ET.fromstring(page_xml)
                
                for shape in page_root.findall(".//v:Shape", NS):
                    shape_info = self._parse_shape(shape, master_names)
                    if shape_info:
                        shapes.append(asdict(shape_info))
            except (KeyError, ET.ParseError):
                pass
        
        return {
            "filename": Path(vsdx_path).name,
            "page_width": page_width,
            "page_height": page_height,
            "shapes": shapes,
        }

    def _parse_shape(self, shape: ET.Element, master_names: Dict[str, str]) -> Optional[ShapeInfo]:
        """Parse a shape element into ShapeInfo."""
        shape_id = shape.get("ID")
        if not shape_id:
            return None
        
        cells = {c.get("N"): c.get("V") for c in shape.findall("v:Cell", NS)}
        
        # Get text
        text_el = shape.find("v:Text", NS)
        text = "".join(text_el.itertext()).strip() if text_el is not None else ""
        
        # Check for heading text in User section
        if not text:
            for row in shape.findall(".//v:Row[@N='visHeadingText']", NS):
                cell = row.find("v:Cell[@N='Value']", NS)
                if cell is not None:
                    text = cell.get("V", "")
        
        master_id = shape.get("Master")
        master_name = master_names.get(master_id, "") if master_id else ""
        
        is_connector = "BeginX" in cells and "EndX" in cells
        
        def safe_float(v, default=0.0):
            try:
                return float(v) if v else default
            except (ValueError, TypeError):
                return default
        
        return ShapeInfo(
            id=shape_id,
            master_name=master_name,
            text=text[:100],  # Limit text length
            x=safe_float(cells.get("PinX")),
            y=safe_float(cells.get("PinY")),
            width=safe_float(cells.get("Width"), 1.0),
            height=safe_float(cells.get("Height"), 1.0),
            is_connector=is_connector,
            begin_x=safe_float(cells.get("BeginX")) if is_connector else None,
            begin_y=safe_float(cells.get("BeginY")) if is_connector else None,
            end_x=safe_float(cells.get("EndX")) if is_connector else None,
            end_y=safe_float(cells.get("EndY")) if is_connector else None,
        )

    def _extract_masters(self, vsdx_path: str) -> List[Dict[str, Any]]:
        """Extract master information from template."""
        masters = []
        
        with zipfile.ZipFile(vsdx_path, "r") as z:
            try:
                masters_xml = z.read("visio/masters/masters.xml")
                rels_xml = z.read("visio/masters/_rels/masters.xml.rels")
            except KeyError:
                return masters
            
            masters_root = ET.fromstring(masters_xml)
            rels_root = ET.fromstring(rels_xml)
            rels_map = {rel.get("Id"): rel.get("Target") for rel in rels_root}
            
            r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            
            for master in masters_root.findall(".//v:Master", NS):
                master_id = master.get("ID")
                name = master.get("NameU") or master.get("Name")
                
                if not master_id or not name:
                    continue
                
                # Get dimensions from master shape
                width, height = 1.0, 1.0
                rel = master.find("v:Rel", NS)
                if rel is not None:
                    rel_id = rel.get(f"{{{r_ns}}}id")
                    target = rels_map.get(rel_id)
                    if target:
                        try:
                            master_xml = z.read(f"visio/masters/{target}")
                            master_root = ET.fromstring(master_xml)
                            shape = master_root.find(".//v:Shape", NS)
                            if shape is not None:
                                cells = {c.get("N"): c.get("V") for c in shape.findall("v:Cell", NS)}
                                width = float(cells.get("Width", 1.0))
                                height = float(cells.get("Height", 1.0))
                        except (KeyError, ET.ParseError, ValueError):
                            pass
                
                masters.append({
                    "id": master_id,
                    "name": name,
                    "width": width,
                    "height": height,
                })
        
        return masters

    def _llm_analyze_input(self, source_data: Dict[str, Any]) -> Dict[str, Any]:
        """Use LLM to analyze the input diagram structure."""
        shapes_desc = []
        for s in source_data["shapes"][:50]:  # Limit for prompt size
            if s["is_connector"]:
                shapes_desc.append(
                    f"  - Connector ID={s['id']}: from ({s['begin_x']:.1f},{s['begin_y']:.1f}) "
                    f"to ({s['end_x']:.1f},{s['end_y']:.1f}), text='{s['text']}'"
                )
            else:
                shapes_desc.append(
                    f"  - Shape ID={s['id']}: master='{s['master_name']}', "
                    f"pos=({s['x']:.1f},{s['y']:.1f}), size=({s['width']:.1f}x{s['height']:.1f}), "
                    f"text='{s['text']}'"
                )
        
        prompt = f"""Analyze this Visio diagram and provide a structured understanding.

DIAGRAM: {source_data['filename']}
Page size: {source_data['page_width']:.1f} x {source_data['page_height']:.1f} inches

SHAPES ({len(source_data['shapes'])} total):
{chr(10).join(shapes_desc)}

Analyze the diagram and respond with JSON:
{{
  "diagram_type": "flowchart|swimlane|org_chart|network|other",
  "diagram_description": "brief description of what this diagram represents",
  "swimlane_count": number of swimlanes (0 if not a swimlane diagram),
  "swimlane_labels": ["list of swimlane header texts"],
  "flow_step_count": number of process/decision steps,
  "flow_direction": "top_to_bottom|left_to_right|other",
  "start_shape_id": "ID of the start shape",
  "end_shape_id": "ID of the end shape",
  "header_shape_ids": ["IDs of header/title shapes"],
  "process_shape_ids": ["IDs of process step shapes"],
  "decision_shape_ids": ["IDs of decision diamond shapes"],
  "connector_ids": ["IDs of connector lines"],
  "layout_structure": {{
    "rows": [
      {{"y_position": float, "shape_ids": ["ids at this row"], "row_type": "header|flow|other"}}
    ]
  }},
  "special_notes": "any special observations about the diagram structure"
}}"""

        response = self._call_llm(prompt, system="You are a Visio diagram analyst. Respond only with valid JSON.")
        return self._parse_json_response(response)

    def _llm_analyze_template(
        self, template_data: Dict[str, Any], template_masters: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """Use LLM to analyze template capabilities and styling."""
        masters_desc = "\n".join([
            f"  - {m['name']}: size {m['width']:.2f}x{m['height']:.2f}"
            for m in template_masters
        ])
        
        shapes_desc = []
        for s in template_data["shapes"][:30]:
            if not s["is_connector"]:
                shapes_desc.append(
                    f"  - Shape: master='{s['master_name']}', pos=({s['x']:.1f},{s['y']:.1f}), text='{s['text']}'"
                )
        
        prompt = f"""Analyze this Visio template and understand its styling capabilities.

TEMPLATE: {template_data['filename']}
Page size: {template_data['page_width']:.1f} x {template_data['page_height']:.1f} inches

AVAILABLE MASTERS:
{masters_desc}

SAMPLE SHAPES IN TEMPLATE:
{chr(10).join(shapes_desc) if shapes_desc else "  (no sample shapes)"}

Analyze the template and respond with JSON:
{{
  "style_description": "brief description of the visual style",
  "color_scheme": "description of colors used",
  "has_swimlanes": true/false,
  "swimlane_master": "name of swimlane master if available",
  "container_master": "name of container master if available", 
  "process_master": "name of process/activity master",
  "decision_master": "name of decision/diamond master",
  "start_end_master": "name of start/end terminator master",
  "connector_master": "name of connector master",
  "recommended_shape_sizes": {{
    "process": {{"width": float, "height": float}},
    "decision": {{"width": float, "height": float}},
    "start_end": {{"width": float, "height": float}}
  }},
  "layout_recommendations": {{
    "swimlane_width": float,
    "row_spacing": float,
    "shape_spacing": float
  }},
  "styling_notes": "any special styling considerations"
}}"""

        response = self._call_llm(prompt, system="You are a Visio template analyst. Respond only with valid JSON.")
        return self._parse_json_response(response)

    def _llm_generate_transform_plan(
        self,
        source_data: Dict[str, Any],
        input_analysis: Dict[str, Any],
        template_analysis: Dict[str, Any],
        template_masters: List[Dict[str, Any]],
    ) -> Dict[str, Any]:
        """Use LLM to generate detailed transformation plan."""
        shapes_desc = []
        for s in source_data["shapes"]:
            if s["is_connector"]:
                shapes_desc.append(
                    f"  ID={s['id']}: CONNECTOR from ({s['begin_x']:.2f},{s['begin_y']:.2f}) "
                    f"to ({s['end_x']:.2f},{s['end_y']:.2f}), label='{s['text']}'"
                )
            else:
                shapes_desc.append(
                    f"  ID={s['id']}: master='{s['master_name']}', "
                    f"pos=({s['x']:.2f},{s['y']:.2f}), text='{s['text']}'"
                )
        
        master_names = [m["name"] for m in template_masters]
        
        prompt = f"""Generate a transformation plan to convert this diagram to the new template style.

SOURCE DIAGRAM ANALYSIS:
{json.dumps(input_analysis, indent=2, ensure_ascii=False)}

TEMPLATE ANALYSIS:
{json.dumps(template_analysis, indent=2, ensure_ascii=False)}

AVAILABLE TARGET MASTERS: {master_names}

SOURCE SHAPES:
{chr(10).join(shapes_desc)}

Generate a complete transformation plan. For each shape, specify:
1. Which target master to use
2. New position (adjusted for template layout)
3. Whether to preserve original text

Respond with JSON:
{{
  "transformation_strategy": "description of overall approach",
  "layout_adjustments": {{
    "x_scale": float (scale factor for X coordinates),
    "y_scale": float (scale factor for Y coordinates),
    "x_offset": float (offset to add to X),
    "y_offset": float (offset to add to Y)
  }},
  "shape_transforms": [
    {{
      "shape_id": "original shape ID",
      "target_master": "name from available masters",
      "new_x": float,
      "new_y": float,
      "new_width": float or null (null = inherit from master),
      "new_height": float or null,
      "preserve_text": true/false,
      "transform_notes": "any special handling"
    }}
  ],
  "connector_transforms": [
    {{
      "connector_id": "original connector ID",
      "target_master": "Dynamic connector",
      "begin_x": float,
      "begin_y": float,
      "end_x": float,
      "end_y": float,
      "label": "text label if any"
    }}
  ],
  "post_processing": ["list of any additional adjustments needed"]
}}"""

        response = self._call_llm(
            prompt, 
            system="You are a Visio diagram transformation expert. Generate precise transformation specifications. Respond only with valid JSON."
        )
        return self._parse_json_response(response)

    def _apply_transformations(
        self,
        source_path: str,
        template_path: str,
        output_path: str,
        transform_plan: Dict[str, Any],
        template_masters: List[Dict[str, Any]],
    ) -> None:
        """Apply LLM-generated transformations to create output file."""
        from .master_injector import MasterInjector
        from .visio_rebuilder import VisioRebuilder
        
        # Build mapping from transform plan
        mapping = {}
        for t in transform_plan.get("shape_transforms", []):
            shape_id = str(t.get("shape_id"))
            target_master = t.get("target_master")
            if shape_id and target_master:
                mapping[shape_id] = target_master
        
        for t in transform_plan.get("connector_transforms", []):
            conn_id = str(t.get("connector_id"))
            target_master = t.get("target_master", "Dynamic connector")
            if conn_id:
                mapping[conn_id] = target_master
        
        # Use existing rebuilder with LLM-generated mapping
        rebuilder = VisioRebuilder(
            source_path=source_path,
            target_template_path=template_path,
            mapping=mapping,
        )
        rebuilder.rebuild(output_path)
        
        # Apply additional position adjustments from LLM plan
        self._apply_position_adjustments(output_path, transform_plan)

    def _apply_position_adjustments(
        self, output_path: str, transform_plan: Dict[str, Any]
    ) -> None:
        """Apply LLM-specified position adjustments to the output file."""
        import shutil
        import tempfile
        
        shape_transforms = {
            str(t["shape_id"]): t 
            for t in transform_plan.get("shape_transforms", [])
            if "new_x" in t and "new_y" in t
        }
        
        connector_transforms = {
            str(t["connector_id"]): t
            for t in transform_plan.get("connector_transforms", [])
            if any(k in t for k in ("begin_x", "begin_y", "end_x", "end_y"))
        }
        
        if not shape_transforms and not connector_transforms:
            return
        
        work_dir = Path(tempfile.mkdtemp())
        try:
            with zipfile.ZipFile(output_path, "r") as z:
                z.extractall(work_dir)
            
            page_path = work_dir / "visio/pages/page1.xml"
            if not page_path.exists():
                return
            
            tree = ET.parse(page_path)
            root = tree.getroot()
            
            for shape in root.findall(".//v:Shape", NS):
                shape_id = shape.get("ID")
                
                # Handle shape position and size transforms
                if shape_id in shape_transforms:
                    t = shape_transforms[shape_id]
                    self._update_shape_cells(shape, {
                        "PinX": t.get("new_x"),
                        "PinY": t.get("new_y"),
                        "Width": t.get("new_width"),
                        "Height": t.get("new_height"),
                    })
                
                # Handle connector coordinate transforms
                if shape_id in connector_transforms:
                    t = connector_transforms[shape_id]
                    self._update_shape_cells(shape, {
                        "BeginX": t.get("begin_x"),
                        "BeginY": t.get("begin_y"),
                        "EndX": t.get("end_x"),
                        "EndY": t.get("end_y"),
                    })
            
            # Write back
            tree.write(page_path, encoding="utf-8", xml_declaration=True)
            
            # Repackage
            with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as z:
                for file_path in work_dir.rglob("*"):
                    if file_path.is_file():
                        arcname = file_path.relative_to(work_dir)
                        z.write(file_path, arcname)
        finally:
            shutil.rmtree(work_dir, ignore_errors=True)

    def _update_shape_cells(
        self, shape: ET.Element, cell_updates: Dict[str, Optional[float]]
    ) -> None:
        """Update multiple cell values on a shape element.
        
        Args:
            shape: The Shape XML element
            cell_updates: Dict mapping cell names (e.g., 'PinX') to new values.
                          If value is None, the cell is skipped.
        """
        # Filter out None values
        updates = {k: v for k, v in cell_updates.items() if v is not None}
        if not updates:
            return
        
        # Track which cells we've updated
        updated = set()
        
        # Update existing cells
        for cell in shape.findall("v:Cell", NS):
            cell_name = cell.get("N")
            if cell_name in updates:
                cell.set("V", str(updates[cell_name]))
                # Remove formula if present (we're setting explicit value)
                if "F" in cell.attrib:
                    del cell.attrib["F"]
                updated.add(cell_name)
        
        # Add missing cells for connectors (BeginX/BeginY/EndX/EndY)
        for cell_name, value in updates.items():
            if cell_name not in updated:
                new_cell = ET.Element(f"{{{NS['v']}}}Cell")
                new_cell.set("N", cell_name)
                new_cell.set("V", str(value))
                shape.insert(0, new_cell)

    def _call_llm(self, prompt: str, system: str = "You are a helpful assistant.") -> str:
        """Call the LLM API."""
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": prompt},
                ],
                temperature=0.3,
                response_format={"type": "json_object"},
            )
            return response.choices[0].message.content
        except Exception as e:
            raise RuntimeError(f"LLM API call failed: {e}")

    def _parse_json_response(self, response: str) -> Dict[str, Any]:
        """Parse JSON response from LLM."""
        try:
            clean = response.strip()
            if clean.startswith("```"):
                clean = clean.split("\n", 1)[1]
                if clean.endswith("```"):
                    clean = clean.rsplit("\n", 1)[0]
            return json.loads(clean)
        except json.JSONDecodeError as e:
            print(f"Warning: Failed to parse LLM response: {e}")
            return {}
