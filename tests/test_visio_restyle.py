"""
Unit tests for the Visio Restyle application.
"""
import json
import os
import tempfile
from pathlib import Path
import pytest
import zipfile
import xml.etree.ElementTree as ET

from visio_restyle.auto_mapper import AutoMapper
from visio_restyle.visio_extractor import VisioExtractor
from visio_restyle.visio_rebuilder import VisioRebuilder

# Test data
SAMPLE_DIAGRAM_DATA = {
    "filename": "test.vsdx",
    "pages": [
        {
            "name": "Page-1",
            "shapes": [
                {
                    "id": "1",
                    "text": "Start",
                    "master_name": "Process",
                    "master_id": "1",
                    "position": {"x": 4.25, "y": 8.5},
                    "size": {"width": 1.5, "height": 0.75},
                    "properties": {}
                },
                {
                    "id": "2",
                    "text": "Decision",
                    "master_name": "Decision",
                    "master_id": "2",
                    "position": {"x": 4.25, "y": 6.5},
                    "size": {"width": 1.5, "height": 1.0},
                    "properties": {}
                }
            ],
            "connectors": [
                {
                    "id": "3",
                    "text": "",
                    "master_name": "Connector",
                    "from_shape": "1",
                    "to_shape": "2",
                    "properties": {}
                }
            ]
        }
    ]
}

SAMPLE_MASTERS = {
    "masters": [
        {"name": "ModernProcess", "id": "1", "description": "Modern process shape"},
        {"name": "ModernDecision", "id": "2", "description": "Modern decision shape"},
        {"name": "ModernData", "id": "3", "description": "Modern data shape"},
    ]
}

SAMPLE_MAPPING = {
    "1": "ModernProcess",
    "2": "ModernDecision"
}


class TestJSONSerialization:
    """Test JSON serialization and deserialization."""
    
    def test_diagram_data_serialization(self):
        """Test that diagram data can be serialized to JSON."""
        json_str = json.dumps(SAMPLE_DIAGRAM_DATA, indent=2)
        assert json_str is not None
        assert "filename" in json_str
        assert "pages" in json_str
    
    def test_diagram_data_deserialization(self):
        """Test that diagram data can be deserialized from JSON."""
        json_str = json.dumps(SAMPLE_DIAGRAM_DATA)
        data = json.loads(json_str)
        assert data["filename"] == "test.vsdx"
        assert len(data["pages"]) == 1
        assert len(data["pages"][0]["shapes"]) == 2
        assert len(data["pages"][0]["connectors"]) == 1
    
    def test_masters_serialization(self):
        """Test that masters data can be serialized to JSON."""
        json_str = json.dumps(SAMPLE_MASTERS, indent=2)
        assert json_str is not None
        assert "masters" in json_str
    
    def test_mapping_serialization(self):
        """Test that mapping data can be serialized to JSON."""
        json_str = json.dumps(SAMPLE_MAPPING, indent=2)
        assert json_str is not None
        data = json.loads(json_str)
        assert data["1"] == "ModernProcess"
        assert data["2"] == "ModernDecision"


class TestLLMMapper:
    """Test LLM mapper functionality."""
    
    def test_mapping_validation(self):
        """Test that mapping format is valid."""
        # Mapping should be a dict with string keys and string values
        assert isinstance(SAMPLE_MAPPING, dict)
        for key, value in SAMPLE_MAPPING.items():
            assert isinstance(key, str)
            assert isinstance(value, str)
    
    def test_prompt_building(self):
        """Test that prompt can be built from diagram and masters."""
        from visio_restyle.llm_mapper import LLMMapper
        
        # We can't call _build_prompt directly without an API key,
        # but we can test the data structures
        shapes = SAMPLE_DIAGRAM_DATA["pages"][0]["shapes"]
        masters = SAMPLE_MASTERS["masters"]
        
        assert len(shapes) == 2
        assert len(masters) == 3
        assert shapes[0]["master_name"] == "Process"
        assert masters[0]["name"] == "ModernProcess"


class TestFileOperations:
    """Test file I/O operations."""
    
    def test_save_and_load_diagram(self):
        """Test saving and loading diagram JSON."""
        with tempfile.TemporaryDirectory() as tmpdir:
            filepath = Path(tmpdir) / "diagram.json"
            
            # Save
            with open(filepath, 'w') as f:
                json.dump(SAMPLE_DIAGRAM_DATA, f)
            
            # Load
            with open(filepath, 'r') as f:
                loaded_data = json.load(f)
            
            assert loaded_data["filename"] == SAMPLE_DIAGRAM_DATA["filename"]
            assert len(loaded_data["pages"]) == len(SAMPLE_DIAGRAM_DATA["pages"])
    
    def test_save_and_load_masters(self):
        """Test saving and loading masters JSON."""
        with tempfile.TemporaryDirectory() as tmpdir:
            filepath = Path(tmpdir) / "masters.json"
            
            # Save
            with open(filepath, 'w') as f:
                json.dump(SAMPLE_MASTERS, f)
            
            # Load
            with open(filepath, 'r') as f:
                loaded_data = json.load(f)
            
            assert "masters" in loaded_data
            assert len(loaded_data["masters"]) == len(SAMPLE_MASTERS["masters"])
    
    def test_save_and_load_mapping(self):
        """Test saving and loading mapping JSON."""
        from visio_restyle.llm_mapper import LLMMapper
        
        with tempfile.TemporaryDirectory() as tmpdir:
            filepath = Path(tmpdir) / "mapping.json"
            
            # Save
            LLMMapper.save_mapping(None, SAMPLE_MAPPING, str(filepath))
            
            # Load
            loaded_mapping = LLMMapper.load_mapping(str(filepath))
            
            assert loaded_mapping["1"] == SAMPLE_MAPPING["1"]
            assert loaded_mapping["2"] == SAMPLE_MAPPING["2"]


class TestDataValidation:
    """Test data validation and structure."""
    
    def test_shape_has_required_fields(self):
        """Test that shapes have required fields."""
        shape = SAMPLE_DIAGRAM_DATA["pages"][0]["shapes"][0]
        required_fields = ["id", "text", "master_name", "position", "size"]
        
        for field in required_fields:
            assert field in shape, f"Shape missing required field: {field}"
    
    def test_connector_has_required_fields(self):
        """Test that connectors have required fields."""
        connector = SAMPLE_DIAGRAM_DATA["pages"][0]["connectors"][0]
        required_fields = ["id", "from_shape", "to_shape"]
        
        for field in required_fields:
            assert field in connector, f"Connector missing required field: {field}"
    
    def test_position_has_x_and_y(self):
        """Test that position has x and y coordinates."""
        position = SAMPLE_DIAGRAM_DATA["pages"][0]["shapes"][0]["position"]
        assert "x" in position
        assert "y" in position
        assert isinstance(position["x"], (int, float))
        assert isinstance(position["y"], (int, float))
    
    def test_size_has_width_and_height(self):
        """Test that size has width and height."""
        size = SAMPLE_DIAGRAM_DATA["pages"][0]["shapes"][0]["size"]
        assert "width" in size
        assert "height" in size
        assert isinstance(size["width"], (int, float))
        assert isinstance(size["height"], (int, float))


class TestConversionWorkflow:
    """Integration-style test for conversion pipeline without LLM."""

    def test_rebuild_output_contains_shapes(self, tmp_path):
        output_path = tmp_path / "output.vsdx"

        extractor = VisioExtractor("input.vsdx")
        diagram_data = extractor.extract()
        template_masters = VisioExtractor.get_masters_from_visio("template.vsdx")

        mapping = AutoMapper().create_mapping(diagram_data, template_masters)
        assert mapping, "Expected non-empty mapping from auto mapper"

        rebuilder = VisioRebuilder(
            source_path="input.vsdx",
            target_template_path="template.vsdx",
            mapping=mapping,
        )
        rebuilder.rebuild(str(output_path))

        with zipfile.ZipFile(output_path, "r") as z:
            page_xml = z.read("visio/pages/page1.xml").decode("utf-8")
            masters_xml = z.read("visio/masters/masters.xml").decode("utf-8")
            rels_xml = z.read("visio/pages/_rels/page1.xml.rels").decode("utf-8")

        assert "<PageContents" in page_xml
        assert "<Shape" in page_xml

        # Ensure output contains expected template masters.
        assert "Process" in masters_xml
        assert "Decision" in masters_xml

        # Style overrides from template should be applied to shapes.
        assert "FillBkgnd" in page_xml or "LineColor" in page_xml

        # Page relationships should include injected masters.
        assert "master14.xml" in rels_xml
        assert "master17.xml" in rels_xml
        assert "master18.xml" in rels_xml
        assert "master19.xml" in rels_xml

        # Shapes should be scaled when template page is wider.
        def max_pinx(xml_text):
            root = ET.fromstring(xml_text)
            max_val = 0.0
            for shape in root.findall(".//{http://schemas.microsoft.com/office/visio/2012/main}Shape"):
                for cell in shape.findall("{http://schemas.microsoft.com/office/visio/2012/main}Cell"):
                    if cell.get("N") == "PinX":
                        try:
                            max_val = max(max_val, float(cell.get("V")))
                        except (TypeError, ValueError):
                            pass
            return max_val

        input_page = zipfile.ZipFile("input.vsdx").read("visio/pages/page1.xml").decode("utf-8")
        output_max = max_pinx(page_xml)
        input_max = max_pinx(input_page)
        assert output_max > input_max

        ns = {"v": "http://schemas.microsoft.com/office/visio/2012/main"}
        masters_root = ET.fromstring(masters_xml)
        master_name_by_id = {}
        for master in masters_root.findall(".//v:Master", ns):
            master_id = master.get("ID")
            name = master.get("NameU") or master.get("Name")
            if master_id and name:
                master_name_by_id[master_id] = name

        page_root = ET.fromstring(page_xml)
        lane_shapes = []
        container_shapes = []
        for shape in page_root.findall(".//v:Shape", ns):
            master_id = shape.get("Master")
            if not master_id:
                continue
            name = master_name_by_id.get(master_id)
            if name == "Swimlane (vertical)":
                lane_shapes.append(shape)
            elif name == "CFF Container":
                container_shapes.append(shape)

        assert container_shapes, "Expected a container shape mapped to CFF Container"
        assert len(lane_shapes) >= 2, "Expected swimlane shapes to be mapped"

        def get_user_cell(shape, row_name):
            user_section = shape.find("v:Section[@N='User']", ns)
            if user_section is None:
                return None
            row = user_section.find(f"v:Row[@N='{row_name}']", ns)
            if row is None:
                return None
            cell = row.find("v:Cell[@N='Value']", ns)
            return cell.get("V") if cell is not None else None

        num_lanes = None
        for shape in container_shapes:
            num_lanes = get_user_cell(shape, "numLanes")
            if num_lanes:
                break
        assert num_lanes is not None
        assert int(float(num_lanes)) >= 2


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
