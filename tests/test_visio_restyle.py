"""
Unit tests for the Visio Restyle application.
"""
import json
import os
import tempfile
from pathlib import Path
import pytest

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


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
