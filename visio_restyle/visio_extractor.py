"""
Extract Visio diagram information to JSON format.
"""

import json
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Any, Optional
from pathlib import Path
from vsdx import VisioFile


class VisioExtractor:
    """Extract shape and connector information from a Visio diagram."""

    def __init__(self, visio_path: str):
        """Initialize the extractor with a Visio file path.
        
        Args:
            visio_path: Path to the .vsdx file
        """
        self.visio_path = Path(visio_path)
        self.visio_file = None
        self._master_name_by_id = self._load_master_name_map()

    def _load_master_name_map(self) -> Dict[str, str]:
        master_map = {}
        try:
            with zipfile.ZipFile(self.visio_path, 'r') as z:
                if 'visio/masters/masters.xml' not in z.namelist():
                    return master_map
                xml_content = z.read('visio/masters/masters.xml')
            root = ET.fromstring(xml_content)
            ns = {'v': 'http://schemas.microsoft.com/office/visio/2012/main'}
            for master in root.findall('.//v:Master', ns):
                master_id = master.get('ID')
                name = master.get('NameU') or master.get('Name')
                if master_id and name:
                    master_map[master_id] = name
        except Exception:
            return master_map
        return master_map
    
    def extract(self) -> Dict[str, Any]:
        """Extract the complete diagram structure.
        
        Returns:
            Dictionary containing shapes, connectors, and metadata
        """
        self.visio_file = VisioFile(str(self.visio_path))
        
        diagram_data = {
            "filename": self.visio_path.name,
            "pages": []
        }
        
        # Process each page in the diagram
        for page in self.visio_file.pages:
            page_data = {
                "name": page.name,
                "shapes": self._extract_shapes(page),
                "connectors": self._extract_connectors(page)
            }
            diagram_data["pages"].append(page_data)
        
        return diagram_data
    
    def _extract_shapes(self, page) -> List[Dict[str, Any]]:
        """Extract all shapes from a page.
        
        Args:
            page: VisioFile page object
            
        Returns:
            List of shape dictionaries
        """
        shapes = []
        
        # Use child_shapes if available, otherwise shapes
        # child_shapes usually contains top-level shapes
        page_shapes = page.child_shapes if hasattr(page, 'child_shapes') else page.shapes
        
        for shape in page_shapes:
            # Skip connectors (they're handled separately)
            if self._is_connector(shape):
                continue
            
            shape_data = {
                "id": shape.ID,
                "text": shape.text if hasattr(shape, 'text') else "",
                "master_name": self._get_master_name(shape),
                "master_id": self._get_master_id(shape),
                "position": self._get_position(shape),
                "size": self._get_size(shape),
                "properties": self._get_properties(shape)
            }
            shapes.append(shape_data)
        
        return shapes
    
    def _extract_connectors(self, page) -> List[Dict[str, Any]]:
        """Extract all connectors from a page.
        
        Args:
            page: VisioFile page object
            
        Returns:
            List of connector dictionaries
        """
        connectors = []
        
        page_shapes = page.child_shapes if hasattr(page, 'child_shapes') else page.shapes
        
        for shape in page_shapes:
            if not self._is_connector(shape):
                continue
            
            connector_data = {
                "id": shape.ID,
                "text": shape.text if hasattr(shape, 'text') else "",
                "master_name": self._get_master_name(shape),
                "from_shape": self._get_connection_source(shape),
                "to_shape": self._get_connection_target(shape),
                "properties": self._get_properties(shape)
            }
            connectors.append(connector_data)
        
        return connectors
    
    def _is_connector(self, shape) -> bool:
        """Check if a shape is a connector.
        
        Args:
            shape: Shape object
            
        Returns:
            True if the shape is a connector
        """
        # Check if shape has connector properties
        try:
            # Connectors typically have BeginX/EndX cells
            if hasattr(shape, 'cells'):
                return 'BeginX' in shape.cells or shape.Type == '1'
        except:
            pass
        return False
    
    def _get_master_name(self, shape) -> Optional[str]:
        """Get the master name of a shape.
        
        Args:
            shape: Shape object
            
        Returns:
            Master name or None
        """
        try:
            if hasattr(shape, 'master_shape') and shape.master_shape:
                name = shape.master_shape.Name if hasattr(shape.master_shape, 'Name') else None
                if name:
                    return name
            master_id = self._get_master_id(shape)
            if master_id:
                return self._master_name_by_id.get(str(master_id))
            return None
        except:
            return None
    
    def _get_master_id(self, shape) -> Optional[str]:
        """Get the master ID of a shape.
        
        Args:
            shape: Shape object
            
        Returns:
            Master ID or None
        """
        try:
            if hasattr(shape, 'xml') and shape.xml is not None:
                return shape.xml.get('Master')
            if hasattr(shape, 'master_shape') and shape.master_shape:
                return shape.master_shape.ID if hasattr(shape.master_shape, 'ID') else None
            return None
        except:
            return None
    
    def _get_position(self, shape) -> Dict[str, float]:
        """Get the position of a shape.
        
        Args:
            shape: Shape object
            
        Returns:
            Dictionary with x, y coordinates
        """
        try:
            x = float(shape.cells.get('PinX', {}).get('value', 0))
            y = float(shape.cells.get('PinY', {}).get('value', 0))
            return {"x": x, "y": y}
        except:
            return {"x": 0.0, "y": 0.0}
    
    def _get_size(self, shape) -> Dict[str, float]:
        """Get the size of a shape.
        
        Args:
            shape: Shape object
            
        Returns:
            Dictionary with width, height
        """
        try:
            width = float(shape.cells.get('Width', {}).get('value', 1))
            height = float(shape.cells.get('Height', {}).get('value', 1))
            return {"width": width, "height": height}
        except:
            return {"width": 1.0, "height": 1.0}
    
    def _get_properties(self, shape) -> Dict[str, Any]:
        """Get custom properties of a shape.
        
        Args:
            shape: Shape object
            
        Returns:
            Dictionary of properties
        """
        properties = {}
        try:
            if hasattr(shape, 'properties'):
                properties = dict(shape.properties)
        except:
            pass
        return properties
    
    def _get_connection_source(self, connector) -> Optional[str]:
        """Get the source shape ID of a connector.
        
        Args:
            connector: Connector shape object
            
        Returns:
            Source shape ID or None
        """
        try:
            if hasattr(connector, 'connects'):
                for connect in connector.connects:
                    if hasattr(connect, 'from_rel') and connect.from_rel == 'BeginX':
                        return connect.to_shape.ID if hasattr(connect.to_shape, 'ID') else None
        except:
            pass
        return None
    
    def _get_connection_target(self, connector) -> Optional[str]:
        """Get the target shape ID of a connector.
        
        Args:
            connector: Connector shape object
            
        Returns:
            Target shape ID or None
        """
        try:
            if hasattr(connector, 'connects'):
                for connect in connector.connects:
                    if hasattr(connect, 'from_rel') and connect.from_rel == 'EndX':
                        return connect.to_shape.ID if hasattr(connect.to_shape, 'ID') else None
        except:
            pass
        return None
    
    def save_to_json(self, output_path: str) -> None:
        """Extract and save diagram data to JSON file.
        
        Args:
            output_path: Path to save the JSON file
        """
        data = self.extract()
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    
    @staticmethod
    def get_masters_from_visio(visio_path: str) -> List[Dict[str, str]]:
        """Extract list of available masters from a Visio file.
        
        Args:
            visio_path: Path to the .vsdx file
            
        Returns:
            List of master dictionaries with name and description
        """
        masters = []
        
        try:
            import zipfile
            import xml.etree.ElementTree as ET
            
            with zipfile.ZipFile(visio_path, 'r') as z:
                if 'visio/masters/masters.xml' in z.namelist():
                    xml_content = z.read('visio/masters/masters.xml')
                    root = ET.fromstring(xml_content)
                    
                    # Visio namespace
                    ns = {'v': 'http://schemas.microsoft.com/office/visio/2012/main'}
                    
                    for master in root.findall('.//v:Master', ns):
                        # Try NameU (Universal Name) first, then Name
                        name = master.get('NameU') or master.get('Name')
                        master_id = master.get('ID')
                        unique_id = master.get('UniqueID', '')
                        
                        if name and master_id:
                            masters.append({
                                "name": name,
                                "id": master_id,
                                "description": unique_id
                            })
                            
        except Exception as e:
            print(f"Warning: Could not extract masters directly from XML: {e}")
            
        # Fallback to vsdx library if XML extraction failed or returned empty
        if not masters:
            try:
                visio_file = VisioFile(str(visio_path))
                if hasattr(visio_file, 'masters'):
                    for master in visio_file.masters:
                        master_info = {
                            "name": master.Name if hasattr(master, 'Name') else str(master),
                            "id": master.ID if hasattr(master, 'ID') else None,
                            "description": master.UniqueID if hasattr(master, 'UniqueID') else ""
                        }
                        masters.append(master_info)
            except Exception as e:
                print(f"Warning: Could not extract masters via vsdx lib: {e}")
        
        return masters
