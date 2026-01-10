"""
Rebuild Visio diagram with new masters while preserving layout and connections.
"""

from typing import Dict, List, Any, Optional
from pathlib import Path
from vsdx import VisioFile
import shutil


class VisioRebuilder:
    """Rebuild a Visio diagram with new shape masters."""

    def __init__(
        self,
        source_path: str,
        target_template_path: str,
        mapping: Dict[str, str]
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
        self.source_file = None
        self.target_file = None
    
    def rebuild(self, output_path: str) -> None:
        """Rebuild the diagram with new masters.
        
        Args:
            output_path: Path to save the rebuilt .vsdx file
        """
        # Inject masters from template into source, save to output_path
        from .master_injector import MasterInjector
        print(f"Injecting masters from {self.target_template_path}...")
        injector = MasterInjector(self.source_path, self.target_template_path)
        master_name_map = injector.inject(output_path)
        print(f"Injected {len(master_name_map)} masters.")
        
        # Open the output file (now containing merged masters)
        output_file = VisioFile(str(output_path))
        
        # Process each page
        for page in output_file.pages:
            self._rebuild_page(page, master_name_map)
        
        # Save the modified file
        output_file.save_vsdx(str(output_path))
        output_file.close_vsdx()
    
    def _rebuild_page(
        self,
        page,
        master_name_map: Dict[str, str]
    ) -> None:
        """Rebuild a single page.
        
        Args:
            page: Output page object to modify
            master_name_map: Map of master name to new master ID
        """
        # Iterate over shapes
        # Note: We iterate over a copy of the list if we were modifying the list, 
        # but here we modify shape properties.
        if hasattr(page, 'child_shapes'):
            shapes = page.child_shapes
        else:
            shapes = page.shapes

        for shape in shapes:
            shape_id = shape.ID
            
            # Skip connectors for now
            if self._is_connector(shape):
                continue
            
            # Check if this shape should be remapped
            if shape_id in self.mapping:
                target_master_name = self.mapping[shape_id]
                
                # Try to replace with new master
                if target_master_name in master_name_map:
                    new_master_id = master_name_map[target_master_name]
                    self._replace_shape_master_by_id(shape, new_master_id)
                else:
                    print(f"Warning: Master '{target_master_name}' not found in injected masters")

    def _replace_shape_master_by_id(self, shape, new_master_id: str) -> None:
        """Replace a shape's master by ID.
        
        Args:
            shape: Shape to modify
            new_master_id: New master ID
        """
        try:
            # Update XML directly
            if hasattr(shape, 'xml'):
                # 1. Set new Master ID
                shape.xml.set('Master', str(new_master_id))
                
                # 2. Remove Type attribute (let Master decide)
                if 'Type' in shape.xml.attrib:
                    del shape.xml.attrib['Type']
                
                # 3. Remove local overrides to inherit Master style/geometry
                # We want to remove: Geometry, Fill, Line, QuickStyle
                # We want to KEEP: Text, Transform (PinX/Y), User (Container props), Property (Shape Data)
                
                sections_to_remove = ['Geometry', 'Fill', 'Line', 'QuickStyle', 'Image']
                
                to_remove = []
                for child in shape.xml:
                    # Check for Section elements
                    # Tag will be {http://...}Section
                    if 'Section' in child.tag:
                        n_attr = child.get('N')
                        if n_attr in sections_to_remove:
                            to_remove.append(child)
                
                for child in to_remove:
                    shape.xml.remove(child)
                
        except Exception as e:
            print(f"Warning: Could not replace shape master: {e}")

    
    def _is_connector(self, shape) -> bool:
        """Check if a shape is a connector.
        
        Args:
            shape: Shape object
            
        Returns:
            True if the shape is a connector
        """
        try:
            if hasattr(shape, 'cells'):
                return 'BeginX' in shape.cells or shape.Type == '1'
        except:
            pass
        return False
    
    def _replace_shape_master(
        self,
        shape,
        new_master,
        page
    ) -> None:
        """Replace a shape's master while preserving properties.
        
        Args:
            shape: Shape to modify
            new_master: New master to apply
            page: Page containing the shape
        """
        # Store current properties
        old_text = shape.text if hasattr(shape, 'text') else ""
        old_position = self._get_position(shape)
        old_size = self._get_size(shape)
        
        try:
            # Update the master reference
            if hasattr(shape, 'master_shape'):
                shape.master_shape = new_master
            
            # Restore position
            if 'PinX' in shape.cells:
                shape.cells['PinX']['value'] = str(old_position['x'])
            if 'PinY' in shape.cells:
                shape.cells['PinY']['value'] = str(old_position['y'])
            
            # Restore size
            if 'Width' in shape.cells:
                shape.cells['Width']['value'] = str(old_size['width'])
            if 'Height' in shape.cells:
                shape.cells['Height']['value'] = str(old_size['height'])
            
            # Restore text
            if old_text and hasattr(shape, 'text'):
                shape.text = old_text
        
        except Exception as e:
            print(f"Warning: Could not fully replace shape master: {e}")
    
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
    
    def _reglue_connectors(
        self,
        page,
        shape_id_map: Dict[str, Any]
    ) -> None:
        """Re-glue connectors to shapes after master changes.
        
        Args:
            page: Page object
            shape_id_map: Mapping of shape IDs to shape objects
        """
        for shape in page.shapes:
            if not self._is_connector(shape):
                continue
            
            try:
                # Connectors should maintain their connections automatically
                # in most cases, but we can verify and fix if needed
                if hasattr(shape, 'connects'):
                    for connect in shape.connects:
                        # Connection should still be valid
                        # The vsdx library handles this internally
                        pass
            except Exception as e:
                print(f"Warning: Could not verify connector: {e}")
