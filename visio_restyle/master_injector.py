import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import tempfile
import re

# Namespaces
NS = {
    'v': 'http://schemas.microsoft.com/office/visio/2012/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types'
}
for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)

class MasterInjector:
    def __init__(self, source_path, template_path):
        self.source_path = Path(source_path)
        self.template_path = Path(template_path)
        self.work_dir = Path(tempfile.mkdtemp())
        self.source_dir = self.work_dir / "source"
        self.template_dir = self.work_dir / "template"
        self.output_dir = self.work_dir / "output"
        
        # Unzip files
        with zipfile.ZipFile(self.source_path, 'r') as z:
            z.extractall(self.source_dir)
        with zipfile.ZipFile(self.template_path, 'r') as z:
            z.extractall(self.template_dir)
            
        # Prepare output dir (copy source as base)
        shutil.copytree(self.source_dir, self.output_dir)

    def inject(self, output_path):
        try:
            # 1. Get max Master ID in source
            masters_xml_path = self.output_dir / "visio/masters/masters.xml"
            max_id = 0
            existing_masters = []
            
            if masters_xml_path.exists():
                tree = ET.parse(masters_xml_path)
                root = tree.getroot()
                for master in root.findall('.//v:Master', NS):
                    mid = int(master.get('ID', 0))
                    if mid > max_id:
                        max_id = mid
                    existing_masters.append(mid)
            else:
                # Create masters.xml if not exists
                root = ET.Element(f"{{{NS['v']}}}Masters")
                tree = ET.ElementTree(root)
                # Ensure directories exist
                masters_xml_path.parent.mkdir(parents=True, exist_ok=True)
            
            # 2. Read template masters
            template_masters_xml = self.template_dir / "visio/masters/masters.xml"
            if not template_masters_xml.exists():
                print("No masters in template")
                return {}

            t_tree = ET.parse(template_masters_xml)
            t_root = t_tree.getroot()
            
            # Map: Template Master Name -> New ID in Source
            master_name_map = {}
            
            # 3. Copy masters
            rels_path = self.output_dir / "visio/masters/_rels/masters.xml.rels"
            if rels_path.exists():
                rels_tree = ET.parse(rels_path)
                rels_root = rels_tree.getroot()
            else:
                rels_root = ET.Element(f"{{{NS['r']}}}Relationships")
                rels_tree = ET.ElementTree(rels_root)
                rels_path.parent.mkdir(parents=True, exist_ok=True)
                
            # Parse template rels to find master file paths
            t_rels_path = self.template_dir / "visio/masters/_rels/masters.xml.rels"
            t_rels_map = {} # rId -> target
            if t_rels_path.exists():
                tr_tree = ET.parse(t_rels_path)
                for rel in tr_tree.getroot():
                    t_rels_map[rel.get('Id')] = rel.get('Target')

            # Content Types
            ct_path = self.output_dir / "[Content_Types].xml"
            ct_tree = ET.parse(ct_path)
            ct_root = ct_tree.getroot()
            
            for t_master in t_root.findall('.//v:Master', NS):
                max_id += 1
                new_id = str(max_id)
                
                # Get Master Name
                name = t_master.get('NameU') or t_master.get('Name')
                master_name_map[name] = new_id
                
                # Get Rel ID
                r_id = None
                for rel in t_master.findall('.//v:Rel', NS):
                    r_id = rel.get(f"{{{NS['r']}}}id")
                    break
                
                if not r_id or r_id not in t_rels_map:
                    print(f"Skipping master {name}: No relationship found")
                    continue
                    
                t_file_rel = t_rels_map[r_id]
                t_file_name = Path(t_file_rel).name
                
                # Copy Master File
                new_file_name = f"master{new_id}.xml"
                src_file = self.template_dir / "visio/masters" / t_file_name
                dst_file = self.output_dir / "visio/masters" / new_file_name
                
                if src_file.exists():
                    shutil.copy2(src_file, dst_file)
                    
                    # Add to Masters.xml
                    t_master.set('ID', new_id)
                    # Generate new Rel ID
                    new_rel_id = f"rId{new_id}_injected"
                    
                    # Update Rel in Master element
                    for rel in t_master.findall('.//v:Rel', NS):
                        rel.set(f"{{{NS['r']}}}id", new_rel_id)
                    
                    root.append(t_master)
                    
                    # Add to .rels
                    new_rel = ET.SubElement(rels_root, f"{{{NS['r']}}}Relationship")
                    new_rel.set('Id', new_rel_id)
                    new_rel.set('Type', "http://schemas.microsoft.com/office/visio/2012/relationships/master")
                    new_rel.set('Target', new_file_name)
                    
                    # Add to [Content_Types].xml
                    # Check if Override exists
                    part_name = f"/visio/masters/{new_file_name}"
                    found = False
                    for override in ct_root.findall(f".//ct:Override", NS):
                        if override.get('PartName') == part_name:
                            found = True
                            break
                    if not found:
                        override = ET.SubElement(ct_root, f"{{{NS['ct']}}}Override")
                        override.set('PartName', part_name)
                        override.set('ContentType', "application/vnd.ms-visio.master+xml")
                else:
                    print(f"Warning: Master file {src_file} not found")

            # Save XMLs
            tree.write(masters_xml_path, encoding='utf-8', xml_declaration=True)
            rels_tree.write(rels_path, encoding='utf-8', xml_declaration=True)
            ct_tree.write(ct_path, encoding='utf-8', xml_declaration=True)
            
            # Zip output
            shutil.make_archive(str(Path(output_path).with_suffix('')), 'zip', self.output_dir)
            shutil.move(str(Path(output_path).with_suffix('.zip')), output_path)
            
            return master_name_map

        finally:
            shutil.rmtree(self.work_dir, ignore_errors=True)

if __name__ == "__main__":
    # Test
    pass
