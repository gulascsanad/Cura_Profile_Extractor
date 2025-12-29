#!/usr/bin/env python3
"""
Cura Profile Extractor v1.1.0
=============================
Extracts ALL Cura settings into a single, searchable JSON file.

Resolves Cura's 8-layer inheritance system:
  fdmprinter.def.json → creality_base → machine-specific → quality → user overrides

Features:
  - Auto-detects Cura install and AppData paths
  - Extracts preferences, machine settings, G-code, quality profiles, materials
  - Tracks source file for every setting
  - Human-readable output with formatted arrays and G-code
  - Quick-reference summary section
  - Key settings extraction for common values
  - GUI (default) or CLI mode

v1.1.0 Changes:
  - Semicolon-delimited lists now formatted as sorted arrays
  - G-code split into readable line arrays
  - Added _summary section with quick overview
  - Added _key_settings section with important values
  - Added --raw flag to skip formatting

Usage:
  python cura_profile_extractor.py          # GUI mode
  python cura_profile_extractor.py --cli    # CLI mode
  python cura_profile_extractor.py --help   # Help

Author: Brian's 3D Printer Project
Date: 2025-12-28
License: MIT
"""

import argparse
import configparser
import json
import os
import re
import sys
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import unquote

__version__ = "1.1.0"

# =============================================================================
# Path Detection
# =============================================================================

def find_cura_install_path() -> Optional[Path]:
    """Auto-detect Cura installation directory."""
    search_paths = [
        Path(os.environ.get("PROGRAMFILES", "C:/Program Files")),
        Path(os.environ.get("PROGRAMFILES(X86)", "C:/Program Files (x86)")),
        Path(os.environ.get("LOCALAPPDATA", "")),
    ]
    
    candidates = []
    for base in search_paths:
        if not base.exists():
            continue
        # Look for UltiMaker Cura folders
        for item in base.iterdir():
            if item.is_dir() and "cura" in item.name.lower():
                # Check if it has the expected structure
                if (item / "share" / "cura" / "resources").exists():
                    # Extract version from folder name
                    match = re.search(r'(\d+\.\d+\.?\d*)', item.name)
                    version = match.group(1) if match else "0.0.0"
                    candidates.append((version, item))
    
    if not candidates:
        return None
    
    # Return newest version
    candidates.sort(key=lambda x: [int(p) for p in x[0].split('.')], reverse=True)
    return candidates[0][1]


def find_cura_appdata_path() -> Optional[Path]:
    """Auto-detect Cura AppData directory."""
    appdata = Path(os.environ.get("APPDATA", ""))
    if not appdata.exists():
        return None
    
    cura_dir = appdata / "cura"
    if not cura_dir.exists():
        return None
    
    # Find newest version folder
    versions = []
    for item in cura_dir.iterdir():
        if item.is_dir() and re.match(r'^\d+\.\d+', item.name):
            versions.append(item)
    
    if not versions:
        return None
    
    versions.sort(key=lambda x: [int(p) for p in x.name.split('.')[:2]], reverse=True)
    return versions[0]


def get_default_paths() -> Tuple[str, str]:
    """Get default paths with auto-detection."""
    install = find_cura_install_path()
    appdata = find_cura_appdata_path()
    
    install_str = str(install) if install else "C:/Program Files/UltiMaker Cura 5.11.0"
    appdata_str = str(appdata) if appdata else ""
    
    return install_str, appdata_str


# =============================================================================
# File Parsers
# =============================================================================

def parse_cfg_file(filepath: Path) -> Dict[str, Any]:
    """Parse Cura .cfg or .inst.cfg file (INI-style format)."""
    result = {
        "_filepath": str(filepath),
        "_filename": filepath.name,
    }
    
    if not filepath.exists():
        result["_error"] = "File not found"
        return result
    
    try:
        # Read with configparser
        config = configparser.ConfigParser(interpolation=None)
        config.read(filepath, encoding='utf-8')
        
        for section in config.sections():
            result[section] = dict(config[section])
        
        return result
    except Exception as e:
        result["_error"] = str(e)
        return result


def parse_def_json(filepath: Path) -> Dict[str, Any]:
    """Parse Cura .def.json definition file."""
    result = {
        "_filepath": str(filepath),
        "_filename": filepath.name,
    }
    
    if not filepath.exists():
        result["_error"] = "File not found"
        return result
    
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        result.update(data)
        return result
    except Exception as e:
        result["_error"] = str(e)
        return result


def extract_settings_from_def(def_data: Dict[str, Any], prefix: str = "") -> Dict[str, Dict[str, Any]]:
    """
    Recursively extract all settings from a definition file.
    Returns {setting_key: {default_value, type, description, ...}}
    """
    settings = {}
    
    def recurse(node: Dict[str, Any], path: str = ""):
        if "children" in node:
            for key, child in node["children"].items():
                recurse(child, key)
        
        # Extract setting properties
        if "type" in node and node["type"] != "category":
            setting_info = {}
            for prop in ["default_value", "value", "type", "description", "unit", 
                         "minimum_value", "maximum_value", "enabled", "settable_per_mesh",
                         "settable_per_extruder", "options"]:
                if prop in node:
                    setting_info[prop] = node[prop]
            if setting_info:
                settings[path] = setting_info
    
    if "settings" in def_data:
        for category_key, category in def_data["settings"].items():
            recurse(category, category_key)
    
    if "overrides" in def_data:
        for key, override in def_data["overrides"].items():
            if key not in settings:
                settings[key] = {}
            settings[key].update(override)
            settings[key]["_source"] = def_data.get("_filename", "unknown")
    
    return settings


# =============================================================================
# Post-Processing for Human-Readable Output
# =============================================================================

def humanize_output(data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Post-process extracted data for human readability:
    - Split semicolon-delimited strings into arrays
    - Format G-code as readable multiline
    - Clean up nested structures
    """
    
    # Keys that are semicolon-delimited lists
    SEMICOLON_LIST_KEYS = {
        "visible_settings",
        "categories_expanded", 
        "custom_visible_settings",
        "recent_files",
        "expanded_brands",
    }
    
    # Keys that are newline-delimited (G-code)
    GCODE_KEYS = {
        "machine_start_gcode",
        "machine_end_gcode",
        "start_gcode",
        "end_gcode",
    }
    
    def process_value(key: str, value: Any) -> Any:
        """Process a single value based on its key."""
        if value is None:
            return value
            
        # Handle semicolon-delimited lists
        if key in SEMICOLON_LIST_KEYS and isinstance(value, str):
            items = [item.strip() for item in value.split(";") if item.strip()]
            return sorted(items) if len(items) > 10 else items
        
        # Handle G-code - split into lines for readability
        if key in GCODE_KEYS and isinstance(value, str):
            # Replace literal \n with actual newlines, then split
            cleaned = value.replace("\\n", "\n").replace("\\t", "\t")
            lines = [line for line in cleaned.split("\n")]
            return lines if len(lines) > 1 else value
        
        # Handle comma-separated coordinate lists (e.g., machine_head_with_fans_polygon)
        if key == "machine_head_with_fans_polygon" and isinstance(value, str):
            try:
                import ast
                return ast.literal_eval(value)
            except:
                return value
        
        return value
    
    def process_dict(d: Dict[str, Any], depth: int = 0) -> Dict[str, Any]:
        """Recursively process a dictionary."""
        result = {}
        for key, value in d.items():
            if isinstance(value, dict):
                result[key] = process_dict(value, depth + 1)
            elif isinstance(value, list):
                result[key] = [
                    process_dict(item, depth + 1) if isinstance(item, dict) else item
                    for item in value
                ]
            else:
                result[key] = process_value(key, value)
        return result
    
    return process_dict(data)


def create_summary_section(data: Dict[str, Any]) -> Dict[str, Any]:
    """Create a human-friendly summary section at the top of the output."""
    summary = {
        "_note": "This section provides a quick overview. Full details below.",
    }
    
    # Machine info
    if "machine" in data:
        machine = data["machine"]
        summary["machine_name"] = data.get("metadata", {}).get("machine", "Unknown")
        summary["inheritance"] = " → ".join(
            item["name"] for item in machine.get("inheritance_chain", [])
        )
        summary["total_settings"] = len(machine.get("effective_settings", {}))
    
    # G-code summary
    if "gcode" in data:
        gcode = data["gcode"]
        summary["gcode_source"] = gcode.get("source", "Unknown")
        start = gcode.get("start_gcode", "")
        end = gcode.get("end_gcode", "")
        summary["start_gcode_lines"] = len(start) if isinstance(start, list) else start.count("\n") + 1
        summary["end_gcode_lines"] = len(end) if isinstance(end, list) else end.count("\n") + 1
    
    # Quality profiles
    if "quality_builtin" in data:
        summary["builtin_qualities"] = list(data["quality_builtin"].keys())
    
    if "quality_custom" in data:
        summary["custom_profiles"] = list(data["quality_custom"].keys())
    
    # Plugins
    if "plugins" in data:
        summary["plugins"] = [
            f"{info.get('name', pid)} v{info.get('version', '?')}"
            for pid, info in data["plugins"].items()
            if not pid.startswith("_")
        ]
    
    return summary


def extract_key_settings(data: Dict[str, Any]) -> Dict[str, Any]:
    """Extract the most commonly-referenced settings into a quick-reference section."""
    key_settings = {}
    
    # Settings most people care about
    IMPORTANT_SETTINGS = [
        # Layer
        "layer_height", "layer_height_0",
        # Walls
        "wall_thickness", "wall_line_count",
        # Top/Bottom
        "top_layers", "bottom_layers", "top_bottom_thickness",
        # Infill
        "infill_sparse_density", "infill_pattern",
        # Speed
        "speed_print", "speed_infill", "speed_wall", "speed_wall_0", "speed_wall_x",
        "speed_topbottom", "speed_travel", "speed_layer_0",
        # Retraction
        "retraction_enable", "retraction_amount", "retraction_speed",
        "retraction_hop_enabled", "retraction_hop",
        # Temperature (from quality/material)
        "material_print_temperature", "material_bed_temperature",
        # Cooling
        "cool_fan_speed", "cool_fan_speed_min", "cool_fan_speed_max",
        # Support
        "support_enable", "support_type", "support_structure",
        # Adhesion
        "adhesion_type", "skirt_line_count", "brim_width",
        # Machine
        "machine_width", "machine_depth", "machine_height",
        "machine_heated_bed", "machine_nozzle_size",
    ]
    
    effective = data.get("machine", {}).get("effective_settings", {})
    def_changes = data.get("machine", {}).get("definition_changes", {}).get("values", {})
    
    for setting in IMPORTANT_SETTINGS:
        if setting in def_changes:
            key_settings[setting] = {
                "value": def_changes[setting],
                "source": "your_customizations"
            }
        elif setting in effective:
            info = effective[setting]
            value = info.get("effective_value") or info.get("value") or info.get("default_value")
            key_settings[setting] = {
                "value": value,
                "source": info.get("_sources", ["unknown"])[-1] if "_sources" in info else "default"
            }
    
    return key_settings


# =============================================================================
# Core Extraction Logic
# =============================================================================

class CuraExtractor:
    """Main extraction engine."""
    
    def __init__(self, install_path: str, appdata_path: str, log_callback=None):
        self.install_path = Path(install_path)
        self.appdata_path = Path(appdata_path)
        self.log = log_callback or print
        
        # Discovered data
        self.machines: List[str] = []
        self.custom_profiles: List[str] = []
        self.materials: List[str] = []
        self.cura_version: str = "unknown"
    
    def validate_paths(self) -> Tuple[bool, List[str]]:
        """Validate that paths exist and contain expected structure."""
        errors = []
        
        # Check install path
        resources = self.install_path / "share" / "cura" / "resources"
        if not resources.exists():
            errors.append(f"Install path missing resources: {resources}")
        
        definitions = resources / "definitions" if resources.exists() else None
        if definitions and not (definitions / "fdmprinter.def.json").exists():
            errors.append("Missing fdmprinter.def.json in definitions")
        
        # Check appdata path
        if not self.appdata_path.exists():
            errors.append(f"AppData path does not exist: {self.appdata_path}")
        
        if not (self.appdata_path / "cura.cfg").exists():
            errors.append("Missing cura.cfg in AppData")
        
        return len(errors) == 0, errors
    
    def discover(self) -> Dict[str, Any]:
        """Discover available machines, profiles, and materials."""
        result = {
            "machines": [],
            "custom_profiles": [],
            "builtin_qualities": [],
            "materials": [],
            "plugins": [],
        }
        
        # Extract version from path
        match = re.search(r'(\d+\.\d+\.?\d*)', str(self.install_path))
        if match:
            self.cura_version = match.group(1)
        
        # Discover machines from machine_instances
        machine_dir = self.appdata_path / "machine_instances"
        if machine_dir.exists():
            for f in machine_dir.glob("*.global.cfg"):
                # Decode URL-encoded filename
                name = unquote(f.stem.replace(".global", ""))
                result["machines"].append(name)
                self.machines.append(name)
        
        # Discover custom quality profiles
        quality_changes = self.appdata_path / "quality_changes"
        if quality_changes.exists():
            seen = set()
            for f in quality_changes.glob("*.inst.cfg"):
                cfg = parse_cfg_file(f)
                name = cfg.get("general", {}).get("name", f.stem)
                if name not in seen:
                    result["custom_profiles"].append(name)
                    seen.add(name)
            self.custom_profiles = list(seen)
        
        # Discover built-in quality profiles
        quality_dir = self.install_path / "share" / "cura" / "resources" / "quality" / "creality" / "base"
        if quality_dir.exists():
            for f in quality_dir.glob("base_global_*.inst.cfg"):
                cfg = parse_cfg_file(f)
                name = cfg.get("general", {}).get("name", f.stem)
                result["builtin_qualities"].append(name)
        
        # Discover materials
        materials_dir = self.install_path / "share" / "cura" / "resources" / "materials"
        if materials_dir.exists():
            for f in list(materials_dir.glob("*.xml.fdm_material"))[:20]:  # Limit for performance
                result["materials"].append(f.stem.replace(".xml", ""))
        
        # Discover plugins
        packages_file = self.appdata_path / "packages.json"
        if packages_file.exists():
            try:
                with open(packages_file, 'r', encoding='utf-8') as f:
                    packages = json.load(f)
                for pkg_id, pkg_info in packages.get("installed", {}).items():
                    name = pkg_info.get("package_info", {}).get("display_name", pkg_id)
                    result["plugins"].append(name)
            except:
                pass
        
        return result
    
    def extract_all(self, machine_name: str, options: Dict[str, bool]) -> Dict[str, Any]:
        """
        Extract all requested data for a specific machine.
        
        Options:
            preferences: bool
            machine_settings: bool
            gcode: bool
            quality_builtin: bool
            quality_custom: bool
            materials: bool
            plugins: bool
        """
        output = {
            "metadata": {
                "cura_version": self.cura_version,
                "extracted_at": datetime.now().isoformat(),
                "machine": machine_name,
                "extractor_version": __version__,
            }
        }
        
        self.log(f"Starting extraction for machine: {machine_name}")
        
        # 1. Preferences
        if options.get("preferences", True):
            self.log("  → Extracting preferences...")
            output["preferences"] = self._extract_preferences()
        
        # 2. Machine settings (includes definition chain)
        if options.get("machine_settings", True):
            self.log("  → Extracting machine settings...")
            output["machine"] = self._extract_machine(machine_name)
        
        # 3. G-code (Start/End)
        if options.get("gcode", True):
            self.log("  → Extracting G-code...")
            output["gcode"] = self._extract_gcode(machine_name)
        
        # 4. Extruder settings
        if options.get("machine_settings", True):
            self.log("  → Extracting extruder settings...")
            output["extruders"] = self._extract_extruders(machine_name)
        
        # 5. Built-in quality profiles
        if options.get("quality_builtin", True):
            self.log("  → Extracting built-in quality profiles...")
            output["quality_builtin"] = self._extract_builtin_qualities()
        
        # 6. Custom quality profiles
        if options.get("quality_custom", True):
            self.log("  → Extracting custom quality profiles...")
            output["quality_custom"] = self._extract_custom_qualities()
        
        # 7. Plugins
        if options.get("plugins", True):
            self.log("  → Extracting plugins...")
            output["plugins"] = self._extract_plugins()
        
        # Add summary and key settings sections
        self.log("  → Generating summary...")
        summary = create_summary_section(output)
        key_settings = extract_key_settings(output)
        
        # Reorder to put summary first
        ordered_output = {
            "_summary": summary,
            "_key_settings": key_settings,
        }
        ordered_output.update(output)
        
        self.log("Extraction complete!")
        return ordered_output
    
    def _extract_preferences(self) -> Dict[str, Any]:
        """Extract cura.cfg preferences."""
        cfg_path = self.appdata_path / "cura.cfg"
        return parse_cfg_file(cfg_path)
    
    def _extract_machine(self, machine_name: str) -> Dict[str, Any]:
        """Extract machine configuration with full inheritance chain."""
        result = {
            "inheritance_chain": [],
            "effective_settings": {},
            "definition_changes": {},
        }
        
        # Find machine instance file
        machine_dir = self.appdata_path / "machine_instances"
        machine_file = None
        for f in machine_dir.glob("*.global.cfg"):
            name = unquote(f.stem.replace(".global", ""))
            if name == machine_name:
                machine_file = f
                break
        
        if not machine_file:
            result["_error"] = f"Machine not found: {machine_name}"
            return result
        
        # Parse machine instance
        machine_cfg = parse_cfg_file(machine_file)
        result["instance"] = machine_cfg
        
        # Get container stack
        containers = machine_cfg.get("containers", {})
        result["container_stack"] = containers
        
        # Find definition changes (layer 6 - has G-code!)
        settings_name = containers.get("6", "")
        if settings_name:
            def_changes_dir = self.appdata_path / "definition_changes"
            for f in def_changes_dir.glob("*.inst.cfg"):
                cfg = parse_cfg_file(f)
                if cfg.get("general", {}).get("name", "") == settings_name:
                    result["definition_changes"] = cfg
                    break
        
        # Build inheritance chain
        definitions_dir = self.install_path / "share" / "cura" / "resources" / "definitions"
        
        # Find the base definition (layer 7)
        base_def_name = containers.get("7", "creality_ender3pro")
        def_file = definitions_dir / f"{base_def_name}.def.json"
        
        chain = []
        current_def = base_def_name
        while current_def:
            def_path = definitions_dir / f"{current_def}.def.json"
            if def_path.exists():
                def_data = parse_def_json(def_path)
                chain.append({
                    "name": current_def,
                    "file": str(def_path),
                    "inherits": def_data.get("inherits"),
                })
                current_def = def_data.get("inherits")
            else:
                break
        
        result["inheritance_chain"] = chain
        
        # Extract effective settings from chain (bottom-up)
        effective = {}
        for def_info in reversed(chain):
            def_path = Path(def_info["file"])
            def_data = parse_def_json(def_path)
            
            # Extract settings from this definition
            settings = extract_settings_from_def(def_data)
            for key, value in settings.items():
                if key not in effective:
                    effective[key] = {"_sources": []}
                effective[key].update(value)
                effective[key]["_sources"].append(def_info["name"])
        
        # Apply definition_changes overrides
        if "values" in result.get("definition_changes", {}):
            for key, value in result["definition_changes"]["values"].items():
                if key not in effective:
                    effective[key] = {"_sources": []}
                effective[key]["effective_value"] = value
                effective[key]["_sources"].append("definition_changes")
        
        result["effective_settings"] = effective
        return result
    
    def _extract_gcode(self, machine_name: str) -> Dict[str, str]:
        """Extract Start and End G-code."""
        result = {
            "start_gcode": "",
            "end_gcode": "",
            "source": "unknown",
        }
        
        # First check definition_changes (user customizations)
        def_changes_dir = self.appdata_path / "definition_changes"
        for f in def_changes_dir.glob("*_settings.inst.cfg"):
            if machine_name.lower().replace(" ", "_") in f.name.lower().replace("+", "_"):
                cfg = parse_cfg_file(f)
                values = cfg.get("values", {})
                if "machine_start_gcode" in values:
                    result["start_gcode"] = values["machine_start_gcode"]
                    result["source"] = str(f)
                if "machine_end_gcode" in values:
                    result["end_gcode"] = values["machine_end_gcode"]
                    result["source"] = str(f)
                break
        
        # If not found, fall back to definition chain
        if not result["start_gcode"]:
            definitions_dir = self.install_path / "share" / "cura" / "resources" / "definitions"
            for def_name in ["creality_ender3pro", "creality_base", "fdmprinter"]:
                def_path = definitions_dir / f"{def_name}.def.json"
                if def_path.exists():
                    def_data = parse_def_json(def_path)
                    overrides = def_data.get("overrides", {})
                    settings = def_data.get("settings", {}).get("machine_settings", {}).get("children", {})
                    
                    if "machine_start_gcode" in overrides:
                        result["start_gcode"] = overrides["machine_start_gcode"].get("default_value", "")
                        result["source"] = str(def_path)
                    if "machine_end_gcode" in overrides:
                        result["end_gcode"] = overrides["machine_end_gcode"].get("default_value", "")
                    
                    if result["start_gcode"]:
                        break
        
        return result
    
    def _extract_extruders(self, machine_name: str) -> Dict[str, Any]:
        """Extract extruder configurations."""
        result = {}
        
        extruder_dir = self.appdata_path / "extruders"
        if not extruder_dir.exists():
            return result
        
        for f in extruder_dir.glob("*.extruder.cfg"):
            cfg = parse_cfg_file(f)
            metadata = cfg.get("metadata", {})
            
            # Check if this extruder belongs to our machine
            if metadata.get("machine", "") == machine_name or machine_name in str(f):
                position = metadata.get("position", "0")
                result[f"extruder_{position}"] = cfg
                
                # Get extruder settings
                settings_name = cfg.get("containers", {}).get("6", "")
                if settings_name:
                    for sf in (self.appdata_path / "definition_changes").glob("*.inst.cfg"):
                        scfg = parse_cfg_file(sf)
                        if scfg.get("general", {}).get("name", "") == settings_name:
                            result[f"extruder_{position}_settings"] = scfg
                            break
        
        return result
    
    def _extract_builtin_qualities(self) -> Dict[str, Any]:
        """Extract built-in quality profiles."""
        result = {}
        
        quality_dir = self.install_path / "share" / "cura" / "resources" / "quality" / "creality" / "base"
        if not quality_dir.exists():
            return result
        
        for f in quality_dir.glob("base_global_*.inst.cfg"):
            cfg = parse_cfg_file(f)
            name = cfg.get("general", {}).get("name", f.stem)
            quality_type = cfg.get("metadata", {}).get("quality_type", "unknown")
            result[quality_type] = {
                "name": name,
                "file": str(f),
                "settings": cfg.get("values", {}),
            }
        
        return result
    
    def _extract_custom_qualities(self) -> Dict[str, Any]:
        """Extract custom quality profiles from AppData."""
        result = {}
        
        quality_dir = self.appdata_path / "quality_changes"
        if not quality_dir.exists():
            return result
        
        for f in quality_dir.glob("*.inst.cfg"):
            cfg = parse_cfg_file(f)
            name = cfg.get("general", {}).get("name", f.stem)
            
            # Group by profile name (there may be global + per-extruder files)
            if name not in result:
                result[name] = {
                    "files": [],
                    "settings": {},
                }
            
            result[name]["files"].append(str(f))
            
            # Merge settings
            if "values" in cfg:
                result[name]["settings"].update(cfg["values"])
            
            # Include metadata
            if "metadata" in cfg:
                result[name]["metadata"] = cfg["metadata"]
        
        return result
    
    def _extract_plugins(self) -> Dict[str, Any]:
        """Extract installed plugins."""
        result = {}
        
        packages_file = self.appdata_path / "packages.json"
        if not packages_file.exists():
            return result
        
        try:
            with open(packages_file, 'r', encoding='utf-8') as f:
                packages = json.load(f)
            
            for pkg_id, pkg_info in packages.get("installed", {}).items():
                info = pkg_info.get("package_info", {})
                result[pkg_id] = {
                    "name": info.get("display_name", pkg_id),
                    "version": info.get("package_version", "unknown"),
                    "description": info.get("description", ""),
                    "author": info.get("author", {}).get("display_name", "unknown"),
                    "website": info.get("website", ""),
                }
        except Exception as e:
            result["_error"] = str(e)
        
        return result


# =============================================================================
# GUI Application
# =============================================================================

class CuraExtractorGUI:
    """Tkinter GUI for the extractor."""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(f"Cura Profile Extractor v{__version__}")
        self.root.geometry("900x700")
        self.root.minsize(700, 500)
        
        self.extractor: Optional[CuraExtractor] = None
        self.discovered: Dict[str, Any] = {}
        
        self._create_widgets()
        self._load_defaults()
    
    def _create_widgets(self):
        """Build the GUI."""
        # Main container with padding
        main = ttk.Frame(self.root, padding="10")
        main.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)
        
        row = 0
        
        # === Path Section ===
        ttk.Label(main, text="Paths", font=("", 11, "bold")).grid(row=row, column=0, sticky="w", pady=(0, 5))
        row += 1
        
        # Cura Install Path
        ttk.Label(main, text="Cura Install:").grid(row=row, column=0, sticky="w")
        self.install_path = ttk.Entry(main, width=60)
        self.install_path.grid(row=row, column=1, sticky="ew", padx=5)
        ttk.Button(main, text="Browse", command=self._browse_install).grid(row=row, column=2)
        row += 1
        
        # Cura AppData Path
        ttk.Label(main, text="Cura AppData:").grid(row=row, column=0, sticky="w")
        self.appdata_path = ttk.Entry(main, width=60)
        self.appdata_path.grid(row=row, column=1, sticky="ew", padx=5)
        ttk.Button(main, text="Browse", command=self._browse_appdata).grid(row=row, column=2)
        row += 1
        
        # Detect & Verify Button
        ttk.Button(main, text="Detect & Verify Paths", command=self._detect_and_verify).grid(
            row=row, column=0, columnspan=3, pady=10
        )
        row += 1
        
        ttk.Separator(main, orient="horizontal").grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        row += 1
        
        # === Machine Selection ===
        ttk.Label(main, text="Machine", font=("", 11, "bold")).grid(row=row, column=0, sticky="w", pady=(0, 5))
        row += 1
        
        ttk.Label(main, text="Select Machine:").grid(row=row, column=0, sticky="w")
        self.machine_var = tk.StringVar()
        self.machine_combo = ttk.Combobox(main, textvariable=self.machine_var, state="disabled", width=50)
        self.machine_combo.grid(row=row, column=1, sticky="w", padx=5)
        row += 1
        
        ttk.Separator(main, orient="horizontal").grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        row += 1
        
        # === Extraction Options ===
        ttk.Label(main, text="Extract Options", font=("", 11, "bold")).grid(row=row, column=0, sticky="w", pady=(0, 5))
        row += 1
        
        options_frame = ttk.Frame(main)
        options_frame.grid(row=row, column=0, columnspan=3, sticky="w")
        
        self.opt_preferences = tk.BooleanVar(value=True)
        self.opt_machine = tk.BooleanVar(value=True)
        self.opt_gcode = tk.BooleanVar(value=True)
        self.opt_quality_builtin = tk.BooleanVar(value=True)
        self.opt_quality_custom = tk.BooleanVar(value=True)
        self.opt_plugins = tk.BooleanVar(value=True)
        
        opts = [
            ("Preferences", self.opt_preferences),
            ("Machine Settings", self.opt_machine),
            ("Start/End G-code", self.opt_gcode),
            ("Built-in Qualities", self.opt_quality_builtin),
            ("Custom Profiles", self.opt_quality_custom),
            ("Plugins", self.opt_plugins),
        ]
        
        for i, (label, var) in enumerate(opts):
            cb = ttk.Checkbutton(options_frame, text=label, variable=var, state="disabled")
            cb.grid(row=i // 3, column=i % 3, sticky="w", padx=10, pady=2)
        
        self.option_checkboxes = options_frame.winfo_children()
        row += 1
        
        ttk.Separator(main, orient="horizontal").grid(row=row, column=0, columnspan=3, sticky="ew", pady=10)
        row += 1
        
        # === Action Buttons ===
        btn_frame = ttk.Frame(main)
        btn_frame.grid(row=row, column=0, columnspan=3, pady=10)
        
        self.dry_run_btn = ttk.Button(btn_frame, text="Dry Run (Preview)", command=self._dry_run, state="disabled")
        self.dry_run_btn.pack(side="left", padx=5)
        
        self.extract_btn = ttk.Button(btn_frame, text="Extract All!", command=self._extract, state="disabled")
        self.extract_btn.pack(side="left", padx=5)
        
        ttk.Button(btn_frame, text="Clear Log", command=self._clear_log).pack(side="left", padx=5)
        row += 1
        
        # === Log Output ===
        ttk.Label(main, text="Log", font=("", 11, "bold")).grid(row=row, column=0, sticky="w", pady=(0, 5))
        row += 1
        
        self.log_text = scrolledtext.ScrolledText(main, height=15, width=100, state="disabled", 
                                                   font=("Consolas", 9))
        self.log_text.grid(row=row, column=0, columnspan=3, sticky="nsew", pady=5)
        main.rowconfigure(row, weight=1)
        row += 1
        
        # === Status Bar ===
        self.status_var = tk.StringVar(value="Ready. Click 'Detect & Verify Paths' to begin.")
        ttk.Label(main, textvariable=self.status_var, relief="sunken", anchor="w").grid(
            row=row, column=0, columnspan=3, sticky="ew", pady=(5, 0)
        )
    
    def _load_defaults(self):
        """Load default paths."""
        install, appdata = get_default_paths()
        self.install_path.insert(0, install)
        self.appdata_path.insert(0, appdata)
    
    def _browse_install(self):
        """Browse for Cura install directory."""
        path = filedialog.askdirectory(title="Select Cura Installation Directory")
        if path:
            self.install_path.delete(0, tk.END)
            self.install_path.insert(0, path)
    
    def _browse_appdata(self):
        """Browse for Cura AppData directory."""
        path = filedialog.askdirectory(title="Select Cura AppData Directory")
        if path:
            self.appdata_path.delete(0, tk.END)
            self.appdata_path.insert(0, path)
    
    def _log(self, message: str):
        """Append message to log."""
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")
        self.root.update()
    
    def _clear_log(self):
        """Clear the log."""
        self.log_text.config(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state="disabled")
    
    def _detect_and_verify(self):
        """Validate paths and discover available options."""
        self._clear_log()
        self._log("=" * 60)
        self._log("Detecting and verifying Cura installation...")
        self._log("=" * 60)
        
        install = self.install_path.get().strip()
        appdata = self.appdata_path.get().strip()
        
        if not install or not appdata:
            self._log("ERROR: Please provide both paths.")
            self.status_var.set("Error: Missing paths")
            return
        
        self.extractor = CuraExtractor(install, appdata, self._log)
        
        # Validate
        valid, errors = self.extractor.validate_paths()
        if not valid:
            for err in errors:
                self._log(f"ERROR: {err}")
            self.status_var.set("Validation failed - check log")
            return
        
        self._log("✓ Paths validated successfully")
        
        # Discover
        self._log("\nDiscovering available configurations...")
        self.discovered = self.extractor.discover()
        
        self._log(f"  Cura Version: {self.extractor.cura_version}")
        self._log(f"  Machines found: {len(self.discovered['machines'])}")
        for m in self.discovered['machines']:
            self._log(f"    - {m}")
        
        self._log(f"  Custom profiles: {len(self.discovered['custom_profiles'])}")
        for p in self.discovered['custom_profiles']:
            self._log(f"    - {p}")
        
        self._log(f"  Built-in qualities: {len(self.discovered['builtin_qualities'])}")
        self._log(f"  Plugins: {len(self.discovered['plugins'])}")
        
        # Enable controls
        if self.discovered['machines']:
            self.machine_combo['values'] = self.discovered['machines']
            self.machine_combo.current(0)
            self.machine_combo.config(state="readonly")
            
            for cb in self.option_checkboxes:
                cb.config(state="normal")
            
            self.dry_run_btn.config(state="normal")
            self.extract_btn.config(state="normal")
            
            self.status_var.set(f"Ready. Found {len(self.discovered['machines'])} machine(s).")
            self._log("\n✓ Discovery complete. Select options and click 'Extract All!'")
        else:
            self.status_var.set("No machines found - check paths")
            self._log("\nERROR: No machines found in AppData")
    
    def _get_options(self) -> Dict[str, bool]:
        """Get current option selections."""
        return {
            "preferences": self.opt_preferences.get(),
            "machine_settings": self.opt_machine.get(),
            "gcode": self.opt_gcode.get(),
            "quality_builtin": self.opt_quality_builtin.get(),
            "quality_custom": self.opt_quality_custom.get(),
            "plugins": self.opt_plugins.get(),
        }
    
    def _dry_run(self):
        """Preview extraction without saving."""
        if not self.extractor:
            return
        
        machine = self.machine_var.get()
        if not machine:
            messagebox.showwarning("Warning", "Please select a machine")
            return
        
        self._log("\n" + "=" * 60)
        self._log("DRY RUN - Preview Only (no file written)")
        self._log("=" * 60)
        
        options = self._get_options()
        self._log(f"\nOptions selected:")
        for k, v in options.items():
            self._log(f"  {k}: {'Yes' if v else 'No'}")
        
        self._log(f"\nExtracting for machine: {machine}")
        
        try:
            result = self.extractor.extract_all(machine, options)
            
            # Summary
            self._log("\n" + "-" * 40)
            self._log("Extraction Summary:")
            self._log("-" * 40)
            
            if "preferences" in result:
                self._log(f"  Preferences sections: {len(result['preferences']) - 2}")  # minus _filepath, _filename
            
            if "machine" in result:
                chain_len = len(result['machine'].get('inheritance_chain', []))
                settings_count = len(result['machine'].get('effective_settings', {}))
                self._log(f"  Inheritance chain depth: {chain_len}")
                self._log(f"  Effective settings: {settings_count}")
            
            if "gcode" in result:
                start_len = len(result['gcode'].get('start_gcode', ''))
                end_len = len(result['gcode'].get('end_gcode', ''))
                self._log(f"  Start G-code: {start_len} chars")
                self._log(f"  End G-code: {end_len} chars")
            
            if "quality_builtin" in result:
                self._log(f"  Built-in qualities: {len(result['quality_builtin'])}")
            
            if "quality_custom" in result:
                self._log(f"  Custom profiles: {len(result['quality_custom'])}")
            
            if "plugins" in result:
                self._log(f"  Plugins: {len(result['plugins'])}")
            
            self._log("\n✓ Dry run complete. Click 'Extract All!' to save to file.")
            self.status_var.set("Dry run complete")
            
        except Exception as e:
            self._log(f"\nERROR: {e}")
            self.status_var.set("Dry run failed - check log")
    
    def _extract(self):
        """Run full extraction and save to file."""
        if not self.extractor:
            return
        
        machine = self.machine_var.get()
        if not machine:
            messagebox.showwarning("Warning", "Please select a machine")
            return
        
        # Ask for save location
        default_name = f"cura_profile_{machine.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        filepath = filedialog.asksaveasfilename(
            title="Save Extracted Profile",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialfile=default_name
        )
        
        if not filepath:
            return
        
        self._log("\n" + "=" * 60)
        self._log("FULL EXTRACTION")
        self._log("=" * 60)
        
        options = self._get_options()
        
        try:
            result = self.extractor.extract_all(machine, options)
            
            # Apply human-friendly formatting
            self._log("  → Formatting for readability...")
            result = humanize_output(result)
            
            # Save to file
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(result, f, indent=2, ensure_ascii=False)
            
            self._log(f"\n✓ Profile saved to: {filepath}")
            self._log(f"  File size: {os.path.getsize(filepath):,} bytes")
            
            self.status_var.set(f"Saved: {filepath}")
            
            # Offer to open
            if messagebox.askyesno("Success", f"Profile saved to:\n{filepath}\n\nOpen in default viewer?"):
                os.startfile(filepath)
            
        except Exception as e:
            self._log(f"\nERROR: {e}")
            self.status_var.set("Extraction failed - check log")
            messagebox.showerror("Error", f"Extraction failed:\n{e}")
    
    def run(self):
        """Start the GUI."""
        self.root.mainloop()


# =============================================================================
# CLI Interface
# =============================================================================

def run_cli(args):
    """Run in CLI mode."""
    print(f"Cura Profile Extractor v{__version__}")
    print("=" * 50)
    
    # Get paths
    install_path = args.install or find_cura_install_path()
    appdata_path = args.appdata or find_cura_appdata_path()
    
    if not install_path:
        print("ERROR: Could not detect Cura install path. Use --install to specify.")
        return 1
    if not appdata_path:
        print("ERROR: Could not detect Cura AppData path. Use --appdata to specify.")
        return 1
    
    print(f"Install: {install_path}")
    print(f"AppData: {appdata_path}")
    
    extractor = CuraExtractor(str(install_path), str(appdata_path))
    
    # Validate
    valid, errors = extractor.validate_paths()
    if not valid:
        print("\nValidation errors:")
        for err in errors:
            print(f"  - {err}")
        return 1
    
    # Discover
    print("\nDiscovering configurations...")
    discovered = extractor.discover()
    
    print(f"  Machines: {discovered['machines']}")
    print(f"  Custom profiles: {discovered['custom_profiles']}")
    
    # Select machine
    machine = args.machine
    if not machine:
        if discovered['machines']:
            machine = discovered['machines'][0]
            print(f"\nUsing first machine: {machine}")
        else:
            print("ERROR: No machines found")
            return 1
    
    # Extract
    options = {
        "preferences": not args.no_preferences,
        "machine_settings": not args.no_machine,
        "gcode": not args.no_gcode,
        "quality_builtin": not args.no_builtin,
        "quality_custom": not args.no_custom,
        "plugins": not args.no_plugins,
    }
    
    print(f"\nExtracting for: {machine}")
    result = extractor.extract_all(machine, options)
    
    # Apply human-friendly formatting unless --raw
    if not args.raw:
        print("Formatting for readability...")
        result = humanize_output(result)
    else:
        print("Skipping formatting (--raw mode)")
    
    # Save
    output_file = args.output or f"cura_profile_{machine.replace(' ', '_')}.json"
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, indent=2, ensure_ascii=False)
    
    print(f"\n✓ Saved to: {output_file}")
    print(f"  Size: {os.path.getsize(output_file):,} bytes")
    
    return 0


# =============================================================================
# Entry Point
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Extract ALL Cura settings into a single searchable JSON file.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s                    # Launch GUI (default)
  %(prog)s --cli              # Run in CLI mode with auto-detection
  %(prog)s --cli --machine "Ender3Pro_Sprite_CRTouch" --output my_profile.json
  
For more information, see the header comments in this script.
"""
    )
    
    parser.add_argument("--cli", action="store_true", help="Run in CLI mode instead of GUI")
    parser.add_argument("--install", type=str, help="Cura installation path")
    parser.add_argument("--appdata", type=str, help="Cura AppData path")
    parser.add_argument("--machine", type=str, help="Machine name to extract")
    parser.add_argument("--output", "-o", type=str, help="Output JSON file path")
    parser.add_argument("--no-preferences", action="store_true", help="Skip preferences")
    parser.add_argument("--no-machine", action="store_true", help="Skip machine settings")
    parser.add_argument("--no-gcode", action="store_true", help="Skip G-code")
    parser.add_argument("--no-builtin", action="store_true", help="Skip built-in qualities")
    parser.add_argument("--no-custom", action="store_true", help="Skip custom profiles")
    parser.add_argument("--no-plugins", action="store_true", help="Skip plugins")
    parser.add_argument("--raw", action="store_true", help="Skip human-friendly formatting")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    
    args = parser.parse_args()
    
    if args.cli:
        sys.exit(run_cli(args))
    else:
        app = CuraExtractorGUI()
        app.run()


if __name__ == "__main__":
    main()
