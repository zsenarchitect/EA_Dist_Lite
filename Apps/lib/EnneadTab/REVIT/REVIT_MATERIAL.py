# -*- coding: utf-8 -*-



import REVIT_APPLICATION, REVIT_SELECTION
try:

    from Autodesk.Revit import DB # pyright: ignore
    UIDOC = REVIT_APPLICATION.get_uidoc() 
    DOC = REVIT_APPLICATION.get_doc()
    

    
    
except:
    globals()["UIDOC"] = object()
    globals()["DOC"] = object()


def get_material_by_name(material_name, doc=DOC):
    """Retrieves a material by its name from the document.
    
    Args:
        material_name (str): Name of the material to find
        doc (Document): The Revit document to query. Defaults to active document
        
    Returns:
        Material: The matching material element, or None if not found
    """
    all_materials = DB.FilteredElementCollector(doc).OfClass(DB.Material).ToElements()
    
    for material in all_materials:
        if material.Name == material_name:
            return material
    return None


MATERIAL_MAP_SAMPLE_FOR_YOU_TO_MIMIC = {
    "Curb":{
        "Color": (147, 147, 147),
        "SurfaceForegroundPatternIsSolid": True,
        "SurfaceForegroundPatternColor": (192, 192, 192)
    } ,
    "Road":{
        "Color": (120, 120, 120),
        "SurfaceForegroundPatternIsSolid": True,
        "SurfaceForegroundPatternColor": (150, 150, 150)
    },
    "Grass":{
        "Color": (187, 255, 187),
        "SurfaceForegroundPatternColor": (116, 206, 85),
        "SurfaceBackgroundPatternIsSolid": True,
        "SurfaceBackgroundPatternColor": (238, 252, 233)
    }
}
def update_material_setting(doc, material_map):
    """Updates material properties based on a configuration map.
    
    Material map should follow this structure:
    {
        "Material Name": {
            "Color": (r, g, b),
            "SurfaceForegroundPatternIsSolid": bool,
            "SurfaceForegroundPatternColor": (r, g, b),
            "SurfaceBackgroundPatternIsSolid": bool,
            "SurfaceBackgroundPatternColor": (r, g, b)
        }
    }
    
    Args:
        doc (Document): The Revit document containing materials
        material_map (dict): Configuration dictionary for material properties
        
    Example:
        material_map = {
            "Curb": {
                "Color": (147, 147, 147),
                "SurfaceForegroundPatternIsSolid": True,
                "SurfaceForegroundPatternColor": (192, 192, 192)
            }
        }
    """
    for material_name, material_setting in material_map.items():
        material = get_material_by_name(material_name, doc)
        
        if material is None:
            continue

        if not REVIT_SELECTION.is_changable(material):
            continue

        for setting_name, setting_value in material_setting.items():

            if "IsSolid" in setting_name:
                if setting_value:
                    solid_fill_pattern_id = REVIT_SELECTION.get_solid_fill_pattern_id(doc)
                attr_name = setting_name.replace("IsSolid", "Id")
                setattr(material, attr_name, solid_fill_pattern_id)
                continue
                
            if "color" in setting_name.lower():
                setting_value = DB.Color(setting_value[0], setting_value[1], setting_value[2])
            setattr(material, setting_name, setting_value)

def sanitize_material_name(name):
    """Remove prohibited characters from material name.
    
    Args:
        name (str): Original material name
        
    Returns:
        str: Sanitized material name
    """
    name = name.replace("[imported]", '')
    # Based on Revit error message: "{, }, [, ], |, ;, less-than sign, greater-than sign, ?, `, ~"
    # Also including other common problematic characters: \, :
    prohibited_chars = '{}[]|;<>?`~\\:'
    for char in prohibited_chars:
        name = name.replace(char, '_')
    name = name.strip()
    return name