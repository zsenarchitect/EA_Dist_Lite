from pyrevit import forms
from Autodesk.Revit import DB # pyright:ignore

def pick_category(doc):
    """Displays UI for selecting Revit categories from a predefined list.
    
    Args:
        doc (Document): The Revit document to query categories from
        
    Returns:
        list: Collection of selected Category objects, or None if selection canceled
    """
    cate_list = [("OST_Grids", "Grids"),
            ("OST_Levels", "Levels"),
            ("OST_Rooms", "Rooms"),
            ("OST_Areas", "Areas"),
            ("OST_Furniture", "Furniture"),
            ("OST_Parking", "Parking"),]
    class MyOption(forms.TemplateListItem):
        @property
        def name(self):
            return self.item[1]

    selected_cates = forms.SelectFromList.show([MyOption(cate) for cate in cate_list], 
                                      title = "Select Categorie(s) to bind", 
                                      multiselect = True)
    if not selected_cates:
        return
    
    selected_cate_ids = [getattr(DB.BuiltInCategory , cate[0]) for cate in selected_cates]
    selected_cates = [DB.Category.GetCategory(doc, cate_id) for cate_id in selected_cate_ids]


    return selected_cates
  


class RevitCategory:
    """Wrapper class for Revit Category objects providing enhanced functionality.
    
    Args:
        category (Category): The Revit Category object to wrap
    """
    
    def __init__(self, category):
        self.category = category

    @property
    def sub_category(self):
        """Gets the sub-category if the category is nested.
        
        Returns:
            Category: The sub-category object if present, None otherwise
        """
        if self.category.Parent:
            return self.category
        return None
    
    @property
    def sub_category_name(self):
        """Gets the name of the sub-category.
        
        Returns:
            str: Name of the sub-category if present, empty string otherwise
        """
        if self.sub_category:
            return self.sub_category.Name
        return ""
    
    @property
    def root_category(self):
        """Gets the root category, traversing up the parent hierarchy.
        
        Returns:
            Category: The root category object
        """
        if self.category.Parent:
            return self.category.Parent
        return self.category

    @property
    def root_category_name(self):
        """Gets the name of the root category.
        
        Returns:
            str: Name of the root category
        """
        return self.root_category.Name

    @property
    def pretty_name(self):
        """Formats a human-readable category name including hierarchy.
        
        Returns:
            str: Formatted string showing root and sub-category if present
        """
        if not self.sub_category:
            return "[{}]".format(self.root_category_name)
        
        return "[{}]: {}".format(self.root_category_name, self.sub_category_name)

def get_or_create_subcategory_with_material(doc, subc_name):
    """
    Get or create a subcategory by name under the family document's root category, and assign a mapped material if available.
    Args:
        doc: The Revit family document
        subc_name: Name of the subcategory to get or create
        mapping_key: Optional key to use for material mapping lookup (defaults to subc_name)
    Returns:
        The subcategory (DB.Category) object
    """
    parent_category = doc.OwnerFamily.FamilyCategory
    subCs = parent_category.SubCategories
    for subC in subCs:
        if subC.Name == subc_name:
            return subC
    # Create if not found
    new_subc = doc.Settings.Categories.NewSubcategory(parent_category, subc_name)
    # Assign material if mapping exists
    try:
        from EnneadTab.REVIT import REVIT_MATERIAL
        from EnneadTab import DATA_FILE
        recent_out_data = DATA_FILE.get_data("rhino2revit_out_paths")
        if recent_out_data and isinstance(recent_out_data.get("layer_material_mapping"), dict):
            # Try to find material data using the subcategory name as key
            mat_data = recent_out_data["layer_material_mapping"].get(subc_name)
            if not mat_data:
                # If not found, try with sanitized name (in case the mapping uses sanitized keys)
                import re
                sanitized_key = subc_name.replace('.', '_')
                prohibited = r'[\\:\{\}\[\]\|;<>\\?`~]'
                sanitized_key = re.sub(prohibited, '', sanitized_key)
                mat_data = recent_out_data["layer_material_mapping"].get(sanitized_key)
                if mat_data:
                    print("Found material data using sanitized key: '{}' -> '{}'".format(subc_name, sanitized_key))
            
            if mat_data and isinstance(mat_data, dict):
                mat_name = mat_data.get("material_name")
                mat_color = mat_data.get("material_color")
                print("Found material data for subcategory '{}': material_name='{}'".format(subc_name, mat_name))
            else:
                print("No material data found for subcategory '{}'".format(subc_name))
        else:
            print("No valid layer_material_mapping found in recent_out_data")
        if mat_name:
            # Sanitize material name before using it
            original_mat_name = mat_name
            mat_name = REVIT_MATERIAL.sanitize_material_name(mat_name)
            if mat_name != original_mat_name:
                print("Material name sanitized: '{}' -> '{}'".format(original_mat_name, mat_name))
            
            print("Looking for material with name: '{}'".format(mat_name))
            material = REVIT_MATERIAL.get_material_by_name(mat_name, doc)
            if material is None and mat_color:
                print("Creating new material with name: '{}'".format(mat_name))
                try:
                    mat_id = DB.Material.Create(doc, mat_name)
                    material = doc.GetElement(mat_id)
                    print("Successfully created material with ID: {}".format(mat_id))
                    # Parse RGB tuple to DB.Color object
                    if mat_color and len(mat_color) == 3:
                        color_obj = DB.Color(mat_color[0], mat_color[1], mat_color[2])
                        material.Color = color_obj
                        # Set solid fill pattern and color for surface foreground
                        from EnneadTab.REVIT import REVIT_SELECTION
                        solid_fill_pattern_id = REVIT_SELECTION.get_solid_fill_pattern_id(doc)
                        material.SurfaceForegroundPatternId = solid_fill_pattern_id
                        material.SurfaceForegroundPatternColor = color_obj
                except Exception as create_error:
                    print("Failed to create material '{}': {}".format(mat_name, str(create_error)))
            elif material is None:
                print("Creating new material with name: '{}' (no color data)".format(mat_name))
                try:
                    mat_id = DB.Material.Create(doc, mat_name)
                    material = doc.GetElement(mat_id)
                    print("Successfully created material with ID: {}".format(mat_id))
                except Exception as create_error:
                    print("Failed to create material '{}': {}".format(mat_name, str(create_error)))
            else:
                print("Found existing material: '{}'".format(mat_name))
            if material:
                print("Assigning material '{}' to subcategory '{}'".format(material.Name, subc_name))
                new_subc.Material = material
    except Exception as e:
        print("[get_or_create_subcategory_with_material] Material assignment failed: {}".format(e))
    return new_subc