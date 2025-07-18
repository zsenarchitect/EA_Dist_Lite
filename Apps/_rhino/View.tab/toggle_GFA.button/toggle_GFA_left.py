# -*- coding: utf-8 -*-
__title__ = "ToggleGFA"
__doc__ = """Toggles Gross Floor Area (GFA) visualization and calculation.

Features:
- Processes layers marked with [GFA] to calculate and display area information
- Real-time updates as geometry changes
- Supports area factor multipliers using \{factor\} syntax in layer names
- Excel data export capabilities
- Automatic unit conversion (mm/m -> SQM, inch/ft -> SQFT)
- Dynamic merging of coplanar surfaces at same elevation
- Support for single surfaces and polysurfaces
- Live comparison of how much is off from target.

Usage:
- Add [GFA] to layer names to include in calculation
- Optional \{factor\} at end of layer name for area multipliers (e.g. \{0.5\})
- Right-click to export to Excel or generate checking surfaces or set target area for each keyword.
"""
__is_popular__ = True


import re
import Rhino # pyright: ignore
import System # pyright: ignore
import scriptcontext as sc # pyright: ignore
import rhinoscriptsyntax as rs # pyright: ignore
import sys
sys.path.append("..\lib")



import time
import random


from EnneadTab import ERROR_HANDLE, LOG, NOTIFICATION, TIME, EXCEL

from EnneadTab.RHINO import RHINO_LAYER, RHINO_OBJ_DATA, RHINO_PROJ_DATA

def try_catch_error(func):
    
    def wrapper(*args, **kwargs):

        try:
            out = func(*args, **kwargs)
            return out
        except Exception as e:
            print ( str(e))
            error = ERROR_HANDLE.get_alternative_traceback()
            print (error)
            if random.random() < 0.1:
                NOTIFICATION.messenger(error)
    return wrapper


class EA_GFA_Conduit(Rhino.Display.DisplayConduit):
    """Display conduit for real-time GFA visualization and calculation.
    
    Monitors document changes and updates area calculations automatically.
    Handles layer changes, object modifications, additions and deletions.
    """
    
    def __init__(self):
        self.data = []
        self.cached_data = []
        sc.sticky["EA_GFA_IS_BAKING"] = False
        sc.sticky["reset_timestamp"] = time.time()
        self.current_objs = get_current_objs()
        
        data = RHINO_PROJ_DATA.get_plugin_data()
        self.target_dict = data.get(RHINO_PROJ_DATA.DocKeys.GFA_TARGET_DICT, {})



       
    @ERROR_HANDLE.try_catch_error()
    def check_doc_update_after_adding(self,sender, e):
        layer = sc.doc.Layers.FindIndex(e.TheObject.Attributes.LayerIndex)
        if layer and layer.FullPath:
            if "[GFA]" not in layer.FullPath:
                return

        # print("added")
        # this is added because single action to any object will trigger 
        # replace geo method, that is a equal to delete and add in a very fast sequence. 
        # So to ensure cached data is cleared after, need to add this. 
        # IF in furutre is there a reanother hook that monitor direct modification or replace geo event, can use that instead.
        self.cached_data = [] 
        self.reset_conduit_data("Document content changed, recalculating...")
  
    @ERROR_HANDLE.try_catch_error()
    def check_doc_updated_after_deleting(self,sender, e):
        layer = sc.doc.Layers.FindIndex(e.TheObject.Attributes.LayerIndex)
        if layer and layer.FullPath:
            if "[GFA]" not in layer.FullPath:
                return
        
        self.reset_conduit_data("Document content changed, recalculating...")

    @ERROR_HANDLE.try_catch_error()
    def check_doc_updated_after_layertable_changed(self,sender, e):
        layer = sc.doc.Layers.FindIndex(e.LayerIndex)
        if layer and layer.FullPath:
            if "[GFA]" not in layer.FullPath:
                return
        self.reset_conduit_data("Layer table changed, recalculating...")

    @ERROR_HANDLE.try_catch_error()
    def check_doc_update_after_modifying(self,sender, e):
        layer = sc.doc.Layers.FindIndex(e.NewAttributes.LayerIndex)
        if layer and layer.FullPath:
            if "[GFA]" not in layer.FullPath:
                return
        self.reset_conduit_data("Obj attribute modifed, recalculating...")
    
    
    def reset_conduit_data(self, note=None):
        
        if time.time() - sc.sticky["reset_timestamp"] < 2:
            return
        if note:
            NOTIFICATION.messenger(note)
        self.cached_data = []
        self.current_objs = get_current_objs()
        sc.sticky["reset_timestamp"] = time.time()
        # print ("cached data is now empty")
        # self.is_reseted = True
        self.target_dict = RHINO_PROJ_DATA.get_plugin_data().get(RHINO_PROJ_DATA.DocKeys.GFA_TARGET_DICT, {})   

    
    def add_hook(self):
        Rhino.RhinoDoc.AddRhinoObject  += self.check_doc_update_after_adding
        Rhino.RhinoDoc.DeleteRhinoObject  += self.check_doc_updated_after_deleting
        Rhino.RhinoDoc.UndeleteRhinoObject   += self.check_doc_update_after_adding
        Rhino.RhinoDoc.ModifyObjectAttributes  += self.check_doc_update_after_modifying
        Rhino.RhinoDoc.LayerTableEvent   += self.check_doc_updated_after_layertable_changed
        
    def remove_hook(self):
        Rhino.RhinoDoc.AddRhinoObject  -= self.check_doc_update_after_adding
        Rhino.RhinoDoc.DeleteRhinoObject  -= self.check_doc_updated_after_deleting
        Rhino.RhinoDoc.UndeleteRhinoObject   -= self.check_doc_update_after_adding
        Rhino.RhinoDoc.ModifyObjectAttributes  -= self.check_doc_update_after_modifying
        Rhino.RhinoDoc.LayerTableEvent   -= self.check_doc_updated_after_layertable_changed

        
    def get_layer_target(self, layer):
        return next((value for key, value in self.target_dict.items() if key in layer), None)

    @try_catch_error
    def PostDrawObjects(self, e):
        # Additional safety check for display object
        if e is None or e.Display is None:
            return
            
        if self.current_objs != get_current_objs():
            self.reset_conduit_data(None)

        if not self.cached_data:
            try:
                self.data = [(x, get_area_and_crv_geo_from_layer(x))   for x in get_schedule_layers()]
                self.cached_data = self.data[:]
            except Exception as ex:
                print("Error getting schedule layers data: {}".format(str(ex)))
                return
        else:
            self.data = self.cached_data

        for data in self.data:
            layer, values = data
            total_area, edges, faces, note = values

            if not rs.IsLayer(layer):
                print ("!!!!!!!!!!!!!!!!!! This layer is no longer existing ....{}".format(layer))
                self.data.remove(data)
                continue

            color = rs.LayerColor(layer)
            thickness = 10

            for edge in edges:
                try:
                    if edge is not None:
                        e.Display.DrawCurve(edge, color, thickness)
                except Exception as ex:
                    print("Error drawing edge in layer {}: {}".format(layer, str(ex)))
                    continue

            for face in faces:
                try:
                    abstract_face = Rhino.Geometry.AreaMassProperties.Compute(face)
                    if abstract_face:
                        area = abstract_face.Area
                        text = convert_area_to_good_unit(area)
                        
                        factor =  self.layer_factor(layer)
                        if factor != 1:
                            text = text + " x {} = {}".format(factor, convert_area_to_good_unit(area * factor))
                        
                        pt3D = abstract_face.Centroid
                        e.Display.DrawDot(pt3D, text, color, System.Drawing.Color.White)
                    else:
                        print ("!!! Check for geo cleanness in layer: " + layer)
                except Exception as ex:
                    print ("!!! Error processing face in layer {}: {}".format(layer, str(ex)))
                    continue

    def layer_factor(self, layer):
        # if layer name contain syntax such as 'abcd{0.5}' or 'xyz{0}', extract 0.5 and 0 as the factor.
        # if no curly bracket is found, return 1.
        pattern = re.compile(r".*{(.*)}")
        match = pattern.search(layer)
        if match:
            try:
                return float(match.group(1))
            except:
                return 1
        return 1
    
    @try_catch_error
    def DrawForeground(self, e):
        # Additional safety check for display object
        if e is None or e.Display is None:
            return





        
        color = rs.CreateColor([87, 85, 83])
        color_hightlight = rs.CreateColor([150, 85, 83])
        backgroud_color = rs.CreateColor([240, 240, 240])
        #color = System.Drawing.Color.Red
        position_X_offset = 20
        position_Y_offset = 40
        size = 40
        
        # Check if viewport exists and has bounds
        if e.Viewport is None or e.Viewport.Bounds is None:
            return
            
        bounds = e.Viewport.Bounds
        pt = Rhino.Geometry.Point2d(bounds.Left + position_X_offset, bounds.Top + position_Y_offset)


        # background_rec = e.Display.DrawRoundedRectangle(Rhino.Geometry.Point2d(pt[0] + 50, pt[1] + 50),
        #                                                 200,
        #                                                 200,
        #                                                 50,
        #                                                 color_hightlight,
        #                                                 3,
        #                                                 backgroud_color)

        
        text = "Ennead GFA Schedule Mode"
        e.Display.Draw2dText(text, color, pt, False, size)
        
        # Safe access to sticky data
        reset_timestamp = sc.sticky.get("reset_timestamp", time.time())
        recent_time_text = TIME.get_formatted_time(reset_timestamp)
        recent_time_text = "Last data cache update: " + recent_time_text
        pt = Rhino.Geometry.Point2d(pt[0], pt[1] + 35)
        e.Display.Draw2dText(recent_time_text, color, pt, False, 10)
        pt = Rhino.Geometry.Point2d(pt[0], pt[1] + 15)
        e.Display.Draw2dText("Including bracket [GFA] in layer name to allow scheduling.(Do not begin layer name with '[' bracket.)", color, pt, False, 10)
        pt = Rhino.Geometry.Point2d(pt[0], pt[1] + 10)
        e.Display.Draw2dText("Including curly bracket {factor} at end of layer name to allow factor multiply, such as {0.5} or {0} or whatever.", color_hightlight, pt, False, 10)
        pt = Rhino.Geometry.Point2d(pt[0], pt[1] + 10)
        e.Display.Draw2dText("Right click toggle button to bake data to Excel and/or generate checking srf.", color, pt, False, 10)
        pt = Rhino.Geometry.Point2d(pt[0], pt[1] + 10)
        e.Display.Draw2dText("The generated checking srf will also show area text dot but will not be exported to Excel.", color, pt, False, 10)
        pt = Rhino.Geometry.Point2d(pt[0], pt[1] + 10)
        e.Display.Draw2dText("For clarity reason, the generated checking srf layer will be purged.", color, pt, False, 10)
        pt = Rhino.Geometry.Point2d(pt[0], pt[1] + 15)
        e.Display.Draw2dText("Area unit auto-mapped from your Rhino unit: mm--> m" + u"\u00B2" + ", m--> m" + u"\u00B2" + ", inch--> ft" + u"\u00B2" + ", ft--> ft" + u"\u00B2" + "", color_hightlight, pt, False, 10)
        pt = Rhino.Geometry.Point2d(pt[0], pt[1] + 10)
        e.Display.Draw2dText("Areas from same layer will try to merge dynamically if on same elevation.", color_hightlight, pt, False, 10)
        pt = Rhino.Geometry.Point2d(pt[0], pt[1] + 10)
        e.Display.Draw2dText("Accepting single surface(Z+ or Z- normal) and polysurface(open or enclosed, only check the face with Z- normal). ", color, pt, False, 10)
        pt = Rhino.Geometry.Point2d(pt[0], pt[1] + 10)
        e.Display.Draw2dText("Target area can be set for each keyword in the right toggle button.", color_hightlight, pt, False, 10)


        pt = Rhino.Geometry.Point2d(pt[0], pt[1] + 10)
        size = 20
        offset = 20

        grand_total = 0
        #sub_title = "X" * 10
        #sub_total = 0
        

        for data in self.data:

            layer, values = data
            if not rs.IsLayer(layer):
                print ("!!!!!!!!!!!!!!!!!! This layer is no longer existing ....{}".format(layer))
                self.data.remove(data)
                continue
            


            layer_tatal_area = values[0]
            note = values[3]
            #print note
            if layer_tatal_area == 0:
                continue
            
            factor =  self.layer_factor(layer)
            if factor != 1:
                layer_tatal_area *= factor

            grand_total += layer_tatal_area
            #sub_total += area
            text = "{}: {}".format(RHINO_LAYER.rhino_layer_to_user_layer(layer), convert_area_to_good_unit(layer_tatal_area))
            if note:
                text += note
            target = self.get_layer_target(layer)
            diff = 0
            if target:
                area_num, area_unit = convert_area_to_good_unit(layer_tatal_area, use_commas = False).split(" ", maxsplit = 1)
                diff = float(target) - float(area_num)
                text += " (Target: {:,.2f}{})".format(float(target), area_unit)
                if diff != 0:
                    text += " [{:,.2f}{} {}]".format(abs(diff), area_unit, "over" if diff < 0 else "under")
            pt = Rhino.Geometry.Point2d(pt[0], pt[1] + offset)
            color = rs.LayerColor(layer)
            e.Display.Draw2dText(text, color, pt, False, size)
            
            if diff:  # Only process if there's a difference
                # Pre-calculate common values
                temp_pt = Rhino.Geometry.Point2d(pt[0]-20, pt[1])
                
                if diff > 0:
                    e.Display.Draw2dText(u"\u25BC", 
                                       System.Drawing.Color.FromArgb(255, 145, 145),  # Soft red
                                       temp_pt, 
                                       False, 
                                       size)
                else:
                    e.Display.Draw2dText(u"\u25B2", 
                                       System.Drawing.Color.FromArgb(145, 255, 145),  # Soft green
                                       temp_pt, 
                                       False, 
                                       size)
 

            # draw curve from crv geo

        pt = Rhino.Geometry.Point2d(pt[0], pt[1] + 25)
        color = rs.CreateColor([87, 85, 83])
        #print "C"
        pt0 = System.Drawing.Point(pt[0], pt[1] )
        pt1 = System.Drawing.Point(pt[0] + 500, pt[1] )
        #print "A"
        e.Display.Draw2dLine(pt0, pt1, color, 5)
        #print "B"
        pt = Rhino.Geometry.Point2d(pt[0], pt[1] + 10)
        text = "Grand Total Area: {}".format(convert_area_to_good_unit(grand_total))
        e.Display.Draw2dText(text, color, pt, False, 30)
        #print "drawingsssss finish"




        #print "Done"
        #True = text center around thge define pt, Flase = lower-left

        root_layer = "GFA Internal Check Srfs"

        if "EA_GFA_IS_BAKING_EXCEL" not in sc.sticky:
            sc.sticky["EA_GFA_IS_BAKING_EXCEL"] = False
        
        if sc.sticky["EA_GFA_IS_BAKING_EXCEL"]:
            import toggle_GFA_right
            data_collection = []
            row = 0
            data_collection.append(EXCEL.ExcelDataItem("Layer", row, "A", is_bold = True, text_alignment = EXCEL.TextAlignment.Center, cell_color = (200, 200, 200), 
                                                        top_border_style = EXCEL.BorderStyle.Thick,
                                                        bottom_border_style = EXCEL.BorderStyle.Thick,
                                                        side_border_style = EXCEL.BorderStyle.Thick))
            data_collection.append(EXCEL.ExcelDataItem("Area", row, "B", is_bold = True, text_alignment = EXCEL.TextAlignment.Center, cell_color = (200, 200, 200),
                                                        top_border_style = EXCEL.BorderStyle.Thick,
                                                        bottom_border_style = EXCEL.BorderStyle.Thick,
                                                        side_border_style = EXCEL.BorderStyle.Thick,))
            data_collection.append(EXCEL.ExcelDataItem("Unit", row, "C", is_bold = True, text_alignment = EXCEL.TextAlignment.Center, cell_color = (200, 200, 200),
                                                        top_border_style = EXCEL.BorderStyle.Thick,
                                                        bottom_border_style = EXCEL.BorderStyle.Thick,
                                                        side_border_style = EXCEL.BorderStyle.Thick))
            row += 1

            for data in self.data:
                layer, values = data
                if root_layer in layer:
                    continue
                area = values[0]
                if area == 0:
                    continue

                factor =  self.layer_factor(layer)
                if factor != 1:
                    area *= factor
                area = convert_area_to_good_unit(area, use_commas = False)
                #print area
                area_num, area_unit = area.split(" ", maxsplit = 1)
                #print area_num
                layer = RHINO_LAYER.rhino_layer_to_user_layer(layer)
                
                cell_layer = EXCEL.ExcelDataItem(layer, row, "A", is_bold = True, bottom_border_style=EXCEL.BorderStyle.Thin)
                cell_area = EXCEL.ExcelDataItem(float(area_num), row, "B", bottom_border_style=EXCEL.BorderStyle.Thin)
                cell_unit = EXCEL.ExcelDataItem(area_unit, row, "C", bottom_border_style=EXCEL.BorderStyle.Thin)
                data_collection.append(cell_layer)
                data_collection.append(cell_area)
                data_collection.append(cell_unit)
                row += 1

            toggle_GFA_right.bake_action(data_collection)
            sc.sticky["EA_GFA_IS_BAKING_EXCEL"] = False
            

        if "EA_GFA_IS_BAKING_CRV" not in sc.sticky:
            sc.sticky["EA_GFA_IS_BAKING_CRV"] = False
        if sc.sticky["EA_GFA_IS_BAKING_CRV"]:
            if rs.IsLayer(root_layer):
                purge_layer_and_sub_layers(root_layer)
            for data in self.data:
                layer, values = data
                if root_layer in layer:
                    continue
                if values[0] == 0:
                    continue


                area, edges, faces, note = values
                color = rs.LayerColor(layer)
            
                #attr = sc.doc.ObjectAttributes()
                #attr.LayerIndex

                
                
                data_layer = secure_layer("{}::{}".format(root_layer, layer))

                rs.LayerColor(data_layer, color)
                #rs.DeleteObjects(rs.ObjectsByLayer(data_layer))
                shapes = [sc.doc.Objects.AddBrep(face) for face in faces]
                rs.ObjectLayer(shapes, data_layer)



            sc.sticky["EA_GFA_IS_BAKING_CRV"] = False
            sc.doc.Views.Redraw()





def secure_layer(layer):
    if not rs.IsLayer(layer):
        rs.AddLayer(layer)
    return layer


def purge_layer_and_sub_layers(root_layer):
    if 0 != rs.LayerChildCount(root_layer):

        for layer in rs.LayerChildren(root_layer):
            purge_layer_and_sub_layers(layer)


    rs.PurgeLayer(root_layer)
    return
    


def merge_as_crv_bool_union(breps):
    try:
        crvs = []
        for x in breps:
            try:
                if x.Loops.Count > 0:
                    crv = x.Loops[0].To3dCurve()
                    if crv:
                        crvs.append(crv)
            except Exception as ex:
                print("Error getting curve from brep: {}".format(str(ex)))
                continue
                
        if not crvs:
            return None
            
        factors = [1, 1.5]
        for factor in factors:
            try:
                union_crvs = Rhino.Geometry.Curve.CreateBooleanUnion(crvs, sc.doc.ModelAbsoluteTolerance * factor)
                if union_crvs:
                    union_breps = Rhino.Geometry.Brep.CreatePlanarBreps (union_crvs)
                    return union_breps
            except Exception as ex:
                print("Error in curve boolean union with factor {}: {}".format(factor, str(ex)))
                continue
        # print ("Failed to boolean union as Curve. Returning original breps.")
        return None
    except Exception as ex:
        print("Error in merge_as_crv_bool_union: {}".format(str(ex)))
        return None

def merge_coplaner_srf(breps):
    try:
        for brep in breps:
            try:
                param = rs.SurfaceClosestPoint(brep, RHINO_OBJ_DATA.get_center(brep))
                normal = rs.SurfaceNormal(brep, param)
                if normal[2] < 0:
                    brep.Flip()
            except Exception as ex:
                print("Error processing brep in merge_coplaner_srf: {}".format(str(ex)))
                continue

        # print ("original breps: ", breps)
        factors = [1, 1.5]
        for factor in factors:
            try:
                union_breps = Rhino.Geometry.Brep.CreateBooleanUnion(breps, sc.doc.ModelAbsoluteTolerance * factor, False)
                if union_breps:
                    # print ("unioined breps: ", union_breps)
                    break
            except Exception as ex:
                print("Error in boolean union with factor {}: {}".format(factor, str(ex)))
                continue
        else:
            # print ("Failed to boolean union as Brep. Trying as Curve.")
            try:
                union_breps = merge_as_crv_bool_union(breps)
                if not union_breps:
                    note = " #Warning! Cannot merge coplanar faces. Check your geometry cleanness."
                    return breps, note
            except Exception as ex:
                print("Error in merge_as_crv_bool_union: {}".format(str(ex)))
                note = " #Warning! Cannot merge coplanar faces. Check your geometry cleanness."
                return breps, note
        #print breps
        for brep in union_breps:
            try:
                success = brep.MergeCoplanarFaces(sc.doc.ModelAbsoluteTolerance, sc.doc.ModelAngleToleranceRadians * 1.5)
                if not success:
                    success = brep.MergeCoplanarFaces(sc.doc.ModelAbsoluteTolerance * 1.5, sc.doc.ModelAngleToleranceRadians * 3)
            except Exception as ex:
                print("Error merging coplanar faces: {}".format(str(ex)))
                continue

        return union_breps, None
    except Exception as ex:
        print("Error in merge_coplaner_srf: {}".format(str(ex)))
        return breps, " #Warning! Error in merge_coplaner_srf."

# @lru_cache(maxsize=None)
def get_merged_data(faces):
    height_dict = dict()
    for face in faces:
        try:
            face = face.DuplicateFace (False)# this step turn brepface to brep
            centroid_data = rs.SurfaceAreaCentroid(face)
            if not centroid_data or not centroid_data[0]:
                print("Warning: Could not get centroid for face")
                continue
            z = centroid_data[0][2]

            """
            need to put all 'similar' height to same key
            """
            for h_key in height_dict.keys():
                if abs(h_key - z) < sc.doc.ModelAbsoluteTolerance * 1.5:
                    height_dict[h_key].append(face)
                    break
            else:
                h_key = z
                height_dict[h_key] = [face]
        except Exception as ex:
            print("Error processing face in get_merged_data: {}".format(str(ex)))
            continue

        """
        if height_dict.has_key(h_key):
            height_dict[h_key].append(face)
        else:
            height_dict[h_key] = [face]
        """
    main_note = None
    for h_key, faces in height_dict.items():
        try:
            #faces = [brepface_to_brep(x) for x in faces]
            #height_dict[h_key] = merge_coplaner_srf(rs.BooleanUnion(faces, delete_input = False))
            faces, note = merge_coplaner_srf(faces)
            if note:
                main_note = note
            height_dict[h_key] = faces
        except Exception as ex:
            print("Error merging faces at height {}: {}".format(h_key, str(ex)))
            continue

    out_faces = []
    for faces in height_dict.values():
        out_faces.extend(faces)
    sum = 0
    out_crvs = []
    for face in out_faces:
        try:
            abstract_face = Rhino.Geometry.AreaMassProperties.Compute(face)
            if abstract_face:
                sum += abstract_face.Area
            else:
                main_note = " #Warning!! Check your geometry cleanness."
            out_crvs.extend(face.Edges )
        except Exception as ex:
            print("Error computing area for face: {}".format(str(ex)))
            continue


    return sum, out_crvs, out_faces, main_note


def get_schedule_layers():
    def is_good_layer(x):
        if not "[GFA]" in x:
            return False
        if not rs.IsLayer(x):
            return False
        if not rs.IsLayerVisible(x):
            return False
        return True


    try:
        good_layers = filter(is_good_layer, RHINO_LAYER.get_layers_in_order())
    except Exception as e:
        good_layers = filter(is_good_layer, rs.LayerNames())

    return good_layers




#"{:,}".format(total_amount)
# "Total cost is: ${:,.2f}".format(total_amount)
# Total cost is: $10,000.00
def convert_area_to_good_unit(area, use_commas = True):
    # only provide sqm or sqft, depeding on which system it is using. mm and m -->sqm, in, ft -->sqft

    unit = rs.UnitSystemName(capitalize=False, singular=True, abbreviate=False, model_units=True)

    def get_factor(unit):
        if unit == "millimeter":
            return 1.0/(1000**2), "m" + u"\u00B2"
        if unit == "meter":
            return 1.0, "m" + u"\u00B2"
        if unit == "inch":
            return 1.0/(12 * 12), "ft" + u"\u00B2"
        if unit == "foot":
            return 1.0, "ft" + u"\u00B2"
        return 1.0, "{0} x {0}".format(unit)
    factor, unit_text = get_factor(unit)

    area *= factor
    if use_commas:
        return "{:,.2f} {}".format(area, unit_text)
    else:
        return "{:.2f} {}".format(area, unit_text)
    pass



def get_current_objs():
    temp = set()
    layers = get_schedule_layers()
    for layer in layers:
        temp.update(get_objs_from_layer(layer))
    return temp


def get_objs_from_layer(layer):
    def is_good_obj(x):
        if not (rs.IsPolysurface(x) or rs.IsSurface(x)):
            return False
        if rs.IsObjectHidden(x):
            return False
        return True
    objs = rs.ObjectsByLayer(layer)
    objs = filter(is_good_obj , objs)
    
    return objs


def get_area_and_crv_geo_from_layer(layer):

    
    objs = get_objs_from_layer(layer)

    #print objs
    sum_area = 0
    crvs = []
    out_faces = []
    for obj in objs:


        #print "is polysurf"
        #prepare for the 2nd version to be faster
        brep = rs.coercebrep(obj)
        if not brep:
            print ("This layer [{}] has a non-brep object".format(layer))
            print (rs.ObjectType(obj))
            key = "EA_GFA_display_conduit"
            if sc.sticky.has_key(key):
                conduit = sc.sticky[key]
                conduit.reset_conduit_data()
            continue
        faces = brep.Faces

        is_single_face = False
        if faces.Count == 1:
            is_single_face = True

        for face in faces:
            #print face
            try:
                if is_facing_down(face, allow_facing_up = is_single_face):
                    #print "get down face"

                    area_props = Rhino.Geometry.AreaMassProperties.Compute(face)
                    if area_props:
                        sum_area += area_props.Area

                    #get crv geo but not add to doc
                    crv_geo = get_crv_geo_from_surface(brep, face)

                    crvs.extend(crv_geo)

                    #face_geo = get_face_geo_from_face(brep, face)
                    #print 123
                    #print face
                    #print face_geo
                    #print 321
                    out_faces.append(face)
            except Exception as ex:
                print("Error processing face in layer {}: {}".format(layer, str(ex)))
                continue




    sum_area, edges, faces, note = get_merged_data(out_faces)

    return sum_area, edges, faces, note


def brepface_to_brep(face):
    face_id = face.SurfaceIndex
    #print face_id
    brep = face.Brep
    surf = brep.Faces.ExtractFace(face_id)
    #print surf
    return surf

def get_face_geo_from_face(brep, face):
    face_id = face.SurfaceIndex
    #print face_id
    surf = brep.Faces.ExtractFace(face_id)
    #print surf
    return surf


def get_crv_geo_from_surface(brep, face):
    try:
        edge_ids = face.AdjacentEdges()
        #print edge_ids
        #print brep.Edges
        edges = [brep.Edges[x].EdgeCurve for x in list(edge_ids)]
        #print edges
        tolerance = sc.doc.ModelAbsoluteTolerance * 2.1
        joined_crvs = Rhino.Geometry.Curve.JoinCurves(edges, tolerance)
        return joined_crvs if joined_crvs else []
    except Exception as ex:
        print("Error getting curve geometry from surface: {}".format(str(ex)))
        return []

def is_facing_down(face, allow_facing_up = False):
    try:
        param = rs.SurfaceClosestPoint(face, RHINO_OBJ_DATA.get_center(face))
        normal = rs.SurfaceNormal(face, param)
        #print normal
        Z_down = [0,0,-1]
        #print rs.VectorAngle(normal, Z_down)
        #print rs.VectorCrossProduct(normal, Z_down)
        t = 0.001

        # make quick check for upfacing single surface
        if allow_facing_up:
            Z_up = [0,0,1]
            if -t < rs.VectorAngle(normal, Z_up) < t:
                return True

        if -t < rs.VectorAngle(normal, Z_down) < t:
            return True
        #if rs.VectorCrossProduct(normal, Z_down) < 0:
            #return False
        return False
    except Exception as ex:
        print("Error checking face orientation: {}".format(str(ex)))
        return False






@LOG.log(__file__, __title__)
@ERROR_HANDLE.try_catch_error()
def toggle_GFA():
    conduit = None
    key = "EA_GFA_display_conduit"
    if sc.sticky.has_key(key):
        conduit = sc.sticky[key]
    else:
        # create a conduit and place it in sticky
        conduit = EA_GFA_Conduit()
        sc.sticky[key] = conduit

    # Toggle enabled state for conduit. Every time this script is
    # run, it will turn the conduit on and off
    conduit.Enabled = not conduit.Enabled
    if conduit.Enabled: 
        conduit.add_hook()
        try: # this is helpful so user resume conduit will trigger refresh, becasue user might be modifying the doc while the conduit is off.
            conduit.reset_conduit_data()
        except Exception as e:
            print (e)
        print ("conduit enabled")
    else: 
        conduit.remove_hook()
        print ("conduit disabled")
    sc.doc.Views.Redraw()

    print( "Tool Finished")



    
if __name__ == "__main__":
    toggle_GFA()





    """
    to-do:
    right click to set desired GFA to each layer name, save to rhino doc data. and live compare how much is off from target.
    """
