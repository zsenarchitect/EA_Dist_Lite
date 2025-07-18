__title__ = "RelocateAll_1643"
__doc__ = """Batch Relocation Tool for 1643 Project

Relocates all Rhino content (objects, cameras, named views) to new project coordinates.
Features:
- Moves all objects from world origin to new base point (440,160,0)
- Updates all named views and camera positions
- Maintains relative positions and orientations
- Preserves camera-target relationships
- Automatically updates view transformations
- Plays completion sound notification

Workflow:
1. Calculates vector from world origin to new base point
2. Moves all objects in the document
3. Updates all named views with new camera positions
4. Redraws viewport to reflect changes
5. Provides audio feedback on completion

Usage:
- Run tool to relocate entire project to new coordinates
- All content moves together maintaining spatial relationships
- Named views are automatically updated for new location"""

import rhinoscriptsyntax as rs
import Rhino # pyright: ignore

from EnneadTab import LOG, ERROR_HANDLE, SOUND


@LOG.log(__file__, __title__)
@ERROR_HANDLE.try_catch_error()
def relocate_left():
    # # Ask the user to pick a point in space
    # point = rs.GetPoint("Pick a new base point")
    # if point is None: return

    point = rs.CreatePoint(440,160,0)
    
    # Calculate the vector from the World Origin to the new point
    vector = rs.VectorCreate(point, [0,0,0])
    
    # Move all objects
    objects = rs.AllObjects(select=False)
    if objects:
        rs.MoveObjects(objects, vector)
    
    # Move all named views (cameras)
    start_pt, end_pt = point, Rhino.Geometry.Point3d(0,0,0)
    
    
    all_views = rs.NamedViews()
    source_orgin = end_pt
    source_x_pt = rs.PointAdd(end_pt, Rhino.Geometry.Point3d(1,0,0))
    source_array = [source_orgin, source_x_pt]
    
    target_orgin = start_pt
    target_x_pt = rs.PointAdd(start_pt, Rhino.Geometry.Point3d(1,0,0))
    target_array = [target_orgin, target_x_pt]
    
    
    for view in sorted(all_views, reverse=True):
        print (view)
        rs.RestoreNamedView(view, view = None, restore_bitmap = False)
        #current_cam = rs.ViewCamera(view = view, camera_location = None)
        current_cam, current_cam_target = rs.ViewCameraTarget(view = view, camera = None, target = None)

        temp_cam_pt = rs.AddPoint( current_cam )
        temp_cam_target_pt = rs.AddPoint( current_cam_target )

        new_cam = rs.OrientObject(temp_cam_pt, source_array, target_array)
        new_cam_target = rs.OrientObject(temp_cam_target_pt, source_array, target_array)
        rs.ViewCameraTarget(view = view, camera = new_cam, target = new_cam_target)

        rs.AddNamedView( view, view )#first view is the view name to override

        rs.DeleteObject(temp_cam_pt)
        rs.DeleteObject(temp_cam_target_pt)


    # Redraw the viewport to reflect changes
    rs.Redraw()

    SOUND.play_sound()




if __name__ == "__main__":
    relocate_left() 