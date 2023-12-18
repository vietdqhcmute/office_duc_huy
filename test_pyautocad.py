# from pyautocad import Autocad, cache
from win32com.client import VARIANT
from win32com.client import Dispatch
import pythoncom, pyautocad
from pyautocad import aDouble, APoint
from shapely import Polygon, get_coordinates

#Functiton to draw intersection between two Polylines
    #Functon to change pyautocad_polygon to shaply_polygon
def cad_to_shapely_polygon (polyline):
    Coordinate_shapely = []
    Coordinate_cad = polyline.Coordinates
    for i in range(0, len(Coordinate_cad),2):
        Coordinate_shapely.append([Coordinate_cad[i], Coordinate_cad[i+1]])   
    return Coordinate_shapely

    #Function to find points of intersection using sharply
def draw_intersect (model_space, polyline_1, polyline_2):
    from shapely import Polygon, get_coordinates
    from win32com.client import Dispatch
    def close_aVariant(obj):
        from pythoncom import VT_ARRAY, VT_R8
        from win32com.client import VARIANT
        return VARIANT(VT_ARRAY | VT_R8 , obj)
    Dispatch("AutoCAD.Application").ActiveDocument.Layers.Add("Intersect")
    polygons = []
    intersection = Polygon(cad_to_shapely_polygon(polyline_1)).intersection(Polygon(cad_to_shapely_polygon(polyline_2)))
    points = get_coordinates(intersection)
    polygon = []
    for i in range(0, len(points)):
        point = (points[i][0], points[i][1])
        if point in polygon:
            polygon.append(point)
            polygons.append(tuple(polygon))
            polygon=[]
        else:
            polygon.append(point)
    #Format points into win32com.client format then draw polyline
    for polygon in polygons:
        formated_points = []
        for point in polygon:
            clone_point = list(point)
            clone_point.append(0.0)
            formated_points.extend(clone_point)
        polyline = model_space.AddPolyline(close_aVariant(formated_points))
        polyline.Layer = "Intersect"

# ---------------------------------------------------------------------------------------

#Set up cad
acad = Dispatch("AutoCAD.Application")
cad = acad.ActiveDocument
model_space = cad.ModelSpace

#Check polyline in each layer
layer_1 = "Ranh_38"
layer_2 = "0"
polylines_1 = [x for x in model_space if x.EntityName == "AcDbPolyline" and x.Layer == layer_1]
polylines_2 = [x for x in model_space if x.EntityName == "AcDbPolyline" and x.Layer == layer_2]

#Draw polyline of intersections
for polyline_1 in polylines_1:
    for polyline_2 in polylines_2:
        draw_intersect(model_space, polyline_1, polyline_2)