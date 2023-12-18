#Function to hatch polyline in AutoCAD
def hatch_aVariant(obj):
    from pythoncom import VT_ARRAY, VT_DISPATCH
    from win32com.client import VARIANT
    return VARIANT(VT_ARRAY | VT_DISPATCH, ([obj]))
def _hatch(win32com_model_space, acadloop):
    hatch = win32com_model_space.AddHatch(0, "ANSI31", True)
    hatch.PatternScale = 20
    try:
        hatch.AppendOuterLoop(hatch_aVariant(acadloop))
        hatch.Evaluate()
    except:
        pass
    # How to create win32_model_space
        # from win32com.client import Dispatch
        # acad = Dispatch("AutoCAD.Application")
        # doc = acad.ActiveDocument
        # win32_model_space = doc.ModelSpace

#Function to Close Polyline in AutoCAD
def close_aVariant(obj):
    from pythoncom import VT_ARRAY, VT_R8
    from win32com.client import VARIANT
    return VARIANT(VT_ARRAY | VT_R8 , (obj))
def _close(lst_open):
    from pyautocad import aDouble
    from win32com.client import Dispatch
    for polyline in lst_open:
        point = list(polyline.Coordinate(0))
        point = close_aVariant(aDouble(tuple(point)))
        polyline.AddVertex(len(polyline.Coordinates)/2, point)
    Dispatch("AutoCAD.Application").ActiveDocument.Regen(1)

#Function to check if a point is inside a Polygon
def is_inside_polygon(model_space, x, y, acadloop):
    from pyautocad import APoint
    ys = []
    for i in range(1, len(acadloop.Coordinates),2):
        ys.append(acadloop.Coordinates[i])
    ymax = (max(ys)*110)/100
    point1 = close_aVariant(APoint(x, y))
    point2 = close_aVariant(APoint(x, ymax))
    line = model_space.AddLine(point1, point2)
    intersect_points = acadloop.IntersectWith(line, 0)
    clone = (len(intersect_points)/3)%2
    line.Delete()
    if clone == 0:
        return False
    else:
        return True
    
#Function to draw Polylines from list of polygons
    #Format of polygons: [((x1, yi), (x2, y2), ...) , (('x1, 'y1), ('x2, 'y2))]
def _cad_polyline(model_space, polygons):
    from win32com.client import Dispatch
    Dispatch("AutoCAD.Application").ActiveDocument.Layers.Add("Python hatch")
    def close_aVariant(obj):
        from pythoncom import VT_ARRAY, VT_R8
        from win32com.client import VARIANT
        return VARIANT(VT_ARRAY | VT_R8 , obj)
    formated_points = []
    for polygon in polygons:
        formated_points = []
        for point in polygon:
            clone_point = list(point)
            clone_point.append(0.0)
            formated_points.extend(clone_point)
    polyline = model_space.AddPolyline(close_aVariant(formated_points))
    polyline.Layer = "Python hatch"


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
