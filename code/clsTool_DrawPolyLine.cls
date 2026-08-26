VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_DrawPolyLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsToolBase

Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "draw_poly_line"
End Property

Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = "Draws a polyline on the active Visio page. " & _
        "Call this tool when the user asks to draw a polyline, a curve connecting two points, or a polygon. " & _
        "If the coordinates of the first and last points match, a polygon will be drawn. " & _
        "Coordinates are specified in document units (inches), unless the user explicitly specifies otherwise."
End Property

Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{" & _
            """xy_arr"":{""type"":""string"",""description"":""Sequence of line point coordinates in the format (x1,y1,x2,y2,...,xn,yn) in inches""}," & _
            """line_weight"":{""type"":""string"",""description"":""Line thickness. Must be specified STRICTLY in points; default value is 0.72""}," & _
            """line_color"":{""type"":""string"",""description"":""Line color. Specified in RGB format. For example, white = 255,255,255, blue = 0,0,255""}," & _
            """fill_color"":{""type"":""string"",""description"":""Polygon fill color. Specified in RGB format. For example, white = 255,255,255, blue = 0,0,255""}" & _
        "}," & _
        """required"":[""xy_arr"",""line_weight"",""line_color""]" & _
    "}"
End Property

Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    On Error GoTo ErrHandler
    
    Dim xy_arr_str As String
    Dim line_weight As String
    Dim line_color As String
    Dim fill_color As String
    xy_arr_str = ExtractJsonNumber(argsJson, "xy_arr")
    line_weight = ExtractJsonNumber(argsJson, "line_weight")
    line_color = ExtractJsonNumber(argsJson, "line_color")
    fill_color = ExtractJsonNumber(argsJson, "fill_color")
    
    If ActivePage Is Nothing Then
        clsToolBase_Execute = "{""status"":""error"",""message"":""No active page""}"
        Exit Function
    End If

    Dim pnts_str() As String
    pnts_str = Split(xy_arr_str, ",")
    
    Dim pnts_dbl() As Double
    ReDim pnts_dbl(UBound(pnts_str))
    Dim i As Integer
    For i = 0 To UBound(pnts_str)
        pnts_dbl(i) = CDbl(Replace(pnts_str(i), ".", ","))
    Next i
    
    Dim shp As Visio.Shape
    Set shp = ActivePage.DrawPolyline(pnts_dbl, 0)
    SetCellVal shp, "LineWeight", line_weight & "pt"
    SetCellFrml shp, "LineColor", "RGB(" & line_color & ")"
    If fill_color <> "" Then
        SetCellFrml shp, "FillForegnd", "RGB(" & fill_color & ")"
    End If

    clsToolBase_Execute = "{""status"":""ok"",""shape_id"":" & shp.ID & "}"
    Exit Function
    
ErrHandler:
    clsToolBase_Execute = "{""status"":""error"",""message"":""" & Err.Description & """}"
End Function

' Helper function for extracting numbers from JSON.
Private Function ExtractJsonNumber(ByVal json As String, ByVal key As String) As String
    ' Use the corrected ExtractJsonString function from m_tools_test.
    ExtractJsonNumber = ExtractJsonString(json, key)
End Function


