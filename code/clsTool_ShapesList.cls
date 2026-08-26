VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_ShapesList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements clsToolBase

Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "get_shapes_list"
End Property

Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = "Returns the list of shapes on the current page. " & _
        "Call this tool when you need to determine which shapes and IDs are available on the active page."
End Property

Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{}," & _
        """required"":[]" & _
    "}"
End Property

Private Function clsToolBase_Execute(ByVal argsJson As String) As String
Dim vsoPage As Visio.page
Dim shp As Visio.Shape
Dim result As String

    On Error GoTo ErrHandler
    
    Set vsoPage = ActivePage
    
    result = "Shape IDs on the active page: "
    For Each shp In vsoPage.Shapes
        result = result & shp.ID
        If shp.AreaIU > 0 Then
            result = result & " (polygon)" & ", "
        Else
            result = result & " (line)" & ", "
        End If
    Next shp
    result = Left(result, Len(result) - 2) & "."
    
    clsToolBase_Execute = "{""status"":""ok"",""message"":""" & result & """}"
    Exit Function
    
ErrHandler:
    clsToolBase_Execute = "{""status"":""error"",""message"":""" & Err.Description & """}"
End Function

