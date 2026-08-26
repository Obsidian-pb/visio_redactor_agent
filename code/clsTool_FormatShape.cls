VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_FormatShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsToolBase

' ============================================================
' Tool: Set shape format
' Assigns arbitrary values to Visio shape cells.
' ============================================================

Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "set_format_shape"
End Property

Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = "Assigns arbitrary values to Visio shape cells in batch mode. " & _
        "Call this tool when the user asks to change arbitrary properties of a Visio shape. " & _
        "The format is specified as a sequence of cell identifiers and the values to assign to them. " & _
        "If you do not know the correct name of the cell that should be changed, check the CELLS.md file located in this document folder."
End Property

Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{" & _
            """shape_id"":{""type"":""string"",""description"":""Identifier of the shape whose format should be changed.""}," & _
            """format_arr"":{""type"":""string"",""description"":""Sequence of cell names and their values in the format 'CELL1:value1;CELL2:value2;...'""}" & _
        "}," & _
        """required"":[""shape_id"",""format_arr""]" & _
    "}"
End Property

Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    On Error GoTo ErrHandler
    
    Dim shp As Visio.Shape
    Dim result As String
    
    Dim shape_id As String
    Dim format_arr As String
    shape_id = ExtractJsonNumber(argsJson, "shape_id")
    format_arr = ExtractJsonNumber(argsJson, "format_arr")
    
    result = "shapeId = " & shape_id & ". "
    
    Set shp = Application.ActivePage.Shapes.ItemFromID(Int(shape_id))

    Dim cell_frmts() As String
    cell_frmts = Split(format_arr, ";")
    
    Dim i As Integer
    Dim cll As String
    Dim frml As String
    For i = 0 To UBound(cell_frmts)
        cll = Split(cell_frmts(i), ":")(0)
        frml = Split(cell_frmts(i), ":")(1)
        
        ' Text transform cells are hidden by default for new shapes.
        ' Ensure the section exists before assigning values to these cells.
        If IsTextTransformCell(cll) Then
            EnsureTextTransformSection shp
        End If
        
        If ShapeHaveCell(shp, cll) Then
            SetCellFrml shp, cll, frml
            result = result & "Cell " & cll & " has been set to " & frml
        Else
            result = result & "Cell " & cll & " does not exist"
        End If
        
        If i = UBound(cell_frmts) Then
            result = result & ". "
        Else
            result = result & ", "
        End If
    Next i
    
    clsToolBase_Execute = "{""status"":""ok"",""message"":""" & result & """}"
    Exit Function
    
ErrHandler:
    clsToolBase_Execute = "{""status"":""error"",""message"":""" & Err.Description & """}"
End Function

' Returns True if the given cell name belongs to the Text Transform section.
Private Function IsTextTransformCell(ByVal cellName As String) As Boolean
    Select Case cellName
        Case "TxtWidth", "TxtHeight", "TxtAngle", "TxtPinX", "TxtPinY"
            IsTextTransformCell = True
        Case Else
            IsTextTransformCell = False
    End Select
End Function

' Helper function for extracting JSON values.
Private Function ExtractJsonNumber(ByVal json As String, ByVal key As String) As String
    ExtractJsonNumber = ExtractJsonString(json, key)
End Function

