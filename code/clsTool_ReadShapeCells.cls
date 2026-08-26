VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_ReadShapeCells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsToolBase

' ============================================================
' Tool name
' ============================================================
Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "read_shape_cells"
End Property

' ============================================================
' Description for the LLM
' ============================================================
Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = "Reads values from specified cells of a Visio shape. " & _
        "Call this tool when the user asks to determine properties of a Visio shape. " & _
        "Returns a list of 'Cell name:Cell value' pairs separated by semicolons. " & _
        "If a cell is not found, 'Cell name/#' is returned."
End Property

' ============================================================
' JSON schema for parameters
' ============================================================
Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{" & _
            """shape_id"":{""type"":""string"",""description"":""Identifier (ID) of the shape whose cell values should be read.""}," & _
            """cells_sequence"":{""type"":""string"",""description"":""Sequence of cell names separated by semicolons (;), for example 'Width;Height;PinX;PinY'""}" & _
        "}," & _
        """required"":[""shape_id"",""cells_sequence""]" & _
    "}"
End Property

' ============================================================
' Tool execution
' ============================================================
Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    On Error GoTo ErrHandler
    
    Dim shp As Visio.Shape
    Dim result As String
    
    Dim shape_id As String
    Dim cells_sequence As String
    shape_id = ExtractJsonString(argsJson, "shape_id")
    cells_sequence = ExtractJsonString(argsJson, "cells_sequence")
    
    ' Get the shape by its ID.
    Set shp = Application.ActivePage.Shapes.ItemFromID(Int(shape_id))
    
    ' Split the cell sequence.
    Dim cellNames() As String
    cellNames = Split(cells_sequence, ";")
    
    Dim i As Integer
    Dim cellName As String
    Dim cellValue As Variant
    
    ' Build the list of "Cell name/Cell value" pairs.
    result = ""
    For i = 0 To UBound(cellNames)
        cellName = Trim(cellNames(i))
        If cellName <> "" Then
            ' Check whether the cell exists.
            If ShapeHaveCell(shp, cellName) Then
                ' Read the cell value through the cellVal function from the t_Shapes module.
                cellValue = cellVal(shp, cellName, visUnitsString, "#")
                result = result & cellName & ":" & CStr(cellValue)
            Else
                result = result & cellName & ":cell does not exist"
            End If
            
            ' Add a separator between pairs.
            If i < UBound(cellNames) Then
                result = result & ";"
            End If
        End If
    Next i
    
    clsToolBase_Execute = "{""status"":""ok"",""message"":""" & JsonString(result) & """}"
    Exit Function
    
ErrHandler:
    clsToolBase_Execute = "{""status"":""error"",""message"":""" & JsonString(Err.Description) & """}"
End Function

