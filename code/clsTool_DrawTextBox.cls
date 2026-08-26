VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_DrawTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsToolBase

' ============================================================
' Tool: Draw text box
' Creates a text shape (label) on the Visio page.
' The containing rectangle is defined by its bottom-left corner
' (x1, y1) and top-right corner (x2, y2) in document units.
' ============================================================

Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "draw_text_box"
End Property

Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = "Creates a text shape (label) on a Visio page. " & _
        "Call this tool when the user asks to create a label, text box, caption, or title. " & _
        "Parameters: x1, y1 are the coordinates of the bottom-left corner and " & _
        "x2, y2 are the coordinates of the top-right corner of the rectangle " & _
        "that contains the text (in document units); text is the label content; " & _
        "angle is an optional rotation angle, 0 by default."
End Property

Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{" & _
            """x1"":{""type"":""string"",""description"":""X coordinate of the bottom-left corner of the text rectangle in document units""}," & _
            """y1"":{""type"":""string"",""description"":""Y coordinate of the bottom-left corner of the text rectangle in document units""}," & _
            """x2"":{""type"":""string"",""description"":""X coordinate of the top-right corner of the text rectangle in document units""}," & _
            """y2"":{""type"":""string"",""description"":""Y coordinate of the top-right corner of the text rectangle in document units""}," & _
            """text"":{""type"":""string"",""description"":""Label content""}," & _
            """angle"":{""type"":""string"",""description"":""Text rotation angle in degrees; optional, 0 by default""}" & _
        "}," & _
        """required"":[""x1"",""y1"",""x2"",""y2"",""text""]" & _
    "}"
End Property

Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    On Error GoTo ErrHandler

    Dim x1 As String, y1 As String, x2 As String, y2 As String
    Dim text As String, angle As String

    x1 = ExtractJsonString(argsJson, "x1")
    y1 = ExtractJsonString(argsJson, "y1")
    x2 = ExtractJsonString(argsJson, "x2")
    y2 = ExtractJsonString(argsJson, "y2")
    text = ExtractJsonString(argsJson, "text")
    angle = ExtractJsonString(argsJson, "angle")

    If ActivePage Is Nothing Then
        clsToolBase_Execute = "{""status"":""error"",""message"":""No active page""}"
        Exit Function
    End If

    ' Convert the string corner coordinates into numbers (decimal-separator safe).
    Dim dX1 As Double, dY1 As Double, dX2 As Double, dY2 As Double
    dX1 = CDbl(Replace(x1, ".", ","))
    dY1 = CDbl(Replace(y1, ".", ","))
    dX2 = CDbl(Replace(x2, ".", ","))
    dY2 = CDbl(Replace(y2, ".", ","))

    ' Create the rectangle shape using the two opposite corners.
    Dim shp As Visio.Shape
    Set shp = ActivePage.DrawRectangle(dX1, dY1, dX2, dY2)

    ' Set the text content.
    shp.Characters.text = text

    ' Make the shape transparent with no fill and no outline.
    shp.Cells("FillForegnd").FormulaForceU = "0"
    shp.Cells("FillPattern").FormulaForceU = "0"
    shp.Cells("LinePattern").FormulaForceU = "0"

    ' If an angle is specified and non-zero, add the Text Transform section
    ' and set the independent text rotation angle.
    If angle <> "" And angle <> "0" Then
        EnsureTextTransformSection shp
        shp.CellsSRC(visSectionObject, visRowTextXForm, visXFormAngle).FormulaU = angle & " deg"
    End If

    clsToolBase_Execute = "{""status"":""ok"",""shape_id"":" & shp.ID & "}"
    Exit Function

ErrHandler:
    clsToolBase_Execute = "{""status"":""error"",""message"":""" & Err.Description & """}"
End Function

