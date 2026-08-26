VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_SetShapeText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsToolBase

' ============================================================
'  Tool: Set shape text
'  Sets the text content (caption) of a Visio shape.
' ============================================================

Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "set_shape_text"
End Property

Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = "Sets the text content (caption) of a Visio shape. " & _
        "Use this tool when the user asks to change, set or assign text to a Visio shape."
End Property

Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{" & _
            """shape_id"":{""type"":""string"",""description"":""Identifier (ID) of the shape whose text is to be changed.""}," & _
            """text"":{""type"":""string"",""description"":""New text to be assigned to the shape.""}" & _
        "}," & _
        """required"":[""shape_id"",""text""]" & _
    "}"
End Property

Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    On Error GoTo ErrHandler

    Dim shp As Visio.Shape
    Dim shape_id As String
    Dim textValue As String

    shape_id = ExtractJsonString(argsJson, "shape_id")
    textValue = ExtractJsonString(argsJson, "text")

    If Trim$(shape_id) = "" Then
        clsToolBase_Execute = "{""status"":""error"",""message"":" & JsonString("Shape ID is not specified") & "}"
        Exit Function
    End If

    ' Get the shape by its ID
    Set shp = Application.ActivePage.Shapes.ItemFromID(Int(shape_id))

    ' Set the shape text via the Characters object
    shp.Characters.text = textValue

    clsToolBase_Execute = "{""status"":""ok"",""message"":" & JsonString("Shape " & shape_id & " text '" & textValue & "' set successfully") & "}"
    Exit Function

ErrHandler:
    clsToolBase_Execute = "{""status"":""error"",""message"":" & JsonString(Err.Description) & "}"
End Function

