VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_ReadShapeText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsToolBase

' ============================================================
'  Tool: Read shape text
'  Reads the text content (caption) of a Visio shape.
' ============================================================

Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "read_shape_text"
End Property

Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = "Reads the text content (caption) of a Visio shape. " & _
        "Use this tool when the user asks to read, find out or show the text, " & _
        "caption or content of a Visio shape. Returns the shape text or an empty string if there is no text."
End Property

Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{" & _
            """shape_id"":{""type"":""string"",""description"":""Identifier (ID) of the shape whose text is to be read.""}" & _
        "}," & _
        """required"":[""shape_id""]" & _
    "}"
End Property

Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    On Error GoTo ErrHandler

    Dim shp As Visio.Shape
    Dim shape_id As String

    shape_id = ExtractJsonString(argsJson, "shape_id")

    If Trim$(shape_id) = "" Then
        clsToolBase_Execute = "{""status"":""error"",""message"":" & JsonString("Shape ID is not specified") & "}"
        Exit Function
    End If

    ' Get the shape by its ID
    Set shp = Application.ActivePage.Shapes.ItemFromID(Int(shape_id))

    ' Read the shape text via the Characters object
    clsToolBase_Execute = "{""status"":""ok"",""message"":" & JsonString(shp.Characters.text) & "}"
    Exit Function

ErrHandler:
    clsToolBase_Execute = "{""status"":""error"",""message"":" & JsonString(Err.Description) & "}"
End Function

