VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHarnessToolsHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' Harness tools handler class
' ============================================================

Private tools As Collection   ' Collection of clsToolBase instances

' Initializes the tool registry and registers all available tools.
Public Sub Init()
    Set tools = New Collection
    
    ' Register tool instances.
    Dim toolInstance As clsToolBase

    ' Get the page size.
    Set toolInstance = New clsTool_PageSize
    tools.Add toolInstance, toolInstance.name
    ' Draw a polyline.
    Set toolInstance = New clsTool_DrawPolyLine
    tools.Add toolInstance, toolInstance.name
    ' Read a text file.
    Set toolInstance = New clsTool_ReadTextFile
    tools.Add toolInstance, toolInstance.name
    ' Write a text file.
    Set toolInstance = New clsTool_WriteTextFile
    tools.Add toolInstance, toolInstance.name
    ' Get the list of shapes.
    Set toolInstance = New clsTool_ShapesList
    tools.Add toolInstance, toolInstance.name
    ' Apply shape formatting.
    Set toolInstance = New clsTool_FormatShape
    tools.Add toolInstance, toolInstance.name
    ' Read shape cell values.
    Set toolInstance = New clsTool_ReadShapeCells
    tools.Add toolInstance, toolInstance.name
    ' Read shape text.
    Set toolInstance = New clsTool_ReadShapeText
    tools.Add toolInstance, toolInstance.name
    ' Set shape text.
    Set toolInstance = New clsTool_SetShapeText
    tools.Add toolInstance, toolInstance.name
    ' Create a text box shape.
    Set toolInstance = New clsTool_DrawTextBox
    ' Create a text box shape.
    Set toolInstance = New clsTool_DrawTextBox
    tools.Add toolInstance, toolInstance.name
    
End Sub


' Generate JSON for tools[]
Public Function GetToolsJson() As String
    If tools Is Nothing Or tools.Count = 0 Then
        GetToolsJson = "[]"
        Exit Function
    End If
    
    Dim i As Long
    Dim result As String
    result = "["
    
    For i = 1 To tools.Count
        Dim tool As clsToolBase
        Set tool = tools.item(i)
        
        ' name and Description are escaped via JsonString so that quotes,
        ' backslashes and newlines in tool metadata cannot break the JSON.
        result = result & "{" & _
            """type"":""function""," & _
            """function"":{" & _
                """name"":" & JsonString(tool.name) & "," & _
                """description"":" & JsonString(tool.Description) & "," & _
                """parameters"":" & tool.Schema & _
            "}" & _
        "}"
        
        If i < tools.Count Then result = result & ","
    Next i
    
    result = result & "]"
    GetToolsJson = result
End Function

' Execute tool by name
Public Function Execute(ByVal toolName As String, ByVal argsJson As String) As String
    If tools Is Nothing Then
        Execute = "{""error"":" & JsonString("Tool registry is not initialized") & "}"
        Exit Function
    End If
    
    On Error GoTo ErrHandler
    
    Dim tool As clsToolBase
    Set tool = tools.item(toolName)
    
    Execute = tool.Execute(argsJson)
    Exit Function
    
ErrHandler:
    Execute = "{""error"":" & JsonString("Tool not found: " & toolName) & "}"
End Function

