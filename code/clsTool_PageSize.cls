VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTool_PageSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements clsToolBase

Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "get_page_size"
End Property

Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = "Returns the width and height of the active Visio page in inches. " & _
        "Call this tool when the user asks about the size, width, or height of the current or active page, " & _
        "or when you need to determine the page size."
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
Dim width As Double, height As Double

    On Error GoTo ErrHandler
    
    Set vsoPage = ActivePage
    width = cellVal(vsoPage.PageSheet, "PageWidth")
    height = cellVal(vsoPage.PageSheet, "PageHeight")
    
    clsToolBase_Execute = "{""status"":""ok"",""message"":""Page width is " & Replace(width, ",", ".") & ", page height is " & Replace(height, ",", ".") & """}"
    Exit Function
    
ErrHandler:
    clsToolBase_Execute = "{""status"":""error"",""message"":""" & Err.Description & """}"
End Function

