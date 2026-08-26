Attribute VB_Name = "m_TextTransform"
Option Explicit

' ============================================================
'  Text Transform section helper
'  Makes TxtWidth, TxtHeight, TxtAngle, TxtPinX, TxtPinY
'  available for any shape.
' ============================================================

Public Sub EnsureTextTransformSection(ByRef shp As Visio.Shape)
    ' Add the Text Transform section if it has not been added before.
    ' This section is hidden by default for newly created shapes.
    If shp Is Nothing Then Exit Sub
    
    On Error GoTo ErrHandler
    
    If Not shp.SectionExists(visSectionObject, visRowTextXForm) Then
        shp.AddRow visSectionObject, visRowTextXForm, visTagDefault
        
        ' Pin of the text block relative to the shape origin
        shp.CellsSRC(visSectionObject, visRowTextXForm, visXFormPinX).FormulaU = "Width*0.5"
        shp.CellsSRC(visSectionObject, visRowTextXForm, visXFormPinY).FormulaU = "Height*0.5"
        
        ' Text block size: by default follows shape size
        shp.CellsSRC(visSectionObject, visRowTextXForm, visXFormWidth).FormulaU = "Width*1"
        shp.CellsSRC(visSectionObject, visRowTextXForm, visXFormHeight).FormulaU = "Height*1"
        
        ' Local pin of the text block
        shp.CellsSRC(visSectionObject, visRowTextXForm, visXFormLocPinX).FormulaU = "TxtWidth*0.5"
        shp.CellsSRC(visSectionObject, visRowTextXForm, visXFormLocPinY).FormulaU = "TxtHeight*0.5"
        
        ' Independent text rotation
        shp.CellsSRC(visSectionObject, visRowTextXForm, visXFormAngle).FormulaU = "0 deg"
    End If
    
    Exit Sub
ErrHandler:
    Debug.Print "EnsureTextTransformSection error: " & Err.Description
End Sub
