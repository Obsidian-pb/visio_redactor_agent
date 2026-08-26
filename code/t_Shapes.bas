Attribute VB_Name = "t_Shapes"


Public Function cellVal(ByRef shps As Variant, ByVal cellName As String, Optional ByVal dataType As VisUnitCodes = visNumber, _
                        Optional defaultValue As Variant = 0, Optional escapeError As Boolean = True) As Variant
'Функция возвращает значение ячейки с указанным названием. Если такой ячейки нет, возвращает 0
Dim shp As Visio.Shape
Dim tmpVal As Variant
    
    If escapeError Then On Error GoTo ex
    
'    Debug.Print TypeName(shps)
    If typeName(shps) = "Shape" Then        'Если фигура
        Set shp = shps
        If shp.CellExists(cellName, 0) Then
            Select Case dataType
                Case Is = visNumber
                    cellVal = shp.Cells(cellName).result(dataType)
                Case Is = visUnitsString
                    cellVal = shp.Cells(cellName).ResultStr(dataType)
                Case Is = visDate
                    cellVal = shp.Cells(cellName).result(dataType)
'                    Debug.Assert cellName <> "Prop.FireTime"
'                    If cellName = "Prop.FireTime" Then Stop
                    If cellVal = 0 Then
                        cellVal = CDate(shp.Cells(cellName).ResultStr(visUnitsString))
                    End If
                Case Else
                    cellVal = shp.Cells(cellName).result(dataType)
            End Select
        Else
            cellVal = defaultValue
        End If
        Exit Function
    ElseIf typeName(shps) = "Shapes" Or typeName(shps) = "Collection" Then     'Если коллекция
        For Each shp In shps
            tmpVal = cellVal(shp, cellName, dataType, defaultValue)
            If tmpVal <> defaultValue Then
                cellVal = tmpVal
                Exit Function
            End If
        Next shp
    End If
    
cellVal = defaultValue
Exit Function
ex:
    cellVal = defaultValue
End Function





Public Sub SetCellVal(ByRef shp As Visio.Shape, ByVal cellName As String, ByVal NewVal As Variant)
'Set cell with cellName value. If such cell does not exists, does nothing
Dim cll As Visio.Cell
    
    On Error GoTo ex
    
    If shp.CellExists(cellName, 0) Then
        '!!!Need to test!!!
        shp.Cells(cellName).FormulaForce = """" & NewVal & """"
    End If
    
Exit Sub
ex:
    Debug.Print "Error in t_Shapes modul in 'Otcheti'! " & shp.name & ", " & cellName & ", " & NewVal
End Sub

Public Sub SetCellFrml(ByRef shp As Visio.Shape, ByVal cellName As String, ByVal NewFrml As Variant)
'Set cell with cellName formula. If such cell does not exists, does nothing
Dim cll As Visio.Cell
    
    On Error GoTo ex
    
    If shp.CellExists(cellName, 0) Then
        '!!!Need to test!!!
        shp.Cells(cellName).FormulaForceU = NewFrml
    End If
    
Exit Sub
ex:

End Sub

Public Function ShapeHaveCell(ByRef shp As Visio.Shape, ByVal cellName As String, _
                              Optional ByVal val As Variant = "", Optional ByVal delimiter As Variant = ";") As Boolean
'Функция возвращает Истина, если такая ячейка есть
'Если указано значение ячейки, также проверяется и она. Если указано несколько значений, то проверяются все они
Dim vals() As String
Dim curval As String
Dim i As Integer

On Error GoTo ex
    
    If shp.CellExists(cellName, 0) Then
        If val <> "" Then
            ' Проверяем одно ли значение передано в атрибуте val
            If InStr(1, val, delimiter) > 0 Then
                '---Если значений несколько
                vals = Split(val, delimiter)
                For i = 0 To UBound(vals)
                    curval = vals(i)
                    If shp.Cells(cellName).ResultStr(visUnitsString) = curval Then
                        ShapeHaveCell = True
                        Exit Function
                    ElseIf shp.Cells(cellName).result(visNumber) = curval Then
                        ShapeHaveCell = True
                        Exit Function
                    Else
                        ShapeHaveCell = False
                    End If
                Next i
            Else
                '---Если значение одно
                If shp.Cells(cellName).ResultStr(visUnitsString) = val Then
                    ShapeHaveCell = True
                ElseIf shp.Cells(cellName).result(visNumber) = val Then
                    ShapeHaveCell = True
                Else
                    ShapeHaveCell = False
                End If
            End If
        Else
            ShapeHaveCell = True
        End If
    Else
        ShapeHaveCell = False
    End If
    
Exit Function
ex:
    ShapeHaveCell = False
End Function


Public Function IsLineGeometry(ByVal shp As Visio.Shape, ByVal seci As Integer) As Boolean
Dim rwi As Integer
Dim rowTypeTag As VisRowTags

    For rwi = 1 To shp.rowCount(seci) - 1
        rowTypeTag = shp.RowType(seci, rwi)

        If rowTypeTag = VisRowTags.visTagMoveTo Or rowTypeTag = VisRowTags.visTagLineTo Then
            IsLineGeometry = True
        Else
            IsLineGeometry = False
            Exit Function
        End If
    Next rwi
End Function

Public Function IsEllipseGeometry(ByVal shp As Visio.Shape, ByVal seci As Integer) As Boolean
Dim rwi As Integer
Dim rowTypeTag As VisRowTags

    For rwi = 1 To shp.rowCount(seci) - 1
        rowTypeTag = shp.RowType(seci, rwi)

        If rowTypeTag = VisRowTags.visTagEllipse Then
            IsEllipseGeometry = True
            Exit Function
        Else
            IsEllipseGeometry = False
        End If
    Next rwi
End Function

Public Function IsWallsOnPage(ByRef page As Visio.page) As Boolean
Dim shp As Visio.Shape

    For Each shp In page.Shapes
        If isWall(shp) Then
            IsWallsOnPage = True
            Exit Function
        End If
    Next shp
IsWallsOnPage = False
End Function

Public Sub GetLineCoordinates(ByVal shp As Visio.Shape, ByRef startX As Double, ByRef startY As Double, _
                               ByRef endX As Double, ByRef endY As Double)
' Процедура получения координат начала и конца линии
Dim geomSection As Integer
Dim rowCount As Integer
    
    ' Получаем первую строку геометрии
    geomSection = visSectionFirstComponent
    
    On Error Resume Next
    rowCount = shp.rowCount(geomSection)  ', visRowFirst)
    
    If rowCount >= 2 Then
        ' Начало линии (первая точка)
        startX = shp.CellsSRC(geomSection, visRowFirst, visX).result(VALUE_MODE)
        startY = shp.CellsSRC(geomSection, visRowFirst, visY).result(VALUE_MODE)
        
        ' Конец линии (вторая точка)
        endX = shp.CellsSRC(geomSection, visRowFirst + 1, visX).result(VALUE_MODE)
        endY = shp.CellsSRC(geomSection, visRowFirst + 1, visY).result(VALUE_MODE)
    Else
        ' Если не удалось получить из геометрии, используем PinX/PinY и размер
        startX = shp.Cells("PinX").result(VALUE_MODE)
        startY = shp.Cells("PinY").result(VALUE_MODE)
        endX = startX + shp.Cells("Width").result(VALUE_MODE)
        endY = startY
    End If
    On Error GoTo 0
End Sub






