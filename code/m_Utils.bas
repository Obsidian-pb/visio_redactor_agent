Attribute VB_Name = "m_Utils"
Option Explicit

' =============================================
' Вывод лога
' =============================================
Public Sub Log(msg)
    If DEBUG_MODE = True Then
        Debug.Print msg
    End If
End Sub



'' =============================================
'' Дописывает строки в файл history.md
'' =============================================
'Public Sub AppendToFile(ByVal s As String)
'    Dim filePath As String
'    Dim docPath As String
'
'    ' Определяем путь к папке текущего документа
'    docPath = ActiveDocument.Path
'    If Right(docPath, 1) <> "\" Then
'        docPath = docPath & "\"
'    End If
'
'    filePath = docPath & "history.md"
'
'    ' Открываем файл для дописывания (Append)
'    Open filePath For Append As #1
'
'    ' Записываем строку (без лишних кавычек)
'    Print #1, s
'
'    ' Закрываем файл
'    Close #1
'End Sub

' Универсальное чтение значения (если нет - возвращает defaultValue)
Public Function GetSettingFromRegistry(ByVal fullKey As String, ByVal defaultValue As String) As String
    Dim shell As Object
    Dim val As Variant

    On Error Resume Next
    Set shell = CreateObject("WScript.Shell")
    val = shell.RegRead(fullKey)
    If Err.Number <> 0 Then
        Err.Clear
        GetSettingFromRegistry = defaultValue
    Else
        GetSettingFromRegistry = CStr(val)
    End If
    On Error GoTo 0
End Function

' Универсальная запись значения в реестр (строка, REG_SZ)
Public Sub SaveSettingToRegistry(ByVal fullKey As String, ByVal value As String)
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.RegWrite fullKey, value, "REG_SZ"
End Sub


Public Sub sss()
Dim shp As Visio.Shape
    
    Set shp = Application.ActivePage.Shapes.ItemFromID(1)
    
    SetCellFrml shp, "Rounding", "1 mm"
End Sub
