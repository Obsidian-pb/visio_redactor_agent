Attribute VB_Name = "m_DiceRoll"
' ============================================================
' Модуль: m_DiceRoll
' Назначение: моделирование броска игровых костей
' Среда: MS Visio 2007 (VBA)
' ------------------------------------------------------------
' Доступные процедуры:
'   RollDie(DiceFaces)          - бросок одной кости
'   RollDice(Count, Faces)      - бросок нескольких костей (сумма)
'   RollDiceDetails(Count, Faces) - детальный отчёт по результатам
' ============================================================

Option Explicit

' ------------------------------------------------------------
' Инициализация генератора случайных чисел один раз
' при первой загрузке модуля.
' ------------------------------------------------------------
Private m_bRandomInit As Boolean

Private Sub EnsureRandomInit()
    If Not m_bRandomInit Then
        Randomize
        m_bRandomInit = True
    End If
End Sub

' ============================================================
' Функция: RollDie
' Назначение: моделирует бросок одной игральной кости.
' Вход:  DiceFaces - число граней (по умолчанию 6).
' Выход: целое число в диапазоне от 1 до DiceFaces.
' ============================================================
Public Function RollDie(Optional DiceFaces As Long = 6) As Long
    If DiceFaces < 1 Then DiceFaces = 1

    EnsureRandomInit

    ' Rnd возвращает значение из диапазона [0; 1).
    ' Формула даёт целое число из диапазона [1; DiceFaces].
    RollDie = Int(DiceFaces * Rnd) + 1
End Function

' ============================================================
' Функция RollDice
' Назначение: моделирует бросок нескольких костей и
'             возвращает сумму выпавших значений.
' Вход:  DiceCount - количество костей (по умолчанию 2);
'        DiceFaces - число граней каждой кости (по умолчанию 6);
'        Bonus     - модификатор, добавляемый к сумме
'                    (по умолчанию 0).
' Выход: суммарный результат броска.
' ============================================================
Public Function RollDice(Optional DiceCount As Long = 2, _
                         Optional DiceFaces As Long = 6, _
                         Optional Bonus As Long = 0) As Long
    Dim i As Long
    Dim lSum As Long

    If DiceCount < 1 Then DiceCount = 1
    If DiceFaces < 1 Then DiceFaces = 1

    EnsureRandomInit

    lSum = 0&
    For i = 1 To DiceCount
        lSum = lSum + RollDie(DiceFaces)
    Next i

    RollDice = lSum + Bonus
End Function

' ============================================================
' Функция RollDiceDetails
' Назначение: выполняет бросок нескольких костей и возвращает
'             строку с подробностями (значения каждой кости
'             и общий результат). Удобно для вывода в диалоге.
' Пример вывода: "Кости (3d6): [4, 2, 6] = 12 (с модиф. +2 = 14)"
' ============================================================
Public Function RollDiceDetails(Optional DiceCount As Integer = 2, _
                                Optional DiceFaces As Long = 6, _
                                Optional Bonus As Long = 0) As String
    Dim i As Integer
    Dim lRoll As Long, lSum As Long
    Dim sOut As String

    If DiceCount < 1 Then DiceCount = 1
    If DiceFaces < 1 Then DiceFaces = 1

    EnsureRandomInit

    lSum = 0&
    sOut = "Кости: " & CStr(DiceCount) & "d" & CStr(DiceFaces) & ": ["

    For i = 1 To DiceCount
        lRoll = RollDie(DiceFaces)          ' значение одной кости
        lSum = lSum + lRoll
        If i > 1 Then sOut = sOut & ", "
        sOut = sOut & CStr(lRoll)
    Next i

    sOut = sOut & "] = " & CStr(lSum)

    If Bonus <> 0& Then
        sOut = sOut & "  (модификатор " & IIf(Bonus > 0, "+", "") & _
               CStr(Bonus) & " => " & CStr(lSum + Bonus) & ")"
    End If

    RollDiceDetails = sOut
End Function

' ============================================================
' Процедура DemoRollDice
' Назначение: демонстрация работы функций модуля.
' Выводит результат в окно Immediate (Ctrl+G).
' ============================================================
Public Sub DemoRollDice()
    Debug.Print "Одиночная кость (1d6): " & RollDie(6)
    Debug.Print "Две кости (2d6): " & RollDice(2, 6)
    Debug.Print "Три кости (3d8): " & RollDice(3, 8)
    Debug.Print "Бросок 2d6 + модификатор +5: " & RollDice(2, 6, 5)
    Debug.Print RollDiceDetails(3, 6, 2)
    Debug.Print RollDiceDetails(1, 20)
End Sub

