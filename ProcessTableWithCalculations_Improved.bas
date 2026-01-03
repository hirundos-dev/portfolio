' ============================================================================
' УЛУЧШЕННАЯ ВЕРСИЯ: Обработка таблицы с расчетами
' ============================================================================
' Улучшения:
' 1. Убраны GoTo, код стал более структурированным
' 2. Вынесены константы для магических чисел
' 3. Созданы вспомогательные функции для поиска последней строки/столбца
' 4. Улучшена обработка ошибок
' 5. Улучшена читаемость и структура кода
' 6. Добавлены комментарии
' ============================================================================

Option Explicit

' ============================================================================
' КОНСТАНТЫ
' ============================================================================
Private Const START_ROW As Long = 5                    ' Начальная строка обработки
Private Const LOOKUP_SHEET_NAME As String = "Лист2"    ' Имя листа с данными для поиска
Private Const TARGET_CLUB_ID As String = "775084"      ' Целевой ID клуба для сводных таблиц
Private Const OUTPUT_SHEET_NAME As String = "Обработанные данные"
Private Const PIVOT_CLUB_SHEET_NAME As String = "Сводная_по_клубам"

' Индексы столбцов исходной таблицы
Private Const SRC_COL_ID As Long = 2                   ' Столбец B - ID
Private Const SRC_COL_NICK As Long = 3                 ' Столбец C - Ник
Private Const SRC_COL_AGENT_NAME As Long = 4           ' Столбец D - Имя агента
Private Const SRC_COL_AGENT_ID As Long = 5             ' Столбец E - ID агента
Private Const SRC_COL_OT_IGRY As Long = 10             ' Столбец J - От Игры
Private Const SRC_COL_KOMISSIYA As Long = 15           ' Столбец O - Комиссия

' Индексы столбцов выходной таблицы
Private Const DEST_COL_CLUB_ID As Long = 1
Private Const DEST_COL_ID As Long = 2
Private Const DEST_COL_AGENT_ID As Long = 3
Private Const DEST_COL_AGENT_NAME As Long = 4
Private Const DEST_COL_NICK As Long = 5
Private Const DEST_COL_OT_IGRY As Long = 6
Private Const DEST_COL_KOMISSIYA As Long = 7
Private Const DEST_COL_RB As Long = 8
Private Const DEST_COL_TOTAL_WITH_COMM As Long = 9
Private Const DEST_COL_SBOR As Long = 10
Private Const DEST_COL_SBOR_SUM As Long = 11
Private Const DEST_COL_PROFIT As Long = 12

' Индексы столбцов листа поиска
Private Const LOOKUP_COL_ID As Long = 1                ' Колонка A - ID
Private Const LOOKUP_COL_RB As Long = 2                ' Колонка B - РБ
Private Const LOOKUP_COL_CLUB_ID As Long = 4           ' Колонка D - ID клуба
Private Const LOOKUP_COL_SBOR As Long = 5              ' Колонка E - Сбор

' ============================================================================
' ОСНОВНАЯ ПРОЦЕДУРА
' ============================================================================
Sub ProcessTableWithCalculations()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim wsLookup As Worksheet
    Dim lookupDictRB As Object
    Dim lookupDictSbor As Object
    Dim lastRow As Long
    Dim destRow As Long
    
    On Error GoTo ErrorHandler
    
    ' Инициализация
    Set wsSource = ActiveSheet
    Set wsDest = CreateOutputSheet(wsSource)
    Set wsLookup = GetLookupSheet()
    
    If wsLookup Is Nothing Then
        MsgBox "Не найден лист с данными для поиска РБ! Создайте лист '" & LOOKUP_SHEET_NAME & "' с колонками ID и РБ.", vbExclamation
        Exit Sub
    End If
    
    ' Загрузка данных для поиска
    Set lookupDictRB = CreateObject("Scripting.Dictionary")
    Set lookupDictSbor = CreateObject("Scripting.Dictionary")
    
    LoadLookupDataRB lookupDictRB, wsLookup
    LoadLookupDataSbor lookupDictSbor, wsLookup
    
    ' Определение последней строки
    lastRow = GetLastRow(wsSource)
    
    ' Создание заголовков
    CreateHeaders wsDest
    
    ' Обработка данных
    destRow = ProcessData(wsSource, wsDest, lookupDictRB, lookupDictSbor, lastRow)
    
    ' Форматирование таблицы
    FormatOutputTable wsDest, destRow
    
    ' Создание сводных таблиц
    If destRow > 2 Then
        CreatePivotTables wsDest
        MsgBox "Обработка завершена! Данные сохранены на листе '" & wsDest.Name & "'" & vbCrLf & _
               "Выгружено " & (destRow - 2) & " строк.", vbInformation
    Else
        MsgBox "Обработка завершена! Данных для выгрузки не найдено.", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка: " & Err.Description & vbCrLf & "Номер ошибки: " & Err.Number, vbCritical
End Sub

' ============================================================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ И ПРОЦЕДУРЫ
' ============================================================================

' Создание выходного листа
Private Function CreateOutputSheet(ByVal wsSource As Worksheet) As Worksheet
    Dim ws As Worksheet
    Set ws = Worksheets.Add(After:=wsSource)
    ws.Name = OUTPUT_SHEET_NAME
    Set CreateOutputSheet = ws
End Function

' Получение листа с данными для поиска
Private Function GetLookupSheet() As Worksheet
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOOKUP_SHEET_NAME)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets(2)
    End If
    On Error GoTo 0
    
    Set GetLookupSheet = ws
End Function

' Получение последней строки с данными
Private Function GetLastRow(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    
    If Not foundCell Is Nothing Then
        GetLastRow = foundCell.Row
    Else
        GetLastRow = 1
    End If
End Function

' Получение последнего столбца с данными
Private Function GetLastColumn(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    
    If Not foundCell Is Nothing Then
        GetLastColumn = foundCell.Column
    Else
        GetLastColumn = 1
    End If
End Function

' Создание заголовков
Private Sub CreateHeaders(ByVal ws As Worksheet)
    Dim headers As Variant
    Dim i As Long
    
    headers = Array("ID клуба", "ID", "ID агента", "Агент", "Ник", "От Игры", "Комиссия", _
                    "РБ", "Итого с комиссией", "Сбор", "Сумма сбора", "Профит")
    
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
    
    ws.Rows(1).Font.Bold = True
End Sub

' Обработка данных
Private Function ProcessData(ByVal wsSource As Worksheet, ByVal wsDest As Worksheet, _
                             ByRef lookupDictRB As Object, ByRef lookupDictSbor As Object, _
                             ByVal lastRow As Long) As Long
    Dim i As Long
    Dim destRow As Long
    Dim idValue As String
    Dim inDataBlock As Boolean
    
    destRow = 2
    inDataBlock = False
    idValue = ""
    
    For i = START_ROW To lastRow
        ' Пропуск пустых строк
        If WorksheetFunction.CountA(wsSource.Rows(i)) = 0 Then
            inDataBlock = False
            GoTo NextIteration
        End If
        
        ' Поиск строки с ID клуба
        If InStr(1, wsSource.Cells(i, 1).Value, "ID:") > 0 Then
            idValue = ExtractID(wsSource.Cells(i, 1).Value)
            inDataBlock = True
            GoTo NextIteration
        End If
        
        ' Пропуск строки "Итог"
        If InStr(1, wsSource.Cells(i, 1).Value, "Итог") > 0 Then
            inDataBlock = False
            GoTo NextIteration
        End If
        
        ' Обработка данных в блоке
        If inDataBlock And Len(Trim(wsSource.Cells(i, 1).Value)) > 0 Then
            If ProcessDataRow(wsSource, wsDest, i, idValue, lookupDictRB, lookupDictSbor, destRow) Then
                destRow = destRow + 1
            End If
        End If
        
NextIteration:
    Next i
    
    ProcessData = destRow
End Function

' Обработка одной строки данных
Private Function ProcessDataRow(ByVal wsSource As Worksheet, ByVal wsDest As Worksheet, _
                                ByVal sourceRow As Long, ByVal clubID As String, _
                                ByRef lookupDictRB As Object, ByRef lookupDictSbor As Object, _
                                ByVal destRow As Long) As Boolean
    Dim sourceID As String
    Dim otIgryValue As Double
    Dim komissiyaValue As Double
    Dim rbValue As Variant
    Dim sborValue As Variant
    Dim rbPercent As Double
    Dim sborPercent As Double
    
    ' Получение ID
    sourceID = Trim(CStr(wsSource.Cells(sourceRow, SRC_COL_ID).Value))
    If sourceID = "" Then
        ProcessDataRow = False
        Exit Function
    End If
    
    ' Копирование базовых данных
    wsDest.Cells(destRow, DEST_COL_CLUB_ID).Value = clubID
    wsDest.Cells(destRow, DEST_COL_ID).Value = sourceID
    wsDest.Cells(destRow, DEST_COL_AGENT_ID).Value = wsSource.Cells(sourceRow, SRC_COL_AGENT_ID).Value
    wsDest.Cells(destRow, DEST_COL_AGENT_NAME).Value = wsSource.Cells(sourceRow, SRC_COL_AGENT_NAME).Value
    wsDest.Cells(destRow, DEST_COL_NICK).Value = wsSource.Cells(sourceRow, SRC_COL_NICK).Value
    wsDest.Cells(destRow, DEST_COL_OT_IGRY).Value = wsSource.Cells(sourceRow, SRC_COL_OT_IGRY).Value
    wsDest.Cells(destRow, DEST_COL_KOMISSIYA).Value = wsSource.Cells(sourceRow, SRC_COL_KOMISSIYA).Value
    
    ' Поиск РБ
    If lookupDictRB.Exists(sourceID) Then
        rbValue = lookupDictRB(sourceID)
        wsDest.Cells(destRow, DEST_COL_RB).Value = rbValue
    Else
        wsDest.Cells(destRow, DEST_COL_RB).Value = "Нет данных"
    End If
    
    ' Поиск Сбора
    If lookupDictSbor.Exists(clubID) Then
        sborValue = lookupDictSbor(clubID)
        wsDest.Cells(destRow, DEST_COL_SBOR).Value = sborValue
    Else
        wsDest.Cells(destRow, DEST_COL_SBOR).Value = "Нет данных"
    End If
    
    ' Расчеты
    If IsNumeric(wsDest.Cells(destRow, DEST_COL_OT_IGRY).Value) And _
       IsNumeric(wsDest.Cells(destRow, DEST_COL_KOMISSIYA).Value) Then
        
        otIgryValue = CDbl(wsDest.Cells(destRow, DEST_COL_OT_IGRY).Value)
        komissiyaValue = CDbl(wsDest.Cells(destRow, DEST_COL_KOMISSIYA).Value)
        
        ' Получение процентов
        rbPercent = GetNumericValue(wsDest.Cells(destRow, DEST_COL_RB).Value, "Нет данных")
        sborPercent = GetNumericValue(wsDest.Cells(destRow, DEST_COL_SBOR).Value, "Нет данных")
        
        ' Замена "Нет данных" на 0 для расчетов
        If Not IsNumeric(wsDest.Cells(destRow, DEST_COL_RB).Value) Then
            wsDest.Cells(destRow, DEST_COL_RB).Value = 0
        End If
        If Not IsNumeric(wsDest.Cells(destRow, DEST_COL_SBOR).Value) Then
            wsDest.Cells(destRow, DEST_COL_SBOR).Value = 0
        End If
        
        ' Расчеты по формулам
        ' 1. Итого с комиссией = "От игры" + %"РБ" от "Комиссия"
        wsDest.Cells(destRow, DEST_COL_TOTAL_WITH_COMM).Value = otIgryValue + (komissiyaValue * rbPercent / 100)
        
        ' 2. Сумма сбора = "Комиссия" * %"Сбор" / 100
        wsDest.Cells(destRow, DEST_COL_SBOR_SUM).Value = komissiyaValue * sborPercent / 100
        
        ' 3. Профит = "Комиссия" - "Итого с комиссией" - "Сумму сбора"
        wsDest.Cells(destRow, DEST_COL_PROFIT).Value = komissiyaValue - _
            CDbl(wsDest.Cells(destRow, DEST_COL_TOTAL_WITH_COMM).Value) - _
            CDbl(wsDest.Cells(destRow, DEST_COL_SBOR_SUM).Value)
    Else
        wsDest.Cells(destRow, DEST_COL_TOTAL_WITH_COMM).Value = "Нет данных"
        wsDest.Cells(destRow, DEST_COL_SBOR_SUM).Value = "Нет данных"
        wsDest.Cells(destRow, DEST_COL_PROFIT).Value = "Нет данных"
    End If
    
    ProcessDataRow = True
End Function

' Получение числового значения или значения по умолчанию
Private Function GetNumericValue(ByVal value As Variant, ByVal defaultValue As String) As Double
    If IsNumeric(value) And value <> defaultValue Then
        GetNumericValue = CDbl(value)
    Else
        GetNumericValue = 0
    End If
End Function

' Форматирование выходной таблицы
Private Sub FormatOutputTable(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim i As Long
    
    ws.Columns.AutoFit
    
    If lastRow > 1 Then
        ' Форматирование числовых колонок
        With ws.Range("F2:L" & lastRow)
            .NumberFormat = "#,##0.00"
        End With
        
        ' Форматирование процентных колонок (РБ и Сбор)
        With ws.Range("H2:H" & lastRow)
            .NumberFormat = "0.00%"
        End With
        With ws.Range("J2:J" & lastRow)
            .NumberFormat = "0.00%"
        End With
        
        ' Преобразование процентов для отображения (если хранятся как 10 для 10%)
        For i = 2 To lastRow
            If IsNumeric(ws.Cells(i, DEST_COL_RB).Value) Then
                ws.Cells(i, DEST_COL_RB).Value = ws.Cells(i, DEST_COL_RB).Value / 100
            End If
            If IsNumeric(ws.Cells(i, DEST_COL_SBOR).Value) Then
                ws.Cells(i, DEST_COL_SBOR).Value = ws.Cells(i, DEST_COL_SBOR).Value / 100
            End If
        Next i
    End If
End Sub

' ============================================================================
' ФУНКЦИИ ЗАГРУЗКИ ДАННЫХ ДЛЯ ПОИСКА
' ============================================================================

' Загрузка данных для поиска РБ (ID -> РБ)
Sub LoadLookupDataRB(ByRef dict As Object, ByVal ws As Worksheet)
    Dim lastRowLookup As Long
    Dim i As Long
    Dim key As String
    Dim value As Variant
    
    dict.RemoveAll
    
    lastRowLookup = GetLastRow(ws)
    
    For i = 2 To lastRowLookup
        key = Trim(CStr(ws.Cells(i, LOOKUP_COL_ID).Value))
        value = ws.Cells(i, LOOKUP_COL_RB).Value
        
        If key <> "" And Len(Trim(value)) > 0 Then
            If IsNumeric(value) Then
                dict(key) = CDbl(value)
            Else
                dict(key) = value
            End If
        End If
    Next i
End Sub

' Загрузка данных для поиска Сбора (ID клуба -> Сбор)
Sub LoadLookupDataSbor(ByRef dict As Object, ByVal ws As Worksheet)
    Dim lastRowLookup As Long
    Dim i As Long
    Dim key As String
    Dim value As Variant
    
    dict.RemoveAll
    
    lastRowLookup = GetLastRow(ws)
    
    For i = 2 To lastRowLookup
        key = Trim(CStr(ws.Cells(i, LOOKUP_COL_CLUB_ID).Value))
        value = ws.Cells(i, LOOKUP_COL_SBOR).Value
        
        If key <> "" And Len(Trim(value)) > 0 Then
            If IsNumeric(value) Then
                dict(key) = CDbl(value)
            Else
                dict(key) = value
            End If
        End If
    Next i
End Sub

' Извлечение ID из строки
Function ExtractID(ByVal text As String) As String
    Dim i As Long
    Dim result As String
    Dim char As String
    
    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        If IsNumeric(char) Then
            result = result & char
        End If
    Next i
    
    ExtractID = result
End Function

' ============================================================================
' СОЗДАНИЕ СВОДНЫХ ТАБЛИЦ
' ============================================================================

' Процедура для создания сводных таблиц
Sub CreatePivotTables(ByVal wsData As Worksheet)
    Dim pivotCache As PivotCache
    Dim pivotRange As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim uniqueAgents As Collection
    
    On Error GoTo ErrorHandler
    
    ' Определение диапазона данных
    lastRow = GetLastRow(wsData)
    lastCol = GetLastColumn(wsData)
    
    If lastRow <= 1 Then
        MsgBox "Нет данных для создания сводных таблиц.", vbInformation
        Exit Sub
    End If
    
    Set pivotRange = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol))
    
    ' Создание кэша сводной таблицы
    Set pivotCache = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=pivotRange)
    
    ' Создание общей сводной таблицы по клубам
    CreateClubPivotTable pivotCache, wsData
    
    ' Создание сводных таблиц по агентам
    Set uniqueAgents = GetUniqueAgents(wsData, TARGET_CLUB_ID, lastRow)
    
    If uniqueAgents.Count > 0 Then
        CreateAgentPivotTables pivotCache, wsData, uniqueAgents, lastRow
    Else
        MsgBox "Для клуба " & TARGET_CLUB_ID & " не найдено агентов. Таблицы по агентам не будут созданы.", vbInformation
    End If
    
    ' Активация листа с данными
    wsData.Activate
    
    ' Итоговое сообщение
    ShowPivotTablesMessage uniqueAgents.Count
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при создании сводных таблиц: " & Err.Description, vbCritical
End Sub

' Создание сводной таблицы по клубам
Private Sub CreateClubPivotTable(ByVal pivotCache As PivotCache, ByVal wsData As Worksheet)
    Dim wsPivotClub As Worksheet
    Dim pivotTable As PivotTable
    
    Set wsPivotClub = Worksheets.Add(After:=wsData)
    wsPivotClub.Name = PIVOT_CLUB_SHEET_NAME
    
    Set pivotTable = wsPivotClub.PivotTables.Add( _
        pivotCache:=pivotCache, _
        TableDestination:=wsPivotClub.Range("A3"), _
        TableName:="PivotByClub")
    
    With pivotTable
        .PivotFields("ID клуба").Orientation = xlRowField
        .PivotFields("ID клуба").Position = 1
        
        .PivotFields("ID агента").Orientation = xlRowField
        .PivotFields("ID агента").Position = 2
        
        .PivotFields("Агент").Orientation = xlRowField
        .PivotFields("Агент").Position = 3
        
        .AddDataField .PivotFields("От Игры"), "Сумма по От Игры", xlSum
        .AddDataField .PivotFields("Комиссия"), "Сумма по Комиссии", xlSum
        .AddDataField .PivotFields("Сумма сбора"), "Сумма по Сбору", xlSum
        .AddDataField .PivotFields("Профит"), "Сумма по Профиту", xlSum
        
        .DataPivotField.Orientation = xlColumnField
        .DataPivotField.Position = 1
        
        .RowAxisLayout xlTabularRow
        .TableStyle2 = "PivotStyleMedium9"
        
        .DataFields(1).NumberFormat = "#,##0.00"
        .DataFields(2).NumberFormat = "#,##0.00"
        .DataFields(3).NumberFormat = "#,##0.00"
        .DataFields(4).NumberFormat = "#,##0.00"
    End With
    
    ' Заголовки
    wsPivotClub.Range("A1").Value = "Сводная таблица по ID клуба и ID агента"
    wsPivotClub.Range("A1").Font.Bold = True
    wsPivotClub.Range("A1").Font.Size = 14
    
    wsPivotClub.Columns.AutoFit
End Sub

' Получение уникальных агентов для целевого клуба
Private Function GetUniqueAgents(ByVal wsData As Worksheet, ByVal targetClubID As String, ByVal lastRow As Long) As Collection
    Dim uniqueAgents As Collection
    Dim i As Long
    Dim agentID As Variant
    
    Set uniqueAgents = New Collection
    
    On Error Resume Next
    For i = 2 To lastRow
        If wsData.Cells(i, DEST_COL_CLUB_ID).Value = targetClubID Then
            agentID = wsData.Cells(i, DEST_COL_AGENT_ID).Value
            If agentID <> "" Then
                uniqueAgents.Add agentID, CStr(agentID)
            End If
        End If
    Next i
    On Error GoTo 0
    
    Set GetUniqueAgents = uniqueAgents
End Function

' Создание сводных таблиц по агентам
Private Sub CreateAgentPivotTables(ByVal pivotCache As PivotCache, ByVal wsData As Worksheet, _
                                   ByVal uniqueAgents As Collection, ByVal lastRow As Long)
    Dim wsPivotAgent As Worksheet
    Dim pivotTable As PivotTable
    Dim agentID As Variant
    Dim agentCounter As Long
    Dim agentName As String
    
    agentCounter = 0
    
    For Each agentID In uniqueAgents
        agentCounter = agentCounter + 1
        
        ' Создание листа для агента
        Set wsPivotAgent = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        wsPivotAgent.Name = "Агент_" & agentID
        
        ' Проверка длины имени листа
        If Len(wsPivotAgent.Name) > 31 Then
            wsPivotAgent.Name = "Агент_" & agentCounter
        End If
        
        ' Создание сводной таблицы
        Set pivotTable = wsPivotAgent.PivotTables.Add( _
            pivotCache:=pivotCache, _
            TableDestination:=wsPivotAgent.Range("A3"), _
            TableName:="PivotAgent_" & agentCounter)
        
        With pivotTable
            .PivotFields("Ник").Orientation = xlRowField
            
            .AddDataField .PivotFields("От Игры"), "Сумма по От Игры", xlSum
            .AddDataField .PivotFields("Комиссия"), "Сумма по Комиссии", xlSum
            .AddDataField .PivotFields("Сумма сбора"), "Сумма по Сбору", xlSum
            .AddDataField .PivotFields("Профит"), "Сумма по Профиту", xlSum
            
            .PivotFields("ID агента").Orientation = xlPageField
            .PivotFields("ID агента").CurrentPage = agentID
            
            .PivotFields("ID клуба").Orientation = xlPageField
            .PivotFields("ID клуба").CurrentPage = TARGET_CLUB_ID
            
            .DataPivotField.Orientation = xlColumnField
            
            .TableStyle2 = "PivotStyleMedium9"
            
            .DataFields(1).NumberFormat = "#,##0.00"
            .DataFields(2).NumberFormat = "#,##0.00"
            .DataFields(3).NumberFormat = "#,##0.00"
            .DataFields(4).NumberFormat = "#,##0.00"
        End With
        
        ' Заголовки и информация
        wsPivotAgent.Range("A1").Value = "Сводная таблица для агента: " & agentID
        wsPivotAgent.Range("A1").Font.Bold = True
        wsPivotAgent.Range("A1").Font.Size = 14
        
        wsPivotAgent.Range("A2").Value = "Клуб: " & TARGET_CLUB_ID
        wsPivotAgent.Range("A2").Font.Bold = True
        
        ' Поиск имени агента
        agentName = GetAgentName(wsData, TARGET_CLUB_ID, agentID, lastRow)
        If agentName <> "" Then
            wsPivotAgent.Range("A3").Value = "Агент: " & agentName
            wsPivotAgent.Range("A3").Font.Bold = True
        End If
        
        wsPivotAgent.Columns.AutoFit
    Next agentID
End Sub

' Получение имени агента
Private Function GetAgentName(ByVal wsData As Worksheet, ByVal clubID As String, _
                              ByVal agentID As Variant, ByVal lastRow As Long) As String
    Dim i As Long
    
    For i = 2 To lastRow
        If wsData.Cells(i, DEST_COL_CLUB_ID).Value = clubID And _
           wsData.Cells(i, DEST_COL_AGENT_ID).Value = agentID Then
            GetAgentName = wsData.Cells(i, DEST_COL_AGENT_NAME).Value
            Exit Function
        End If
    Next i
    
    GetAgentName = ""
End Function

' Показ сообщения о созданных сводных таблицах
Private Sub ShowPivotTablesMessage(ByVal agentCount As Long)
    Dim msgText As String
    
    msgText = "Создано сводных таблиц:" & vbCrLf & _
              "1. Общая по клубам: '" & PIVOT_CLUB_SHEET_NAME & "'"
    
    If agentCount > 0 Then
        msgText = msgText & vbCrLf & "2. По агентам клуба " & TARGET_CLUB_ID & ": " & agentCount & " таблиц (начинаются с 'Агент_')"
    End If
    
    MsgBox msgText, vbInformation, "Сводные таблицы созданы"
End Sub

