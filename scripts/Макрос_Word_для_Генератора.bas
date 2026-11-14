Attribute VB_Name = "SberbankWordGenerator"
Option Explicit

' ==============================================================================
' МАКРОС АВТОМАТИЧЕСКОЙ ГЕНЕРАЦИИ ДОГОВОРОВ WORD (ПДКП И ОДКП) ИЗ EXCEL
' Нейтральный деловой стандарт (ПАО Сбербанк: СБР, Аккредитив, Ипотека,
' МСК, Аккредитив с трансферацией, Перезалог из другого банка)
' ==============================================================================

' 1. ГЕНЕРАЦИЯ ОБОИХ ДОКУМЕНТОВ СРАЗУ (ПДКП + ОДКП)
Public Sub Сформировать_Все_Договоры_Word()
    If Not Проверить_Баланс() Then Exit Sub
    Call Создать_Документ_Word("ПДКП", "ПДКП_Порядок_расчетов.docx")
    Call Создать_Документ_Word("ОДКП", "ОДКП_Порядок_расчетов.docx")
    MsgBox "Готово! Файлы ПДКП и ОДКП успешно сформированы, открыты в Word и сохранены на Рабочем столе.", vbInformation, "Генератор договоров"
End Sub

' 2. ГЕНЕРАЦИЯ ТОЛЬКО ПДКП (КРАТКИЙ ТЕКСТ)
Public Sub Сформировать_ПДКП_в_Word()
    If Not Проверить_Баланс() Then Exit Sub
    Call Создать_Документ_Word("ПДКП", "ПДКП_Порядок_расчетов.docx")
End Sub

' 3. ГЕНЕРАЦИЯ ТОЛЬКО ОДКП (ПОЛНЫЙ ТЕХНИЧЕСКИЙ ТЕКСТ)
Public Sub Сформировать_ОДКП_в_Word()
    If Not Проверить_Баланс() Then Exit Sub
    Call Создать_Документ_Word("ОДКП", "ОДКП_Порядок_расчетов.docx")
End Sub

' 4. ЗАПОЛНЕНИЕ ГОТОВОГО ШАБЛОНА WORD ПО ПЛЕЙСХОЛДЕРАМ {{...}}
Public Sub Заполнить_Шаблон_Word_По_Плейсхолдерам()
    If Not Проверить_Баланс() Then Exit Sub
    
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim desktopPath As String
    Dim templatePath As String
    Dim savePath As String
    Dim wsIn As Worksheet
    
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    templatePath = desktopPath & "\ШАБЛОН_ДКП_Сбербанк_с_плейсхолдерами.docx"
    
    If Dir(templatePath) = "" Then
        MsgBox "Шаблон не найден по пути:" & vbCrLf & templatePath & vbCrLf & vbCrLf & _
               "Используйте макрос 'Сформировать_ОДКП_в_Word' для создания документа с нуля.", vbExclamation, "Шаблон не найден"
        Exit Sub
    End If
    
    Set wsIn = ThisWorkbook.Sheets("Вводные")
    
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Open(templatePath)
    
    ' Замена основных плейсхолдеров из листа Вводные
    Call Заменить_Плейсхолдер(wdDoc, "{{PUNKT_OBJEKTA}}", CStr(wsIn.Range("B2").Value))
    Call Заменить_Плейсхолдер(wdDoc, "{{CENA_DOGOVORA_CIFRAMI}}", Format(wsIn.Range("B1").Value, "#,##0"))
    Call Заменить_Плейсхолдер(wdDoc, "{{DATA_ZADATKA}}", CStr(wsIn.Range("B3").Value))
    Call Заменить_Плейсхолдер(wdDoc, "{{NOMER_PDKP}}", CStr(wsIn.Range("B4").Value))
    Call Заменить_Плейсхолдер(wdDoc, "{{DATA_PDKP}}", CStr(wsIn.Range("B5").Value))
    Call Заменить_Плейсхолдер(wdDoc, "{{DATA_KD}}", CStr(wsIn.Range("B6").Value))
    Call Заменить_Плейсхолдер(wdDoc, "{{GOROD_KD}}", CStr(wsIn.Range("B7").Value))
    Call Заменить_Плейсхолдер(wdDoc, "{{NOMER_KREDITA_PRODAVCA}}", CStr(wsIn.Range("B8").Value))
    Call Заменить_Плейсхолдер(wdDoc, "{{DATA_KREDITA_PRODAVCA}}", CStr(wsIn.Range("B9").Value))
    Call Заменить_Плейсхолдер(wdDoc, "{{BANK_PRODAVCA}}", CStr(wsIn.Range("B10").Value))
    
    ' Суммы
    Call Заменить_Плейсхолдер(wdDoc, "{{SUMMA_ZADATKA}}", Format(wsIn.Range("B12").Value, "#,##0"))
    Call Заменить_Плейсхолдер(wdDoc, "{{SUMMA_NAL_DO_SDELKI}}", Format(wsIn.Range("B13").Value, "#,##0"))
    Call Заменить_Плейсхолдер(wdDoc, "{{SUMMA_SBR}}", Format(wsIn.Range("B14").Value, "#,##0"))
    Call Заменить_Плейсхолдер(wdDoc, "{{SUMMA_AKKREDITIVA}}", Format(wsIn.Range("B15").Value, "#,##0"))
    Call Заменить_Плейсхолдер(wdDoc, "{{SUMMA_IPOTEKI}}", Format(wsIn.Range("B16").Value, "#,##0"))
    Call Заменить_Плейсхолдер(wdDoc, "{{SUMMA_MSK}}", Format(wsIn.Range("B17").Value, "#,##0"))
    Call Заменить_Плейсхолдер(wdDoc, "{{SUMMA_TRANSFER}}", Format(wsIn.Range("B18").Value, "#,##0"))
    Call Заменить_Плейсхолдер(wdDoc, "{{SUMMA_PEREZALOG_OBSHAYA}}", Format(wsIn.Range("B19").Value, "#,##0"))
    Call Заменить_Плейсхолдер(wdDoc, "{{SUMMA_DOLGA_PRODAVCA}}", Format(wsIn.Range("B20").Value, "#,##0"))
    Call Заменить_Плейсхолдер(wdDoc, "{{SUMMA_OSTATOK_PRODAVCU}}", Format(wsIn.Range("B19").Value - wsIn.Range("B20").Value, "#,##0"))
    
    ' Готовые сводные блоки ПДКП и ОДКП целиком
    Call Заменить_Длинный_Плейсхолдер(wdDoc, "{{BLOK_RASCHETOV_PDKP}}", CStr(ThisWorkbook.Sheets("ПДКП").Range("A1").Value))
    Call Заменить_Длинный_Плейсхолдер(wdDoc, "{{BLOK_RASCHETOV_ODKP}}", CStr(ThisWorkbook.Sheets("ОДКП").Range("A1").Value))
    
    savePath = desktopPath & "\Готовый_ДКП_" & Format(Now, "yyyymmdd_hhnnss") & ".docx"
    wdDoc.SaveAs2 savePath
    
    MsgBox "Шаблон успешно заполнен и сохранен:" & vbCrLf & savePath, vbInformation, "Готово"
End Sub

' ==============================================================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
' ==============================================================================

Private Function Проверить_Баланс() As Boolean
    Dim checkText As String
    checkText = CStr(ThisWorkbook.Sheets("Проверка").Range("A1").Value)
    
    If InStr(1, checkText, "ОШИБКА", vbTextCompare) > 0 Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("ВНИМАНИЕ! На листе «Проверка» обнаружено расхождение сумм:" & vbCrLf & vbCrLf & _
                     checkText & vbCrLf & vbCrLf & _
                     "Вы действительно хотите продолжить выгрузку в Word?", _
                     vbYesNo + vbExclamation + vbDefaultButton2, "Контроль баланса сумм")
        If ans = vbNo Then
            Проверить_Баланс = False
            Exit Function
        End If
    End If
    Проверить_Баланс = True
End Function

Private Sub Создать_Документ_Word(ByVal sheetName As String, ByVal fileName As String)
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim rawText As String
    Dim paragraphs() As String
    Dim i As Long
    Dim p As Object
    Dim desktopPath As String
    Dim fullPath As String
    
    rawText = CStr(ThisWorkbook.Sheets(sheetName).Range("A1").Value)
    rawText = Replace(rawText, vbCrLf, vbLf)
    rawText = Replace(rawText, vbCr, vbLf)
    paragraphs = Split(rawText, vbLf)
    
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    
    wdApp.Visible = True
    Set wdDoc = wdApp.Documents.Add
    
    With wdDoc.PageSetup
        .TopMargin = 36
        .BottomMargin = 36
        .LeftMargin = 43.2
        .RightMargin = 43.2
    End With
    
    Set p = wdDoc.Paragraphs.First
    p.Range.Text = "ЦЕНА ОБЪЕКТА И ПОРЯДОК РАСЧЕТОВ (" & sheetName & ")"
    p.Range.Font.Name = "Times New Roman"
    p.Range.Font.Size = 11
    p.Range.Font.Bold = True
    p.Format.Alignment = 1 ' wdAlignParagraphCenter
    p.Format.SpaceAfter = 8
    
    For i = LBound(paragraphs) To UBound(paragraphs)
        If Trim(paragraphs(i)) <> "" Then
            wdDoc.Paragraphs.Add
            Set p = wdDoc.Paragraphs.Last
            p.Range.Text = Trim(paragraphs(i))
            p.Range.Font.Name = "Times New Roman"
            p.Range.Font.Size = 10
            p.Range.Font.Bold = False
            p.Format.Alignment = 3 ' wdAlignParagraphJustify
            p.Format.FirstLineIndent = 18
            p.Format.SpaceBefore = 0
            p.Format.SpaceAfter = 4
            p.Format.LineSpacingRule = 4
            p.Format.LineSpacing = 12 * 1.08
        End If
    Next i
    
    desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    fullPath = desktopPath & "\" & fileName
    wdDoc.SaveAs2 fullPath
End Sub

Private Sub Заменить_Плейсхолдер(ByVal wdDoc As Object, ByVal findText As String, ByVal replaceText As String)
    With wdDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = Left(replaceText, 250)
        .Forward = True
        .Wrap = 1
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .Execute Replace:=2
    End With
End Sub

Private Sub Заменить_Длинный_Плейсхолдер(ByVal wdDoc As Object, ByVal findText As String, ByVal longText As String)
    Dim rng As Object
    Set rng = wdDoc.Content
    With rng.Find
        .ClearFormatting
        .Text = findText
        .Forward = True
        .Wrap = 0
        Do While .Execute
            rng.Text = Replace(longText, vbLf, vbCr)
            rng.Collapse 0
        Loop
    End With
End Sub
