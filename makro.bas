Attribute VB_Name = "Module1"
Option Explicit

' ============================================================================
' DATA.xlsx dagi har bir fuqaro (PINFL bo'yicha guruhlangan) uchun
' FORM.xlsx shabloni asosida alohida xlsx fayl yaratadi.
'
' ISHLATISH:
'	Oldindan bo'sh xlsx fayli ochib, kodlar shu yangi ochilgan fayl oynasiga yoziladi
'	 va shu yerda ishga tushiriladi
'
'   Excel 2013: Developer > Visual Basic > Insert > Module > shu kodni joylashtiring
'               (yoki File > Import File orqali .bas faylni to'g'ridan-to'g'ri import qiling)
'   LibreOffice Calc: Tools > Macros > Edit Macros > shu kodni joylashtiring
'
' Kod ishga tushganda: avval DATA.xlsx, keyin FORM.xlsx faylini tanlaysiz,
' so'ng natijalar saqlanadigan papkani ko'rsatasiz.
'
' elmurod vokhidov 
' ============================================================================

Sub Shablonlarni_Avtomatik_Yaratish()

    Dim wbSource As Workbook
    Dim wbTemplate As Workbook
    Dim wbNew As Workbook
    Dim wsSource As Worksheet
    Dim wsNew As Worksheet

    Dim sourceFile As Variant
    Dim templateFile As Variant
    Dim outputFolder As String

    Dim lastRow As Long, r As Long, i As Long, j As Long, k As Long

    Dim arrPinfl() As String
    Dim arrKadastr() As String
    Dim arrFuqaro() As String
    Dim arrOrtiqcha() As Double
    Dim arrSumma() As Double
    Dim arrSoliq() As Variant
    Dim n As Long

    Dim uniquePinfl() As String
    Dim memberOf() As Long
    Dim groupCount As Long

    Dim fileName As String
    Dim fullPath As String

    On Error GoTo XatoYuz

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' --- 1. Ma'lumotlar faylini tanlash ---
    sourceFile = Application.GetOpenFilename( _
        "Excel fayllari (*.xlsx;*.xls),*.xlsx;*.xls", , "Ma'lumotlar faylini (DATA.xlsx) tanlang")
    If sourceFile = False Then
        MsgBox "Ma'lumotlar fayli tanlanmadi.", vbExclamation
        GoTo Chiqish
    End If

    ' --- 2. Shablon faylini tanlash ---
    templateFile = Application.GetOpenFilename( _
        "Excel fayllari (*.xlsx;*.xls),*.xlsx;*.xls", , "Shablon faylini (FORM.xlsx) tanlang")
    If templateFile = False Then
        MsgBox "Shablon fayli tanlanmadi.", vbExclamation
        GoTo Chiqish
    End If

    ' --- 3. Natijalar papkasini tanlash ---
    ' Eslatma: agar FileDialog papka tanlash oynasi ochilmasa (ba'zi LibreOffice
    ' versiyalarida bo'lishi mumkin), pastdagi qatorni izohdan chiqarib,
    ' papka yo'lini qo'lda yozib qo'yishingiz mumkin:
    ' outputFolder = "C:\Natijalar"

    With Application.FileDialog(4) ' 4 = msoFileDialogFolderPicker
        .Title = "Tayyor fayllar saqlanadigan papkani tanlang"
        If .Show <> -1 Then
            MsgBox "Papka tanlanmadi.", vbExclamation
            GoTo Chiqish
        End If
        outputFolder = .SelectedItems(1)
    End With

    ' --- 4. Manba faylini ochish va o'qish ---
    Set wbSource = Workbooks.Open(sourceFile)
    Set wsSource = wbSource.Worksheets(1)

    lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(-4162).Row ' -4162 = xlUp

    n = 0
    ReDim arrPinfl(1 To lastRow)
    ReDim arrKadastr(1 To lastRow)
    ReDim arrFuqaro(1 To lastRow)
    ReDim arrOrtiqcha(1 To lastRow)
    ReDim arrSumma(1 To lastRow)
    ReDim arrSoliq(1 To lastRow)

    Dim pinfl As String, fuqaro As String

    ' 1-qator sarlavha (Hujjat sanasi, PINFL, Kadastr, Fuqaro, Ortiqcha to'lov, Summa, Soliq kodi)
    ' Ma'lumot 2-qatordan boshlanadi
    For r = 2 To lastRow
        pinfl = Trim(CStr(wsSource.Cells(r, "B").Value))
        fuqaro = Trim(CStr(wsSource.Cells(r, "D").Value))

        If pinfl <> "" Or fuqaro <> "" Then
            n = n + 1
            arrPinfl(n) = pinfl
            arrKadastr(n) = Trim(CStr(wsSource.Cells(r, "C").Value))
            arrFuqaro(n) = fuqaro
            arrOrtiqcha(n) = SonniAjrat(wsSource.Cells(r, "E").Value)
            arrSumma(n) = SonniAjrat(wsSource.Cells(r, "F").Value)
            arrSoliq(n) = wsSource.Cells(r, "G").Value
        End If
    Next r

    wbSource.Close SaveChanges:=False

    If n = 0 Then
        MsgBox "Ma'lumotlar topilmadi. B ustunida (PINFL) qiymat bormi tekshiring.", vbExclamation
        GoTo Chiqish
    End If

    ' --- 5. PINFL bo'yicha guruhlash (birinchi uchragan tartibda) ---
    ReDim uniquePinfl(1 To n)
    ReDim memberOf(1 To n)
    groupCount = 0

    Dim foundGroup As Long
    For i = 1 To n
        foundGroup = 0
        For j = 1 To groupCount
            If uniquePinfl(j) = arrPinfl(i) Then
                foundGroup = j
                Exit For
            End If
        Next j
        If foundGroup = 0 Then
            groupCount = groupCount + 1
            uniquePinfl(groupCount) = arrPinfl(i)
            foundGroup = groupCount
        End If
        memberOf(i) = foundGroup
    Next i

    ' --- 6. Shablon faylini ochish ---
    Set wbTemplate = Workbooks.Open(templateFile)

    Const ROW_BASE As Long = 14      ' shablonda ma'lumot yoziladigan birinchi qator
    Const TEMPLATE_ROWS As Long = 2  ' shablonda tayyor holda 14 va 15-qatorlar bor

    Dim members() As Long
    Dim memberCount As Long
    Dim g As Long

    For g = 1 To groupCount

        memberCount = 0
        ReDim members(1 To n)
        For i = 1 To n
            If memberOf(i) = g Then
                memberCount = memberCount + 1
                members(memberCount) = i
            End If
        Next i

        ' Shablon varag'idan yangi ish kitobi yaratish
        wbTemplate.Worksheets(1).Copy
        Set wbNew = ActiveWorkbook
        Set wsNew = wbNew.Worksheets(1)

        If memberCount < TEMPLATE_ROWS Then
            ' faqat 1 ta yozuv bo'lsa - ortiqcha shablon qatorini o'chirish
            For k = 1 To (TEMPLATE_ROWS - memberCount)
                wsNew.Rows(ROW_BASE + memberCount).Delete
            Next k

        ElseIf memberCount > TEMPLATE_ROWS Then
            ' 2 tadan ko'p yozuv bo'lsa - qo'shimcha qatorlar kiritish
            Dim extra As Long, insertAt As Long
            extra = memberCount - TEMPLATE_ROWS
            insertAt = ROW_BASE + TEMPLATE_ROWS  ' 16-qator

            For k = 1 To extra
                wsNew.Rows(insertAt).Insert Shift:=-4121 ' -4121 = xlDown
                wsNew.Rows(ROW_BASE + 1).Copy
                wsNew.Rows(insertAt).PasteSpecial Paste:=-4122 ' -4122 = xlPasteFormats
                Application.CutCopyMode = False
                wsNew.Rows(insertAt).RowHeight = wsNew.Rows(ROW_BASE + 1).RowHeight
                wsNew.Cells(insertAt, "B").Value = wsNew.Cells(ROW_BASE + 1, "B").Value
                wsNew.Cells(insertAt, "C").Value = wsNew.Cells(ROW_BASE + 1, "C").Value
                wsNew.Cells(insertAt, "L").Value = wsNew.Cells(ROW_BASE + 1, "L").Value
                insertAt = insertAt + 1
            Next k
        End If

        ' Ma'lumotlarni shablonga yozish
        For k = 1 To memberCount
            i = members(k)
            wsNew.Cells(ROW_BASE + k - 1, "D").Value = arrPinfl(i)      ' PINFL
            wsNew.Cells(ROW_BASE + k - 1, "E").Value = arrKadastr(i)    ' Kadastr raqami
            wsNew.Cells(ROW_BASE + k - 1, "F").Value = arrFuqaro(i)     ' Fuqaro
            wsNew.Cells(ROW_BASE + k - 1, "G").Value = arrSoliq(i)      ' Soliq kodi
            wsNew.Cells(ROW_BASE + k - 1, "H").Value = arrOrtiqcha(i)   ' Ortiqcha to'lov
            wsNew.Cells(ROW_BASE + k - 1, "K").Value = arrSumma(i)      ' Summa
        Next k

        ' --- Fayl nomi: FUQARO-SOLIQKODI (birinchi yozuv asosida) ---
        i = members(1)
        fileName = arrFuqaro(i) & "-" & CStr(arrSoliq(i))
        fileName = CleanFileName(fileName)
        If fileName = "" Then fileName = "Soliq_tolovchi_" & g

        fullPath = outputFolder & Application.PathSeparator & fileName & ".xlsx"
        If Dir(fullPath) <> "" Then
            fullPath = outputFolder & Application.PathSeparator & fileName & "_" & arrPinfl(i) & ".xlsx"
        End If

        wbNew.SaveAs Filename:=fullPath, FileFormat:=51 ' 51 = xlOpenXMLWorkbook (.xlsx)
        wbNew.Close SaveChanges:=False

    Next g

    wbTemplate.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Tayyor!" & vbCrLf & vbCrLf & _
           groupCount & " ta fuqaro uchun alohida fayl yaratildi." & vbCrLf & _
           "Fayllar tanlangan papkaga saqlandi.", vbInformation, "Jarayon tugadi"

    Exit Sub

XatoYuz:
    MsgBox "Xatolik yuz berdi: " & Err.Description, vbCritical
    Resume Chiqish

Chiqish:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub


' Excel/LibreOffice fayllarida "45 409,64" kabi probel+vergul formatidagi
' matnni haqiqiy songa aylantiradi (lokal sozlamalardan qat'iy nazar)
Function SonniAjrat(ByVal v As Variant) As Double
    Dim s As String

    If VarType(v) = vbDouble Or VarType(v) = vbInteger Or VarType(v) = vbSingle Or VarType(v) = vbLong Then
        SonniAjrat = CDbl(v)
        Exit Function
    End If

    s = Trim(CStr(v))
    s = Replace(s, Chr(160), "")  ' uzilmas probel (non-breaking space)
    s = Replace(s, " ", "")
    s = Replace(s, ",", ".")

    If s = "" Then
        SonniAjrat = 0
    Else
        SonniAjrat = Val(s)  ' Val() har doim "." dan foydalanadi, lokaldan qat'iy nazar
    End If
End Function


Function CleanFileName(ByVal txt As String) As String
    Dim badChars As Variant
    Dim i As Integer

    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For i = LBound(badChars) To UBound(badChars)
        txt = Replace(txt, badChars(i), "_")
    Next i

    CleanFileName = Trim(txt)
End Function
