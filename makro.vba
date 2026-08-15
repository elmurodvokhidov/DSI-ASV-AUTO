Option Explicit
  
  Sub Shablonlarni_Avtomatik_Yaratish()
  
  Dim wbSource As Workbook
  Dim wbTemplate As Workbook
  Dim wbNew As Workbook

  Dim wsSource As Worksheet
  Dim wsNew As Worksheet

  Dim sourceFile As Variant
  Dim templateFile As Variant
  Dim outputFolder As String

  Dim lastRow As Long
  Dim r As Long

  Dim pinfl As String
  Dim kadastr As String
  Dim nomi As String
  Dim sana As Variant
  Dim soliq As Variant
  Dim qaytariladigan As Variant
  Dim summa As Variant

  Dim fileName As String
  Dim fullPath As String

  Application.ScreenUpdating = False
  Application.DisplayAlerts = False

  ' 1. Ma'lumotlar Excel faylini tanlash
  sourceFile = Application.GetOpenFilename( _
      "Excel fayllari (*.xlsx;*.xls),*.xlsx;*.xls", _
      , "Ma'lumotlar faylini tanlang")

  If sourceFile = False Then
      MsgBox "Ma'lumotlar fayli tanlanmadi.", vbExclamation
      GoTo Chiqish
  End If

  ' 2. Shablon Excel faylini tanlash
  templateFile = Application.GetOpenFilename( _
      "Excel fayllari (*.xlsx;*.xls),*.xlsx;*.xls", _
      , "Shablon faylini tanlang")

  If templateFile = False Then
      MsgBox "Shablon fayli tanlanmadi.", vbExclamation
      GoTo Chiqish
  End If

  ' 3. Natijalar uchun papka
  With Application.FileDialog(msoFileDialogFolderPicker)
      .Title = "Tayyor fayllar saqlanadigan papkani tanlang"

      If .Show <> -1 Then
          MsgBox "Papka tanlanmadi.", vbExclamation
          GoTo Chiqish
      End If

      outputFolder = .SelectedItems(1)
  End With

  ' 4. Fayllarni ochish
  Set wbSource = Workbooks.Open(sourceFile)
  Set wbTemplate = Workbooks.Open(templateFile)

  ' Ma'lumotlar birinchi listda
  Set wsSource = wbSource.Worksheets(1)

  ' 5. Oxirgi qatorni aniqlash
  lastRow = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row

  ' 6. Har bir odam uchun alohida fayl
  For r = 7 To lastRow

      ' Manba ma'lumotlarini olish
      sana = wsSource.Cells(r, "B").Value
      pinfl = Trim(CStr(wsSource.Cells(r, "F").Value))
      kadastr = Trim(CStr(wsSource.Cells(r, "H").Value))
      nomi = Trim(CStr(wsSource.Cells(r, "I").Value))
      qaytariladigan = wsSource.Cells(r, "J").Value
      summa = wsSource.Cells(r, "K").Value
      soliq = wsSource.Cells(r, "L").Value

      ' Bo'sh qatorlarni o'tkazib yuborish
      If pinfl <> "" Or nomi <> "" Then

          ' Shablondan yangi workbook yaratish
          wbTemplate.Worksheets(1).Copy

          Set wbNew = ActiveWorkbook
          Set wsNew = wbNew.Worksheets(1)

          ' ==============================
          ' MA'LUMOTLARNI SHABLONGA YOZISH
          ' ==============================

          ' 4 - PINFL
          wsNew.Range("D13").Value = pinfl
  
          ' 4 va 5 orasidagi yangi joy - KADASTR KODI
          wsNew.Range("E13").Value = kadastr

          ' 5 - To'lovchi nomi
          wsNew.Range("F13").Value = nomi

          ' 6 - С налог
          wsNew.Range("G13").Value = soliq

          ' 7 - Переплата
          wsNew.Range("H13").Value = qaytariladigan

          ' 10 - Сумма
          wsNew.Range("J13").Value = summa

          ' 11 - Дата документа
          wsNew.Range("K13").Value = sana
          wsNew.Range("K13").NumberFormat = "dd.mm.yyyy"

          ' ==============================
          ' FAYL NOMI
          ' ==============================

          fileName = pinfl & "_" & nomi

          fileName = CleanFileName(fileName)

          If fileName = "" Then
              fileName = "Soliq_tolovchi_" & r
          End If

          fullPath = outputFolder & "\" & fileName & ".xlsx"

          ' Bir xil nomli fayl bo'lsa, nomiga qator raqamini qo'shish
          If Dir(fullPath) <> "" Then
              fullPath = outputFolder & "\" & fileName & "_" & r & ".xlsx"
          End If

          ' Faylni saqlash
          wbNew.SaveAs Filename:=fullPath, _
                       FileFormat:=xlOpenXMLWorkbook

          wbNew.Close SaveChanges:=False

      End If

  Next r

  ' Manba va shablonni yopish
  wbSource.Close SaveChanges:=False
  wbTemplate.Close SaveChanges:=False

  Application.ScreenUpdating = True
  Application.DisplayAlerts = True

  MsgBox "Tayyor!" & vbCrLf & vbCrLf & _
         "Har bir soliq to'lovchi uchun alohida Excel fayl yaratildi." & vbCrLf & _
         "Fayllar tanlangan papkaga saqlandi.", _
         vbInformation, "Jarayon tugadi"

  Exit Sub

  Chiqish:

  Application.ScreenUpdating = True
  Application.DisplayAlerts = True

  End Sub

  Function CleanFileName(ByVal txt As String) As String
  Dim badChars As Variant
  Dim i As Integer

  badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

  For i = LBound(badChars) To UBound(badChars)
      txt = Replace(txt, badChars(i), "_")
  Next i

  CleanFileName = Trim(txt)

End Function
