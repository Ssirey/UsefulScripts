Sub CopyFromExcel()

   Dim i As Integer 'счётчик для обозначения позиции элемента в последовательности
   Dim address As String
   Dim objExcel As Object
   Set objExcel = CreateObject("Excel.Application")

   Dim exWk As Excel.Workbook
   Dim ws As Excel.Worksheet

   Set objExcel = CreateObject("Excel.Application")
   Set exWk = objExcel.Workbooks.Open("C:\Users\ssire\Desktop\geologyTab.xlsx")
   Set ws = exWk.Sheets(1)

   i = 47

   'цикл Do While будет выполняться до тех пор, пока значение
   'текущего числа Фибоначчи не превысит 1000

   Do While i > 4
   
      ActiveDocument.Content.InsertAfter Text:="Слой "
      ActiveDocument.Content.InsertAfter Text:=ws.Range("K" & i).Value        'layerNumber
      ActiveDocument.Content.InsertAfter Text:=" ("
      ActiveDocument.Content.InsertAfter Text:=ws.Range("P" & i).Value        'upperNum
      ActiveDocument.Content.InsertAfter Text:=" - "
      ActiveDocument.Content.InsertAfter Text:=ws.Range("Q" & i).Value        'lowerNum
      ActiveDocument.Content.InsertAfter Text:=" м). "
      ActiveDocument.Content.InsertAfter Text:=ws.Range("R" & i).Value        'descriptionText
      
      If ws.Range("S" & i).Value <> "" Then
         ActiveDocument.Content.InsertAfter Text:="; структура "
         ActiveDocument.Content.InsertAfter Text:=ws.Range("S" & i).Value        'structure
      End If
      
      If ws.Range("T" & i).Value <> "" Then
         ActiveDocument.Content.InsertAfter Text:="; текстура "
         ActiveDocument.Content.InsertAfter Text:=ws.Range("T" & i).Value        ' texture
      End If

      If ws.Range("V" & i).Value <> "" Then
         ActiveDocument.Content.InsertAfter Text:="; индекс биотурбации "
         ActiveDocument.Content.InsertAfter Text:=ws.Range("V" & i).Value        'bioIndex
         If ws.Range("W" & i) <> "" Then
            ActiveDocument.Content.InsertAfter Text:=", фоссилии: "
            ActiveDocument.Content.InsertAfter Text:=ws.Range("W" & i).Value     'fassils
         End If
      End If

      If ws.Range("X" & i).Value <> "" Then
         ActiveDocument.Content.InsertAfter Text:="; реакция с HCl: "
         ActiveDocument.Content.InsertAfter Text:=ws.Range("X" & i).Value        'reaction
      End If
      ActiveDocument.Content.InsertAfter Text:="." & vbCrLf

      i = i - 1
   Loop

   exWk.Close savechanges:=False
   objExcel.Quit

End Sub
