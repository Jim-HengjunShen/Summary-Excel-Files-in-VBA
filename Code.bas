
 Sub CombineFiles()

    Dim mainFolder As String
    Dim folderName As String
    Dim fileName As String
    Dim folders As New Collection
    Dim folder As Variant

    Dim sourceWB As Workbook
    Dim targetWS As Worksheet
    Dim ws As Worksheet
    Dim nextRow As Long
    

    ' Main folder directory
    mainFolder = "/Users/hengjunshen/Documents/Excel Files/"
    Set targetWS = ThisWorkbook.Worksheets("Sheet1")
    targetWS.Cells.Clear
    
    ' Headers
    targetWS.Range("A1:E1").Value = Array("Company", "Benefit", "2025", "2026", "Change")
    With targetWS.Range("A1:E1").Font
        .Bold = True
        .Size = 16
    End With

    ' Find all folders for different company
    folderName = Dir(mainFolder & "*Company*", vbDirectory)

    Do While folderName <> ""
       'Check if it is a folder
       
        If folderName <> "." And folderName <> ".." Then
            If (GetAttr(mainFolder & folderName) And vbDirectory) <> 0 Then
                folders.Add folderName
            End If
        End If
   'Add the company folder name into the folders
        folderName = Dir()
    Loop
  



  ' Loop through each Company folder name in folders
    For Each folder In folders

    'Find Excel files
        fileName = Dir(mainFolder & folder & "/*.xls*")
        
    'Loop through all Excel files in a company folder
        Do While fileName <> ""
        
            'Only Find Excel files with 1 or 2 in the name
            If (InStr(fileName, "RETIREMENT") > 0 Or InStr(fileName, "GRATUITY") > 0) Then
            
                'Find next empty row in the current worksheet
                nextRow = targetWS.Cells(targetWS.Rows.Count, 1).End(xlUp).Row + 1
                targetWS.Cells(nextRow, 1).Value = folder
            
                'Open Excel file
                Set sourceWB = Workbooks.Open(mainFolder & folder & "/" & fileName)
           
                'Loop through all the sheets in Excel file to copy and paste different values.
                For Each ws In sourceWB.Worksheets
    
                   If InStr(ws.Name, "1") > 0 Then

                       ws.Range("A1").Copy Destination:=targetWS.Cells(nextRow, 2)

                   ElseIf InStr(ws.Name, "2") > 0 Then

                       ws.Range("A1").Copy Destination:=targetWS.Cells(nextRow, 3)
                       
                   ElseIf InStr(ws.Name, "Previous") > 0 Then

                       ws.Range("A1").Copy Destination:=targetWS.Cells(nextRow, 4)
                       
                   End If

               Next ws
            
                ' Close Excel file
                sourceWB.Close SaveChanges:=False
                
                  
                   ' Calculate YoY
                   'Check if the value is not empty and valid
                   If IsNumeric(targetWS.Cells(nextRow, 3).Value) And IsNumeric(targetWS.Cells(nextRow, 4).Value) Then

                     oldValue = targetWS.Cells(nextRow, 3).Value
                     newValue = targetWS.Cells(nextRow, 4).Value

                    If oldValue <> 0 Then

                        changeRate = (newValue - oldValue) / oldValue

                        targetWS.Cells(nextRow, 5).Value = changeRate
                        targetWS.Cells(nextRow, 5).NumberFormat = "0.00%"

                        ' Highlight changes greater than 30%
                        If Abs(changeRate) > 0.3 Then

                            targetWS.Cells(nextRow, 5).Font.Color = vbRed

                        End If

                    Else

                        targetWS.Cells(nextRow, 5).Value = "N/A"

                    End If

                End If

            End If
                
            ' Next file name in the Company folder
            fileName = Dir()

        Loop
        
    ' Next Company folder name in the folders
    Next folder
    
    With targetWS.Range("A1:E" & nextRow)
    .Borders.LineStyle = xlContinuous
    .Borders.Weight = xlThin
    End With

    ' Align columns
    targetWS.Columns("A").HorizontalAlignment = xlLeft
    targetWS.Columns("B").HorizontalAlignment = xlLeft
    targetWS.Columns("C:E").HorizontalAlignment = xlCenter
    
    MsgBox "Macro Finished"
    
End Sub

