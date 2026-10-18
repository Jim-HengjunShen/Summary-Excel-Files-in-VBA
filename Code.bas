
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
    With targetWS.Range("A1:E1")
        .Merge
        .Value = "Benefit Summary"
        .Font.Bold = True
        .Font.Size = 20
        .HorizontalAlignment = xlCenter
    End With
    
    targetWS.Range("A2:E2").Value = Array("Company", "Benefit", "2025", "2026", "Change")
    With targetWS.Range("A2:E2").Font
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
                
                a = 0
                b = 0
           
                'Loop through all the sheets in Excel file to copy and paste different values.
                For Each ws In sourceWB.Worksheets
    
                   If InStr(ws.Name, "1") > 0 Then

                        ws.Range("A1").Copy Destination:=targetWS.Cells(nextRow, 2)

                   ElseIf InStr(ws.Name, "2") > 0 Then

                        ws.Range("A1").Copy Destination:=targetWS.Cells(nextRow, 4)
                       
                        a = 1
                       
                   ElseIf InStr(ws.Name, "Previous") > 0 Then

                        ws.Range("A1").Copy Destination:=targetWS.Cells(nextRow, 3)
                        
                        b = 1
                       
                   End If

               Next ws
            
                ' Close Excel file
                sourceWB.Close SaveChanges:=False
                
                If a <> 1 Then

                    targetWS.Cells(nextRow, 4).Interior.Color = RGB(255, 0, 0)
                    targetWS.Cells(nextRow, 4).Value = "Missing data"

                ElseIf targetWS.Cells(nextRow, 4).Value = 0 Then

                    targetWS.Cells(nextRow, 4).Interior.Color = RGB(255, 255, 0)
                    targetWS.Cells(nextRow, 4).Value = "Zero Value"

                End If
                
                If b <> 1 Then

                    targetWS.Cells(nextRow, 3).Interior.Color = RGB(255, 0, 0)
                    targetWS.Cells(nextRow, 3).Value = "Missing data"

                ElseIf targetWS.Cells(nextRow, 3).Value = 0 Then

                    targetWS.Cells(nextRow, 3).Interior.Color = RGB(255, 255, 0)
                    targetWS.Cells(nextRow, 3).Value = "Zero Value"

                End If
                
                
                
                
                   ' Calculate YoY
                   'Check if the value is not empty and valid
                   If IsNumeric(targetWS.Cells(nextRow, 3).Value) And IsNumeric(targetWS.Cells(nextRow, 4).Value) Then

                     oldValue = targetWS.Cells(nextRow, 3).Value
                     newValue = targetWS.Cells(nextRow, 4).Value

                    If oldValue <> 0 And newValue <> 0 Then

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
    
    
    
    'Problem Summary

    ' Clear previous summary
    targetWS.Range("I:K").Clear

    ' Headers
    With targetWS.Range("I1:K1")
        .Merge
        .Value = "Issue Summary"
        .Font.Bold = True
        .Font.Size = 20
        .HorizontalAlignment = xlCenter
    End With
    
    targetWS.Range("I2:K2").Value = Array("Company", "Benefit", "Note")

    With targetWS.Range("I2:K2").Font
        .Bold = True
        .Size = 16
    End With

    lastRow = targetWS.Cells(targetWS.Rows.Count, 1).End(xlUp).Row
    nextRow = 3

    ' Check each row
    For i = 3 To lastRow

        company = targetWS.Cells(i, 1).Value
        benefit = targetWS.Cells(i, 2).Value

        ' Check 2025 and 2026
        For j = 3 To 4

            issue = ""

            If targetWS.Cells(i, j).Interior.Color = RGB(255, 0, 0) Then

                issue = "Missing " & targetWS.Cells(2, j).Value & " data"

            ElseIf targetWS.Cells(i, j).Interior.Color = RGB(255, 255, 0) Then

                issue = "Zero value in " & targetWS.Cells(2, j).Value

            End If

            If issue <> "" Then

                targetWS.Cells(nextRow, 9).Value = company
                targetWS.Cells(nextRow, 10).Value = benefit
                targetWS.Cells(nextRow, 11).Value = issue

                nextRow = nextRow + 1

            End If

        Next j

        ' Check percentage changes
        If IsNumeric(targetWS.Cells(i, 5).Value) Then

            issue = ""

            If targetWS.Cells(i, 5).Value > 0.3 Then

                issue = "Increase more than 30%"

            ElseIf targetWS.Cells(i, 5).Value < -0.3 Then

                issue = "Decrease more than 30%"

            End If

            If issue <> "" Then

                targetWS.Cells(nextRow, 9).Value = company
                targetWS.Cells(nextRow, 10).Value = benefit
                targetWS.Cells(nextRow, 11).Value = issue

                nextRow = nextRow + 1

            End If

        End If

    Next i

    For i = 3 To lastRow
        company = targetWS.Cells(i, 1).Value

        ' Only check each company once
        If Application.CountIf(targetWS.Range("A$2:A" & i), company) = 1 Then

            retirementCount = Application.CountIfs(targetWS.Range("A:A"), company, targetWS.Range("B:B"), "*Retirement*")
    
            gratuityCount = Application.CountIfs(targetWS.Range("A:A"), company, targetWS.Range("B:B"), "*Gratuity*")
            
            If retirementCount = 0 Then
                targetWS.Cells(nextRow, 9).Value = company
                targetWS.Cells(nextRow, 10).Value = "Retirement Benefit"
                targetWS.Cells(nextRow, 11).Value = "Missing Retirement Benefit Excel file"
                nextRow = nextRow + 1
            End If
            
            If gratuityCount = 0 Then
            
                targetWS.Cells(nextRow, 9).Value = company
                targetWS.Cells(nextRow, 10).Value = "Gratuity Benefit"
                targetWS.Cells(nextRow, 11).Value = "Missing Gratuity Benefit Excel file"
                nextRow = nextRow + 1
                
            End If
            
        End If
        
    Next i

    ' Format summary
    With targetWS.Range("I1:K" & nextRow - 1)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    MsgBox "Macro Finished"
    
End Sub


