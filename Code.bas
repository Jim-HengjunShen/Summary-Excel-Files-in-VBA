Sub CombineFiles()
    
    Dim mainfolder As String
    Dim foldername As String
    Dim filename As String
    Dim excelfile As Workbook
    Dim currentworksheet As Worksheet
    Dim ws As Worksheet
    Dim nextrow As Integer
    Dim a As Integer
    Dim b As Integer
    Dim oldvalue As Integer
    Dim newvalue As Integer
    Dim changerate As Integer
    Dim lastrow As Integer
    Dim i As Integer
    Dim j As Integer
    Dim folders As New Collection
    Dim company As String
    Dim benefit As String
    Dim issue As String
    Dim retirementcount As Integer
    Dim gratuitycount As Integer
    
    
    ' Main folder directory
    mainfolder = "/Users/hengjunshen/Documents/Excel Files/"
    Set currentworksheet = ThisWorkbook.Worksheets("Sheet1")
    currentworksheet.Cells.Clear

    ' Headers
    With currentworksheet.Range("A1:E1")
        .Merge
        .Value = "Benefit Summary"
        .Font.Bold = True
        .Font.Size = 20
        .HorizontalAlignment = xlCenter
    End With

    currentworksheet.Range("A2:E2").Value = Array("Company", "Benefit", "2025", "2026", "Change")
    With currentworksheet.Range("A2:E2").Font
        .Bold = True
        .Size = 16
    End With

    ' Find all folders for different company
    foldername = Dir(mainfolder & "*Company*", vbDirectory)

    Do While foldername <> ""
        'Check if it is a folder
        If foldername <> "." And foldername <> ".." Then
            If (GetAttr(mainfolder & foldername) And vbDirectory) <> 0 Then
                folders.Add foldername
            End If
        End If

        'Add the company folder name into the folders
        foldername = Dir()
    Loop

    ' Loop through each Company folder name in folders
    For Each folder In folders

        'Find Excel files
        filename = Dir(mainfolder & folder & "/*.xls*")

        'Loop through all Excel files in a company folder
        Do While filename <> ""

            'Only Find Excel files with benefit(post-retirement benefit, gratuity benefit) name in the name
            If (InStr(filename, "RETIREMENT") > 0 Or InStr(filename, "GRATUITY") > 0) Then

                'Find next empty row in the current worksheet
                nextrow = currentworksheet.Cells(currentworksheet.Rows.Count, 1).End(xlUp).Row + 1
                currentworksheet.Cells(nextrow, 1).Value = folder

                'Open Excel file
                Set excelfile = Workbooks.Open(mainfolder & folder & "/" & filename)

                a = 0
                b = 0

                'Loop through all the sheets in Excel file to copy and paste different values.
                For Each ws In excelfile.Worksheets

                    If InStr(ws.Name, "1") > 0 Then

                        ws.Range("A1").Copy Destination:=currentworksheet.Cells(nextrow, 2)

                    ElseIf InStr(ws.Name, "2") > 0 Then

                        ws.Range("A1").Copy Destination:=currentworksheet.Cells(nextrow, 4)

                        a = 1

                    ElseIf InStr(ws.Name, "Previous") > 0 Then

                        ws.Range("A1").Copy Destination:=currentworksheet.Cells(nextrow, 3)

                        b = 1

                    End If

                Next ws

                ' Close Excel file
                excelfile.Close
                
                'Check if the data is 0 or missing
                If a <> 1 Then

                    currentworksheet.Cells(nextrow, 4).Interior.Color = RGB(255, 0, 0)
                    currentworksheet.Cells(nextrow, 4).Value = "Missing data"

                ElseIf currentworksheet.Cells(nextrow, 4).Value = 0 Then

                    currentworksheet.Cells(nextrow, 4).Interior.Color = RGB(255, 255, 0)
                    currentworksheet.Cells(nextrow, 4).Value = "Zero Value"

                End If

                If b <> 1 Then

                    currentworksheet.Cells(nextrow, 3).Interior.Color = RGB(255, 0, 0)
                    currentworksheet.Cells(nextrow, 3).Value = "Missing data"

                ElseIf currentworksheet.Cells(nextrow, 3).Value = 0 Then

                    currentworksheet.Cells(nextrow, 3).Interior.Color = RGB(255, 255, 0)
                    currentworksheet.Cells(nextrow, 3).Value = "Zero Value"

                End If

                'Calculate YoY
                'Check if the value is not empty and valid
                If IsNumeric(currentworksheet.Cells(nextrow, 3).Value) And IsNumeric(currentworksheet.Cells(nextrow, 4).Value) Then

                    oldvalue = currentworksheet.Cells(nextrow, 3).Value
                    newvalue = currentworksheet.Cells(nextrow, 4).Value

                    If oldvalue <> 0 And newvalue <> 0 Then

                        changerate = (newvalue - oldvalue) / oldvalue

                        currentworksheet.Cells(nextrow, 5).Value = changerate
                        currentworksheet.Cells(nextrow, 5).NumberFormat = "0.00%"

                        ' Highlight changes greater than 30%
                        If Abs(changerate) > 0.3 Then

                            currentworksheet.Cells(nextrow, 5).Font.Color = vbRed

                        End If

                    Else

                        currentworksheet.Cells(nextrow, 5).Value = "N/A"

                    End If

                End If

            End If

            ' Next file name in the Company folder
            filename = Dir()

        Loop

        ' Next Company folder name in the folders
    Next folder

    With currentworksheet.Range("A1:E" & nextrow)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    currentworksheet.Columns("A").HorizontalAlignment = xlLeft
    currentworksheet.Columns("B").HorizontalAlignment = xlLeft
    currentworksheet.Columns("C:E").HorizontalAlignment = xlCenter

    'Issue Summary

    ' Clear previous summary
    currentworksheet.Range("I:K").Clear

    ' Headers
    With currentworksheet.Range("I1:K1")
        .Merge
        .Value = "Issue Summary"
        .Font.Bold = True
        .Font.Size = 20
        .HorizontalAlignment = xlCenter
    End With

    currentworksheet.Range("I2:K2").Value = Array("Company", "Benefit", "Note")

    With currentworksheet.Range("I2:K2").Font
        .Bold = True
        .Size = 16
    End With

    lastrow = currentworksheet.Cells(currentworksheet.Rows.Count, 1).End(xlUp).Row
    nextrow = 3

    ' Check each row
    For i = 3 To lastrow

        company = currentworksheet.Cells(i, 1).Value
        benefit = currentworksheet.Cells(i, 2).Value

        ' Check 2025 and 2026 data, clarifying missing or zero value, wrote in issue summary
        For j = 3 To 4

            issue = ""

            If currentworksheet.Cells(i, j).Interior.Color = RGB(255, 0, 0) Then

                issue = "Missing " & currentworksheet.Cells(2, j).Value & " data"

            ElseIf currentworksheet.Cells(i, j).Interior.Color = RGB(255, 255, 0) Then

                issue = "Zero value in " & currentworksheet.Cells(2, j).Value

            End If

            If issue <> "" Then

                currentworksheet.Cells(nextrow, 9).Value = company
                currentworksheet.Cells(nextrow, 10).Value = benefit
                currentworksheet.Cells(nextrow, 11).Value = issue

                nextrow = nextrow + 1

            End If

        Next j

        ' Check percentage changes, wrote in issue summary
        If IsNumeric(currentworksheet.Cells(i, 5).Value) Then

            issue = ""

            If currentworksheet.Cells(i, 5).Value > 0.3 Then

                issue = "Increase more than 30%"

            ElseIf currentworksheet.Cells(i, 5).Value < -0.3 Then

                issue = "Decrease more than 30%"

            End If

            If issue <> "" Then

                currentworksheet.Cells(nextrow, 9).Value = company
                currentworksheet.Cells(nextrow, 10).Value = benefit
                currentworksheet.Cells(nextrow, 11).Value = issue

                nextrow = nextrow + 1

            End If

        End If

    Next i

    For i = 3 To lastrow

        company = currentworksheet.Cells(i, 1).Value

        ' Only check each company once, find which benefit excel file is missing in which company folder
        If Application.CountIf(currentworksheet.Range("A$2:A" & i), company) = 1 Then

            retirementcount = Application.CountIfs(currentworksheet.Range("A:A"), company, currentworksheet.Range("B:B"), "*Retirement*")

            gratuitycount = Application.CountIfs(currentworksheet.Range("A:A"), company, currentworksheet.Range("B:B"), "*Gratuity*")

            If retirementcount = 0 Then

                currentworksheet.Cells(nextrow, 9).Value = company
                currentworksheet.Cells(nextrow, 10).Value = "Retirement Benefit"
                currentworksheet.Cells(nextrow, 11).Value = "Missing Retirement Benefit Excel file"
                nextrow = nextrow + 1
    
            End If

            If gratuitycount = 0 Then

                currentworksheet.Cells(nextrow, 9).Value = company
                currentworksheet.Cells(nextrow, 10).Value = "Gratuity Benefit"
                currentworksheet.Cells(nextrow, 11).Value = "Missing Gratuity Benefit Excel file"
                nextrow = nextrow + 1

            End If

        End If

    Next i

    With currentworksheet.Range("I1:K" & nextrow - 1)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    MsgBox "Macro Finished"

End Sub
