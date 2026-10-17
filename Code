
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
    targetWS.Cells.ClearContents

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

    ' Find Excel files
        fileName = Dir(mainFolder & folder & "/*.xls*")
        
    'Loop through all Excel files in a company folder
        Do While fileName <> ""

            'Only Find Excel files with 1 or 2 in the name
            If (InStr(fileName, "1") > 0 Or InStr(fileName, "2") > 0) Then
            
                ' Open Excel file
                Set sourceWB = Workbooks.Open(mainFolder & folder & "/" & fileName)

                ' Find next empty row in the current worksheet
                nextRow = targetWS.Cells(targetWS.Rows.Count, 1).End(xlUp).Row + 1

                If targetWS.Range("A1").Value = "" Then
                    nextRow = 1
                End If
                
                'Loop through all the sheets in Excel file to copy and paste different values.
                For Each ws In sourceWB.Worksheets

                   If InStr(ws.Name, "1") > 0 Then

                       ws.Range("A1").Copy Destination:=targetWS.Cells(nextRow, 1)
                       
                       nextRow = nextRow + 1

                   ElseIf InStr(ws.Name, "2") > 0 Then

                       ws.Range("A1:D4").Copy Destination:=targetWS.Cells(nextRow, 1)
                       
                       nextRow = nextRow + 4

                   End If

               Next ws
            
                ' Close Excel file
                sourceWB.Close SaveChanges:=False
                
                End If

            ' Next file name in the Company folder
            fileName = Dir()

        Loop
        
    ' Next Company folder name in the folders
    Next folder
   MsgBox "Macro"
End Sub

