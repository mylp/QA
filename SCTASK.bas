Sub CreateTechTabs_SCTASK()
    Dim ws As Worksheet
    Dim techCell As Range
    Dim techDict As Object
    Dim newSheet As Worksheet
    Dim lastRow As Long
    Dim templateSheet As Worksheet
    Dim sctaskWb As Workbook
    Dim sctaskFilePath As String
    Dim ticketNumber As String
    Dim closedDate As String
    Dim currentDate As String
    Dim sheetName As String
    Dim counter As Integer
    
    ' Prompt user to select the SCTASK file
    sctaskFilePath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Select SCTASK File")
    If sctaskFilePath = "False" Then Exit Sub ' User canceled file selection
    
    ' Open the SCTASK workbook
    Set sctaskWb = Workbooks.Open(sctaskFilePath)
    
    ' Set the SCTASK worksheet (assuming it's the first sheet)
    Set ws = sctaskWb.Sheets(1)
    
    ' Set the template worksheet
    Set templateSheet = ThisWorkbook.Sheets("Template") ' Change "Template" to your actual template sheet name
    
    ' Create a dictionary to keep track of unique technicians
    Set techDict = CreateObject("Scripting.Dictionary")
    
    ' Find the last row with data in the SCTASK worksheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Set the range for the technician column
    Set techCol = ws.Range("D2:D" & lastRow) ' Assuming tech names are in column D
    
    ' Loop through the technician column to find unique names
    For Each techCell In techCol
        If Not techDict.exists(techCell.Value) Then
            techDict.Add techCell.Value, 1
        End If
    Next techCell
    
    ' Loop through the unique technicians and create a new sheet for each
    For Each techKey In techDict.Keys
        ' Generate a unique sheet name
        sheetName = techKey
        counter = 1
        Do While SheetExists(sheetName)
            sheetName = techKey & "_" & counter
            counter = counter + 1
        Loop
        
        ' Add a new sheet
        Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newSheet.Name = sheetName
        
        ' Copy the template to the new sheet
        templateSheet.UsedRange.Copy Destination:=newSheet.Range("A1")
        
        ' Loop through the ticket column to find the ticket for the technician
        For Each techCell In techCol
            If techCell.Value = techKey Then
                ' Get the required values
                ticketNumber = techCell.Offset(0, -3).Value ' Ticket Number (column A)
                closedDate = techCell.Offset(0, 3).Value ' Closed Date (column G)
                currentDate = Format(Date, "yyyy-mm-dd") ' Current Date
                
                ' Fill the cells in the new sheet
                With newSheet
                    .Cells(2, 3).Value = ticketNumber ' C2
                    .Cells(3, 3).Value = techKey ' C3
                    .Cells(4, 3).Value = techKey ' C4
                    .Cells(3, 5).Value = closedDate ' E3
                    .Cells(4, 5).Value = currentDate ' E4
                End With
                
                Exit For ' Only fill the first ticket for the technician
            End If
        Next techCell
        
        ' Autofit columns to prevent text cutoff
        newSheet.Columns("A:F").AutoFit
    Next techKey
    
    ' Close the SCTASK workbook
    sctaskWb.Close SaveChanges:=False
    
    MsgBox "Tabs created for each technician!"
End Sub

' Function to check if a sheet with the given name already exists
Function SheetExists(sheetName As String) As Boolean
    Dim sheet As Worksheet
    SheetExists = False
    For Each sheet In ThisWorkbook.Sheets
        If sheet.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next sheet
End Function
