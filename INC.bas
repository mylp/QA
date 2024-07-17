Attribute VB_Name = "Module2"
Sub CreateTechTabs_INC()
    Dim ws As Worksheet
    Dim techCell As Range
    Dim techDict As Object
    Dim newSheet As Worksheet
    Dim lastRow As Long
    Dim templateSheet As Worksheet
    Dim incWb As Workbook
    Dim incFilePath As String
    Dim ticketNumber As String
    Dim openedDate As String
    Dim updatedDate As String
    Dim currentDate As String
    
    ' Prompt user to select the INC file
    incFilePath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Select INC File")
    If incFilePath = "False" Then Exit Sub ' User canceled file selection
    
    ' Open the INC workbook
    Set incWb = Workbooks.Open(incFilePath)
    
    ' Set the INC worksheet (assuming it's the first sheet)
    Set ws = incWb.Sheets(1)
    
    ' Set the template worksheet
    Set templateSheet = ThisWorkbook.Sheets("Template") ' Change "Template" to your actual template sheet name
    
    ' Create a dictionary to keep track of unique technicians
    Set techDict = CreateObject("Scripting.Dictionary")
    
    ' Find the last row with data in the INC worksheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Set the range for the technician column
    Set techCol = ws.Range("G2:G" & lastRow) ' Assuming tech names are in column G
    
    ' Loop through the technician column to find unique names
    For Each techCell In techCol
        If Not techDict.exists(techCell.Value) Then
            techDict.Add techCell.Value, 1
        End If
    Next techCell
    
    ' Loop through the unique technicians and create a new sheet for each
    For Each techKey In techDict.Keys
        ' Add a new sheet
        Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newSheet.Name = techKey
        
        ' Copy the template to the new sheet
        templateSheet.UsedRange.Copy Destination:=newSheet.Range("A1")
        
        ' Loop through the ticket column to find the ticket for the technician
        For Each techCell In techCol
            If techCell.Value = techKey Then
                ' Get the required values
                ticketNumber = techCell.Offset(0, -6).Value ' Ticket Number (column A)
                openedDate = techCell.Offset(0, -5).Value ' Opened Date (column B)
                updatedDate = techCell.Offset(0, 1).Value ' Updated Date (column H)
                currentDate = Format(Date, "yyyy-mm-dd") ' Current Date
                
                ' Fill the cells in the new sheet
                With newSheet
                    .Cells(2, 3).Value = ticketNumber ' C2
                    .Cells(3, 3).Value = techKey ' C3
                    .Cells(4, 3).Value = techKey ' C4
                    .Cells(2, 5).Value = openedDate ' E2
                    .Cells(3, 5).Value = updatedDate ' E3
                    .Cells(4, 5).Value = currentDate ' E4
                End With
                
                Exit For ' Only fill the first ticket for the technician
            End If
        Next techCell
        
        ' Autofit columns to prevent text cutoff
        newSheet.Columns("A:F").AutoFit
    Next techKey
    
    ' Close the INC workbook
    incWb.Close SaveChanges:=False
    
    MsgBox "Tabs created for each technician!"
End Sub

