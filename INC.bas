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
    Dim newWb As Workbook
    Dim saveFilePath As String

    ' Prompt user to select the INC file
    incFilePath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx", , "Select INC File")
    If incFilePath = "False" Then Exit Sub ' User canceled file selection

    ' Open the INC workbook
    Set incWb = Workbooks.Open(incFilePath)

    ' Set the INC worksheet (assuming it's the first sheet)
    Set ws = incWb.Sheets(1)

    ' Create a new workbook to store the sheets
    Set newWb = Workbooks.Add

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
        If Not techDict.exists(techCell.Value) And techCell.Value <> "" Then
            techDict.Add techCell.Value, 1
        End If
    Next techCell

    ' Loop through the unique technicians and create a new sheet for each
    For Each techKey In techDict.Keys
        ' Add a new sheet
        Set newSheet = newWb.Sheets.Add(After:=newWb.Sheets(newWb.Sheets.Count))
        newSheet.Name = techKey

        ' Copy the template to the new sheet
        templateSheet.UsedRange.Copy Destination:=newSheet.Range("A1")

        ' Set specific column widths after the template is copied
        With newSheet
            .Columns("A").ColumnWidth = 3.14
            .Columns("B").ColumnWidth = 14.14
            .Columns("C").ColumnWidth = 37.14
            .Columns("D").ColumnWidth = 13.71
            .Columns("E").ColumnWidth = 15
        End With

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
    Next techKey

    ' Save the new workbook with " - Miles" appended to the filename
    Dim baseFileName As String
    baseFileName = Mid(incFilePath, InStrRev(incFilePath, "\") + 1)
    baseFileName = Left(baseFileName, InStrRev(baseFileName, ".") - 1) ' Remove the extension
    saveFilePath = Left(incFilePath, InStrRev(incFilePath, "\")) & baseFileName & " - Miles.xlsx"
    newWb.SaveAs Filename:=saveFilePath
    newWb.Close SaveChanges:=True

    ' Close the INC workbook without saving
    incWb.Close SaveChanges:=False

    MsgBox "Tabs created for each technician with proper column widths!"
End Sub

