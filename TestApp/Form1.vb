Imports Microsoft.Office.Interop.Excel

Public Class Form1
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim oXL As Application
        Dim oWB As Workbook
        Dim oSheet As Worksheet
        Dim oRng As Range

        ' Start Excel and get Application object.
        oXL = CreateObject("Excel.Application")
        oXL.Visible = True

        ' Get a new workbook.
        oWB = oXL.Workbooks.Add
        oSheet = oWB.ActiveSheet

        ' Add table headers going cell by cell.
        oSheet.Cells(1, 1).Value = "First Name"
        oSheet.Cells(1, 2).Value = "Last Name"
        oSheet.Cells(1, 3).Value = "Full Name"
        oSheet.Cells(1, 4).Value = "Salary"

        ' Format A1:D1 as bold, vertical alignment = center.
        With oSheet.Range("A1", "D1")
            .Font.Bold = True
            .VerticalAlignment = XlVAlign.xlVAlignCenter
        End With

        ' Create an array to set multiple values at once.
        Dim saNames(5, 2) As String
        saNames(0, 0) = "John"
        saNames(0, 1) = "Smith"
        saNames(1, 0) = "Tom"
        saNames(1, 1) = "Brown"
        saNames(2, 0) = "Sue"
        saNames(2, 1) = "Thomas"
        saNames(3, 0) = "Jane"

        saNames(3, 1) = "Jones"
        saNames(4, 0) = "Adam"
        saNames(4, 1) = "Johnson"

        ' Fill A2:B6 with an array of values (First and Last Names).
        oSheet.Range("A2", "B6").Value = saNames

        ' Fill C2:C6 with a relative formula (=A2 & " " & B2).
        oRng = oSheet.Range("C2", "C6")
        oRng.Formula = "=A2 & "" "" & B2"

        ' Fill D2:D6 with a formula(=RAND()*100000) and apply format.
        oRng = oSheet.Range("D2", "D6")
        oRng.Formula = "=RAND()*100000"
        oRng.NumberFormat = "$0.00"

        ' AutoFit columns A:D.
        oRng = oSheet.Range("A1", "D1")
        oRng.EntireColumn.AutoFit()

        ' Manipulate a variable number of columns for Quarterly Sales Data.
        Call DisplayQuarterlySales(oSheet)

        ' Make sure Excel is visible and give the user control
        ' of Excel's lifetime.
        oXL.Visible = True
        oXL.UserControl = True

        ' Make sure that you release object references.
        oRng = Nothing
        oSheet = Nothing
        oWB = Nothing
        oXL.Quit()
        oXL = Nothing

        Exit Sub
Err_Handler:
        MsgBox(Err.Description, vbCritical, "Error: " & Err.Number)
    End Sub

    Private Sub DisplayQuarterlySales(ByVal oWS As Worksheet)
        Dim oResizeRange As Range
        Dim oChart As Chart
        Dim oSeries As Series
        Dim iNumQtrs As Integer
        Dim sMsg As String
        Dim iRet As Integer


        ' Determine how many quarters to display data for.
        For iNumQtrs = 4 To 2 Step -1
            sMsg = "Enter sales data for" & Str(iNumQtrs) & " quarter(s)?"
            iRet = MsgBox(sMsg, vbYesNo Or vbQuestion _
               Or vbMsgBoxSetForeground, "Quarterly Sales")
            If iRet = vbYes Then Exit For
        Next iNumQtrs

        ' Starting at E1, fill headers for the number of columns selected.
        oResizeRange = oWS.Range("E1", "E1").Resize(ColumnSize:=iNumQtrs)
        oResizeRange.Formula = "=""Q"" & COLUMN()-4 & CHAR(10) & ""Sales"""

        ' Change the Orientation and WrapText properties for the headers.
        oResizeRange.Orientation = 38
        oResizeRange.WrapText = True

        ' Fill the interior color of the headers.
        oResizeRange.Interior.ColorIndex = 36

        ' Fill the columns with a formula and apply a number format.
        oResizeRange = oWS.Range("E2", "E6").Resize(ColumnSize:=iNumQtrs)
        oResizeRange.Formula = "=RAND()*100"
        oResizeRange.NumberFormat = "$0.00"

        ' Apply borders to the Sales data and headers.
        oResizeRange = oWS.Range("E1", "E6").Resize(ColumnSize:=iNumQtrs)
        oResizeRange.Borders.Weight = XlBorderWeight.xlThin

        ' Add a Totals formula for the sales data and apply a border.
        oResizeRange = oWS.Range("E8", "E8").Resize(ColumnSize:=iNumQtrs)
        oResizeRange.Formula = "=SUM(E2:E6)"
        With oResizeRange.Borders(XlBordersIndex.xlEdgeBottom)
            .LineStyle = XlLineStyle.xlDouble
            .Weight = XlBorderWeight.xlThick
        End With

        ' Add a Chart for the selected data.
        oResizeRange = oWS.Range("E2:E6").Resize(ColumnSize:=iNumQtrs)
        oResizeRange = oWS.Range("E1:H6")
        oChart = oWS.Parent.Charts.Add

        With oChart
            .ChartType = XlChartType.xlColumnClustered
            .SetSourceData(oResizeRange)
            For iRet = 1 To 5
                .SeriesCollection(iRet).Name = "=Tabelle1!$C$" & iRet + 1
            Next iRet
            .Location(XlChartLocation.xlLocationAsObject, oWS.Name)
        End With

        ' Move the chart so as not to cover your data.
        With oWS.Shapes.Item("Chart 1")
            .Top = oWS.Rows(10).Top
            .Left = oWS.Columns(2).Left
        End With

        ' Free any references.
        oChart = Nothing
        oResizeRange = Nothing
Err_Handler:
        If Err.Number <> 0 Then MsgBox(Err.Description, vbCritical, "Error: " & Err.Number)
    End Sub

End Class
