Imports Microsoft.Office.Interop

Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim start_time = Now()
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
            Return
        End If


        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        'xlWorkSheet.Cells(1, 1) = "Sheet 1 content"

        For index As Integer = 1 To 1000
            Dim myvalue As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"}
            Dim range = xlWorkSheet.Range(String.Format("A{0}", index), String.Format("L{0}", index))
            range.Value = myvalue
        Next index

        xlApp.DisplayAlerts = False
        xlWorkBook.SaveAs("d:\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, _
         Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        xlWorkBook.Close(True, misValue, misValue)
        xlApp.Quit()

        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)
        Dim end_time = Now()

        Console.WriteLine((end_time - start_time).TotalSeconds())
        'MessageBox.Show("Excel file created , you can find the file d:\csharp-Excel.xls")
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim start_time = Now()

        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()

        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!!")
            Return
        End If


        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value

        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        'xlWorkSheet.Cells(1, 1) = "Sheet 1 content"

        For index As Integer = 1 To 1000
            'Dim myvalue As String() = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"}
            'Dim range = xlWorkSheet.Range(String.Format("A{0}", index), String.Format("L{0}", index))
            'range.Value = myvalue
            xlWorkSheet.Cells(index, 1) = "Sheet 1 content"
            xlWorkSheet.Cells(index, 2) = "Sheet 1 content"
            xlWorkSheet.Cells(index, 3) = "Sheet 1 content"
            xlWorkSheet.Cells(index, 4) = "Sheet 1 content"
            xlWorkSheet.Cells(index, 5) = "Sheet 1 content"
            xlWorkSheet.Cells(index, 6) = "Sheet 1 content"
            xlWorkSheet.Cells(index, 7) = "Sheet 1 content"
            xlWorkSheet.Cells(index, 8) = "Sheet 1 content"
            xlWorkSheet.Cells(index, 9) = "Sheet 1 content"
            xlWorkSheet.Cells(index, 10) = "Sheet 1 content"
            xlWorkSheet.Cells(index, 11) = "Sheet 1 content"
            xlWorkSheet.Cells(index, 12) = "Sheet 1 content"
        Next index

        xlApp.DisplayAlerts = False
        xlWorkBook.SaveAs("d:\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, _
         Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
        xlWorkBook.Close(True, misValue, misValue)
        xlApp.Quit()

        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)

        Dim end_time = Now()

        Console.WriteLine((end_time - start_time).TotalSeconds())

        'MessageBox.Show("Excel file created , you can find the file d:\csharp-Excel.xls")
    End Sub
End Class
