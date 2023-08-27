Imports Word = Microsoft.Office.Interop.Word
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form1
    Dim sfd As SaveFileDialog = New SaveFileDialog()

    Private Sub btnWord_Click(sender As Object, e As EventArgs) Handles btnWord.Click
        If sfd.ShowDialog() = DialogResult.OK Then
            Dim word As Word.Application = New Word.Application()
            Dim doc As Word.Document = word.Documents.Add()
            doc.Activate()
            word.Selection.TypeText(tbTexto.Text)
            doc.SaveAs2(sfd.FileName)
            word.Visible = True
        End If
    End Sub

    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        If sfd.ShowDialog() = DialogResult.OK Then
            Dim excel As Excel.Application = New Excel.Application()
            Dim file As Excel.Workbook = excel.Workbooks.Add()
            Dim hoja As Excel.Worksheet = file.Worksheets.Add()
            file.Activate()
            hoja.Cells(1, 1).Value = tbTexto.Text
            file.SaveAs(sfd.FileName)
            excel.Visible = True
        End If
    End Sub
End Class
