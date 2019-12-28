Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim WrkCSV As New CSVGrid

        WrkCSV.AddCell("Data1-1", "Title1")
        WrkCSV.AddCell("Data1-2", "Title2")

        WrkCSV.NextRow()
        WrkCSV.AddCell("Data2-1", "Title1")
        WrkCSV.AddCell("Data2-2", "Title2")

        TextBox1.Text = WrkCSV.GetCSV()

    End Sub
End Class