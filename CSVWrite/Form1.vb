Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim WrkCSV As New CSVGrid

        'WrkCSV.NextRow() '←これはあってもなくてもＯＫ
        WrkCSV.AddCell("1-1", "Title1")
        WrkCSV.AddCell("1-2", "Title2")
        WrkCSV.AddCell("1-3", "Title3")

        WrkCSV.NextRow()
        WrkCSV.AddCell("2-1", "Title1")
        WrkCSV.AddCell("2-2", "Title2")
        WrkCSV.AddCell("2-3") 'タイトルなし → 自動で次の位置に配置

        WrkCSV.NextRow()
        WrkCSV.AddCell("3-1") 'タイトルなし → 自動で次の位置に配置
        WrkCSV.AddCell("3-3", "Title3") 'タイトル違い → 自動検索で配置
        WrkCSV.AddCell("3-2A", "Title2") 'タイトル違い → 自動検索で配置
        WrkCSV.AddCell("3-2B", "Title2") '重複 → 上書き

        WrkCSV.NextRow()
        WrkCSV.AddCell("4-3", "Title3") 'タイトル違い → 自動検索で配置
        WrkCSV.AddCell("4-4") 'タイトルなし → 自動で次の位置に配置
        WrkCSV.AddCell("4-5", "") 'タイトルなし → 自動で次の位置に配置

        WrkCSV.NextRow()
        WrkCSV.AddCell("5-3", "Title3") 'タイトル違い → 自動検索で配置
        WrkCSV.AddCell("5-6", "Title6") 'タイトル違い → 自動検索で配置 → 存在しないため新規
        WrkCSV.AddCell("5-7", "") 'タイトルなし → 自動で次の位置に配置

        WrkCSV.NextRow()
        WrkCSV.AddCell("6-8", "Title8") 'タイトル違い → 自動検索で配置
        WrkCSV.AddCell("6-9", "Title9") 'タイトル違い → 自動検索で配置 → 存在しないため新規

        WrkCSV.NextRow()

        TextBox1.Text = WrkCSV.GetCSV()

        '想定しているデータの表示

    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim WrkStr As String = ""

        WrkStr = WrkStr & "Title1,Title2,Title3,,,Title6,,Title8,Title9" & vbCrLf
        WrkStr = WrkStr & "1-1,1-2,1-3,,,,,," & vbCrLf
        WrkStr = WrkStr & "2-1,2-2,2-3,,,,,," & vbCrLf
        WrkStr = WrkStr & "3-1,3-2B,3-3,,,,,," & vbCrLf
        WrkStr = WrkStr & ",,4-3,4-4,4-5,,,," & vbCrLf
        WrkStr = WrkStr & ",,5-3,,,5-6,5-7,," & vbCrLf
        WrkStr = WrkStr & ",,,,,,,6-8,6-9" & vbCrLf

        TextBox2.Text = WrkStr

        Button1.PerformClick()

    End Sub
End Class