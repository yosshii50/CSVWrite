Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class TestWriteForm

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim StTime As Date = Now()
        Dim button = CType(sender, Button)

        Dim WrkStr As String = ""
        For WrkIdx As Integer = 1 To 100000
            WrkStr = WrkStr & "aaaa" & vbCrLf
        Next

        '[CSV]のデータ出力
        Dim sw As New System.IO.StreamWriter(TxtSavePath.Text & button.Name & ".CSV", False, System.Text.Encoding.GetEncoding("shift_jis"))
        sw.Write(WrkStr)
        sw.Close()

        button.Text = (Now - StTime).ToString

        '18秒

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim StTime As Date = Now()
        Dim button = CType(sender, Button)

        '[CSV]のデータ出力
        Dim sw As New System.IO.StreamWriter(TxtSavePath.Text & button.Name & ".CSV", False, System.Text.Encoding.GetEncoding("shift_jis"))
        For WrkIdx As Integer = 1 To 100000
            sw.Write("aaaa" & vbCrLf)
        Next
        sw.Close()

        button.Text = (Now - StTime).ToString

        '0.01秒

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim StTime As Date = Now()
        Dim button = CType(sender, Button)

        '[CSV]のデータ出力
        Dim sw As New System.IO.StreamWriter(TxtSavePath.Text & button.Name & ".CSV", False, System.Text.Encoding.GetEncoding("shift_jis"))
        For WrkRIdx As Integer = 1 To 100000
            For WrkCIdx As Integer = 1 To 100
                sw.Write("aaaa,")
            Next
            sw.Write(vbCrLf)
        Next
        sw.Close()

        button.Text = (Now - StTime).ToString

        '0.6秒

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Dim StTime As Date = Now()
        Dim button = CType(sender, Button)

        '[CSV]のデータ出力
        Dim sw As New System.IO.StreamWriter(TxtSavePath.Text & button.Name & ".CSV", False, System.Text.Encoding.GetEncoding("shift_jis"))
        For WrkRIdx As Integer = 1 To 100000
            Dim WrkStr As String = ""
            For WrkCIdx As Integer = 1 To 100
                WrkStr = WrkStr & "aaaa,"
            Next
            sw.Write(WrkStr & vbCrLf)
        Next
        sw.Close()

        button.Text = (Now - StTime).ToString

        '1.8秒
        'こっちの方が遅い

    End Sub

End Class
