Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class CSVGrid

    Private _CSVCol() As CSVCol 'Idx=1から使用
    Private Structure CSVCells
        Dim _CSVCell() As CSVCell 'Idx=1から使用
    End Structure
    Private _CSVRow() As CSVCells 'Idx=1から使用

    Public Sub Clear()
        Call ClearCol()
        Call ClearRow()
    End Sub

    Public Sub ClearCol()
        ReDim _CSVCol(0)
        _CSVCol(0) = New CSVCol
        AddColPos = 0
    End Sub

    Public Sub ClearRow()
        ReDim _CSVRow(0)
        AddRowPos = 0
    End Sub

    Public Sub NextRow()
        AddColPos = 0
        AddRowPos = AddRowPos + 1
    End Sub

    Private AddRowPos As Integer = 0 '1から使用
    Private AddColPos As Integer = 0 '1から使用
    Public Sub AddCell(ByVal Data As Decimal, ByVal ColTitle As String)
        Call AddCell(Data.ToString, ColTitle)
    End Sub
    Public Sub AddCell(ByVal Data As String)
        Call AddCell(Data, "")
    End Sub
    Public Sub AddCell(ByVal Data As String, ByVal ColTitle As String)

        '------------------------------------------------------------------------------- Rowの領域作成

        If AddRowPos = 0 Then
            AddRowPos = 1
        End If

        If _CSVRow Is Nothing Then
            ReDim _CSVRow(0)
        End If
        If _CSVRow.Count - 1 < AddRowPos Then
            '領域が不足している場合
            ReDim Preserve _CSVRow(AddRowPos)
            ReDim _CSVRow(AddRowPos)._CSVCell(0)
        End If

        '--------------------------------------------------------------------------- Col位置を取得
        AddColPos = AddCellNextPos(AddColPos, ColTitle, _CSVCol)

        '------------------------------------------------------------------------------- Colの領域作成

        If _CSVCol.Count - 1 < AddColPos Then
            '領域が不足している場合

            ReDim Preserve _CSVCol(AddColPos)

            _CSVCol(AddColPos) = New CSVCol
            _CSVCol(AddColPos).ColTitle = ColTitle

        End If

        '------------------------------------------------------------------------------- Cellの領域作成

        If _CSVRow(AddRowPos)._CSVCell.Count - 1 < AddColPos Then
            '領域が不足している場合
            ReDim Preserve _CSVRow(AddRowPos)._CSVCell(AddColPos)
            _CSVRow(AddRowPos)._CSVCell(AddColPos) = New CSVCell
        Else
            If _CSVRow(AddRowPos)._CSVCell(AddColPos) Is Nothing Then
                _CSVRow(AddRowPos)._CSVCell(AddColPos) = New CSVCell
            End If
        End If

        'データセット
        _CSVRow(AddRowPos)._CSVCell(AddColPos).CellData = Data

    End Sub

    'Col位置を取得
    Private Function AddCellNextPos(ByVal WrkColPos As Integer, ByVal WrkColTitle As String, ByRef WrkCSVCol() As CSVCol) As Integer

        WrkColPos = WrkColPos + 1

        '------------------------------------------------------------------------------- Colの位置判定
        If WrkColTitle = "" Then
            'タイトルなし → 自動で次の位置に配置
            Return WrkColPos
        Else

            If WrkCSVCol Is Nothing Then
                ReDim WrkCSVCol(0)
                WrkCSVCol(0) = New CSVCol
            End If
            If WrkCSVCol.Count - 1 >= WrkColPos Then
                '領域範囲内の場合

                If WrkCSVCol(WrkColPos).ColTitle = WrkColTitle Then
                    'タイトル一致 → その場所に配置
                    Return WrkColPos
                Else
                    'タイトル違い → 自動検索で配置

                    '同じタイトルの位置を取得
                    Return SerchTitlePos(WrkColPos, WrkColTitle, WrkCSVCol)

                End If

            Else
                '領域範囲外の場合

                '同じタイトルの位置を取得
                Return SerchTitlePos(WrkColPos, WrkColTitle, WrkCSVCol)

            End If

        End If

        '同じタイトルを検索

    End Function

    '同じタイトルの位置を取得
    Private Function SerchTitlePos(ByVal WrkColPos As Integer, ByVal WrkColTitle As String, ByVal WrkCSVCol() As CSVCol) As Integer

        Dim WrkIdx As Integer

        For WrkIdx = 1 To WrkCSVCol.Count - 1
            If WrkCSVCol(WrkIdx).ColTitle = WrkColTitle Then
                Exit For
            End If
        Next

        If WrkIdx = WrkCSVCol.Count Then
            '同じタイトルが見つからなかった場合
            Return WrkIdx
        Else
            '同じタイトルが見つかった場合
            Return WrkIdx
        End If

    End Function

    'CSV文字列取得
    Public Function GetCSV() As String

        Dim WrkStr As String = ""

        'タイトル部分
        For WrkCol As Integer = 1 To _CSVCol.Count - 1
            If Not _CSVCol(WrkCol) Is Nothing Then
                WrkStr = WrkStr & """" & _CSVCol(WrkCol).ColTitle & """"
            End If
            If WrkCol <> _CSVCol.Count - 1 Then
                WrkStr = WrkStr & ","
            End If
        Next
        WrkStr = WrkStr & vbCrLf

        '内容部分
        For WrkRow As Integer = 1 To AddRowPos
            '行数分繰り返し

            If _CSVRow.Count = WrkRow Then
                '最終行を超えたら終了
                Exit For
            End If

            For WrkCol As Integer = 1 To _CSVCol.Count - 1
                '項目数分繰り返し

                If _CSVRow(WrkRow)._CSVCell.Count > WrkCol Then
                    If Not _CSVRow(WrkRow)._CSVCell(WrkCol) Is Nothing Then

                        If _CSVCol(WrkCol).ColType = CSVCol.ColType_Enum.Num9 Then
                            WrkStr = WrkStr & _CSVRow(WrkRow)._CSVCell(WrkCol).CellData
                        Else
                            WrkStr = WrkStr & """" & _CSVRow(WrkRow)._CSVCell(WrkCol).CellData & """"
                        End If

                    End If
                End If
                If WrkCol <> _CSVCol.Count - 1 Then
                    WrkStr = WrkStr & ","
                End If

            Next
            WrkStr = WrkStr & vbCrLf
        Next

        Return WrkStr
    End Function

    Public Function SaveCSV(ByVal WrkCSVFileName As String) As Boolean

        Dim WrkFile As System.IO.StreamWriter = Nothing

        If FileOutOpen(WrkFile, WrkCSVFileName) = False Then
            Return False
        End If


        WrkFile.Write(Me.GetCSV())
        WrkFile.Close()

        Return True
    End Function

    'CSV保存用にオープン
    Private Function FileOutOpen(ByRef File As System.IO.StreamWriter, ByVal FilePathName As String) As Boolean

        Dim PathName As String

        PathName = System.IO.Path.GetDirectoryName(FilePathName) 'パス部分のみ取得

        'フォルダが存在するか確認
        If System.IO.Directory.Exists(PathName) Then
            '存在する場合
            '何もしない
        Else
            '存在しない場合

            Try

                'フォルダ作成
                System.IO.Directory.CreateDirectory(PathName)

            Catch ex As Exception
                Call MessageBox.Show(ex.ToString, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop)

                '再帰的呼び出し
                '↑自動的に再帰作成されるので、処理不要

            End Try

        End If

        Try
            'File = My.Computer.FileSystem.OpenTextFileWriter(FilePathName, False, System.Text.Encoding.Default)
            File = New System.IO.StreamWriter(FilePathName, False, System.Text.Encoding.GetEncoding("shift_jis"))

        Catch ex As System.IO.IOException
            Call MessageBox.Show("ファイルへの書き込みに失敗しました。" & vbCrLf & vbCrLf & ex.ToString, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Return False
        Catch ex As Exception
            Call MessageBox.Show(ex.ToString, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Return False
        End Try

        Return True
    End Function

End Class
