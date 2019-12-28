Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class CSVGrid

    Private WrkCSVCol() As CSVCol 'Idx=1から使用
    Private Structure CSVCells
        Dim WrkCSVCell() As CSVCell 'Idx=1から使用
    End Structure
    Private WrkCSVRow() As CSVCells 'Idx=1から使用

    Public Sub Clear()
        Call ClearCol()
        Call ClearRow()
    End Sub

    Public Sub ClearCol()
        Erase WrkCSVCol
        AddColPos = 0
    End Sub

    Public Sub ClearRow()
        Erase WrkCSVRow
        AddRowPos = 0
    End Sub

    Public Sub NextRow()
        AddColPos = 0
        AddRowPos = AddRowPos + 1
    End Sub

    Private AddRowPos As Integer = 0 '1から使用
    Private AddColPos As Integer = 0 '1から使用
    Public Sub AddCell(ByVal Data As String, ByVal ColTitle As String)

        '------------------------------------------------------------------------------- Rowの領域作成

        If AddRowPos = 0 Then
            AddRowPos = 1
        End If

        If AddRowPos = 1 And AddColPos = 0 Then
            '最初の場合
            ReDim WrkCSVRow(AddRowPos)
        End If

        If WrkCSVRow.Count - 1 < AddRowPos Then
            ReDim Preserve WrkCSVRow(AddRowPos)
        End If

        '------------------------------------------------------------------------------- Colの領域作成

        If WrkCSVCol Is Nothing Then
            AddColPos = 1
            ReDim WrkCSVCol(AddColPos)

            WrkCSVCol(AddColPos) = New CSVCol
            WrkCSVCol(AddColPos).ColTitle = ColTitle

        Else
            AddColPos = AddColPos + 1

            If WrkCSVCol.Count - 1 < AddColPos Then
                '領域が不足している場合

                ReDim Preserve WrkCSVCol(AddColPos)

                WrkCSVCol(AddColPos) = New CSVCol
                WrkCSVCol(AddColPos).ColTitle = ColTitle

            End If

        End If

        'タイトルが一致しているか確認

        '一致していない場合
        '次のCol領域を検索

        '------------------------------------------------------------------------------- Cellの領域作成

        If WrkCSVRow(AddRowPos).WrkCSVCell Is Nothing Then
            ReDim WrkCSVRow(AddRowPos).WrkCSVCell(AddColPos) '←こっちは省略できるかも ※※※
        Else
            ReDim Preserve WrkCSVRow(AddRowPos).WrkCSVCell(AddColPos)
        End If

        WrkCSVRow(AddRowPos).WrkCSVCell(AddColPos) = New CSVCell

        'データセット
        WrkCSVRow(AddRowPos).WrkCSVCell(AddColPos).CellData = Data

    End Sub

    'CSV文字列取得
    Public Function GetCSV() As String

        Dim WrkStr As String = ""

        For Each WrkCol As CSVCol In WrkCSVCol
            If Not WrkCol Is Nothing Then
                WrkStr = WrkStr & WrkCol.ColTitle & ","
            End If
        Next
        WrkStr = WrkStr & vbCrLf

        For WrkRow As Integer = 1 To AddRowPos
            For Each WrkCell As CSVCell In WrkCSVRow(WrkRow).WrkCSVCell
                If Not WrkCell Is Nothing Then
                    WrkStr = WrkStr & WrkCell.CellData & ","
                End If
            Next
            WrkStr = WrkStr & vbCrLf
        Next

        Return WrkStr
    End Function

End Class
