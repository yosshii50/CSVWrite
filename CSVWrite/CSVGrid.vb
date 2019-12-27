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

        WrkCSVRow(AddRowPos).Wrk()
        CSVCell(2)







        WrkCSVRow(AddRowPos)



        WrkCSVCell.cou()


        If IsNextRow = True Then

        End If

        If WrkData Is Nothing Then
            ReDim WrkData(, )
        Else
            ReDim Preserve WrkData(, )
        End If

        If WrkCSVCol.ColType = CSVCol.ColType_Enum.StrX Then
            Return """" & Me.CellData & """"
        Else
        End If



    End Sub



    'カラムズ
    Private _CSVCols() As CSVCol
    Public ReadOnly Property CSVCols() As CSVCol()
        Get
            Return _CSVCols
        End Get
    End Property

    'カラムの追加
    Public Sub CSVColAdd(ByVal ColTitle As String, ByVal ColType As CSVCol.ColType_Enum, ByVal ColData As String)

        Dim WrkCSVCol As CSVCol

        WrkCSVCol = New CSVCol

        With WrkCSVCol
            .ColTitle = ColTitle
            .ColType = ColType
        End With

        If _CSVCols Is Nothing Then
            ReDim _CSVCols(0)
        Else
            ReDim Preserve _CSVCols(_CSVCols.Count)
        End If

        _CSVCols(_CSVCols.Count - 1) = WrkCSVCol

    End Sub


End Class
