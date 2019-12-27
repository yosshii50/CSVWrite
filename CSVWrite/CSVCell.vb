Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class CSVCell

    'セルデータ
    Public CellData As String

    '参照セル
    Private _RefCell As CSVCell
    Public Property RefCell() As CSVCell
        Get
            Return _RefCell
        End Get
        Set(ByVal value As CSVCell)
            _RefCell = value
        End Set
    End Property

    'CSV用データ生成
    Public Function CSVData() As String
        If Not _RefCell Is Nothing Then
            Return _RefCell.CSVData()
        Else
            Return Me.CellData
        End If
    End Function

End Class
