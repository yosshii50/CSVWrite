Option Explicit On '型宣言を強制
Option Strict On 'タイプ変換を厳密に

Public Class CSVCol

    'カラムタイトル
    Private _ColTitle As String
    Public Property ColTitle() As String
        Get
            Return _ColTitle
        End Get
        Set(ByVal value As String)
            _ColTitle = value
        End Set
    End Property

    'カラムタイプ
    Public Enum ColType_Enum
        Auto '自動判断
        StrX '文字/["]で括る
        Num9 '数字/["]で括らない
    End Enum
    Private _ColType As ColType_Enum = ColType_Enum.Auto '自動判断
    Public Property ColType() As ColType_Enum
        Get
            Return _ColType
        End Get
        Set(ByVal value As ColType_Enum)
            _ColType = value
        End Set
    End Property

    '数字以外の存在 False:数字のみ True:数字以外あり
    Private _StrExists As Boolean = False
    Public Property StrExists() As Boolean
        Get
            Return _StrExists
        End Get
        Set(ByVal value As Boolean)
            _StrExists = value
        End Set
    End Property

End Class
