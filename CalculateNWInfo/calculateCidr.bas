Attribute VB_Name = "calculateCidr"
Function GetCIDR(IP As String, mask As String) As String
    Dim maskParts() As String
    Dim i As Integer, bitCount As Integer
    
    ' サブネットマスクを分割
    maskParts = Split(mask, ".")
    
    ' 4つの要素があるかチェック
    If UBound(maskParts) <> 3 Then
        GetCIDR = "エラー: 無効なサブネットマスク"
        Exit Function
    End If

    ' CIDRの計算
    bitCount = 0
    For i = 0 To 3
        ' 数値でない場合エラー
        If Not IsNumeric(maskParts(i)) Then
            GetCIDR = "エラー: 無効なマスク"
            Exit Function
        End If

        ' 8ビットごとに1の数をカウント
        bitCount = bitCount + CountBits(CInt(maskParts(i)))
    Next i

    ' CIDR表記で返す
    GetCIDR = IP & "/" & bitCount
End Function

' 8ビットの整数値の中に含まれる1のビット数をカウントする関数
Function CountBits(n As Integer) As Integer
    Dim count As Integer
    count = 0
    While n > 0
        count = count + (n And 1)
        n = n \ 2 ' 右シフト (整数除算)
    Wend
    CountBits = count
End Function

