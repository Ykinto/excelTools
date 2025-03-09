Attribute VB_Name = "calculateIP"
Function GetIPRange(IP As String, mask As String) As String
    Dim ipParts() As String, maskParts() As String
    Dim netParts(3) As Integer, broadParts(3) As Integer
    Dim netStr(3) As String, broadStr(3) As String
    Dim i As Integer
    
    ' IPアドレスとサブネットマスクを「.」で分割
    ipParts = Split(IP, ".")
    maskParts = Split(mask, ".")

    ' 4つの要素がない場合はエラー
    If UBound(ipParts) <> 3 Or UBound(maskParts) <> 3 Then
        GetIPRange = "エラー: IPまたはサブネットマスクが無効"
        Exit Function
    End If

    ' 計算
    For i = 0 To 3
        ' 数値でない場合はエラー
        If Not IsNumeric(ipParts(i)) Or Not IsNumeric(maskParts(i)) Then
            GetIPRange = "エラー: 無効なIPまたはマスク"
            Exit Function
        End If

        netParts(i) = CInt(ipParts(i)) And CInt(maskParts(i)) ' ネットワークアドレス
        broadParts(i) = netParts(i) Or (255 - CInt(maskParts(i))) ' ブロードキャストアドレス

        ' 文字列として格納
        netStr(i) = CStr(netParts(i))
        broadStr(i) = CStr(broadParts(i))
    Next i

    ' 結果を返す（文字列として結合）
    GetIPRange = Join(netStr, ".") & " ～ " & Join(broadStr, ".")
End Function

