Attribute VB_Name = "calculateIP"
Function GetIPRange(IP As String, mask As String) As String
    Dim ipParts() As String, maskParts() As String
    Dim netParts(3) As Integer, broadParts(3) As Integer
    Dim netStr(3) As String, broadStr(3) As String
    Dim i As Integer
    
    ' IP�A�h���X�ƃT�u�l�b�g�}�X�N���u.�v�ŕ���
    ipParts = Split(IP, ".")
    maskParts = Split(mask, ".")

    ' 4�̗v�f���Ȃ��ꍇ�̓G���[
    If UBound(ipParts) <> 3 Or UBound(maskParts) <> 3 Then
        GetIPRange = "�G���[: IP�܂��̓T�u�l�b�g�}�X�N������"
        Exit Function
    End If

    ' �v�Z
    For i = 0 To 3
        ' ���l�łȂ��ꍇ�̓G���[
        If Not IsNumeric(ipParts(i)) Or Not IsNumeric(maskParts(i)) Then
            GetIPRange = "�G���[: ������IP�܂��̓}�X�N"
            Exit Function
        End If

        netParts(i) = CInt(ipParts(i)) And CInt(maskParts(i)) ' �l�b�g���[�N�A�h���X
        broadParts(i) = netParts(i) Or (255 - CInt(maskParts(i))) ' �u���[�h�L���X�g�A�h���X

        ' ������Ƃ��Ċi�[
        netStr(i) = CStr(netParts(i))
        broadStr(i) = CStr(broadParts(i))
    Next i

    ' ���ʂ�Ԃ��i������Ƃ��Č����j
    GetIPRange = Join(netStr, ".") & " �` " & Join(broadStr, ".")
End Function

