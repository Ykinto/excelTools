Attribute VB_Name = "calculateCidr"
Function GetCIDR(IP As String, mask As String) As String
    Dim maskParts() As String
    Dim i As Integer, bitCount As Integer
    
    ' �T�u�l�b�g�}�X�N�𕪊�
    maskParts = Split(mask, ".")
    
    ' 4�̗v�f�����邩�`�F�b�N
    If UBound(maskParts) <> 3 Then
        GetCIDR = "�G���[: �����ȃT�u�l�b�g�}�X�N"
        Exit Function
    End If

    ' CIDR�̌v�Z
    bitCount = 0
    For i = 0 To 3
        ' ���l�łȂ��ꍇ�G���[
        If Not IsNumeric(maskParts(i)) Then
            GetCIDR = "�G���[: �����ȃ}�X�N"
            Exit Function
        End If

        ' 8�r�b�g���Ƃ�1�̐����J�E���g
        bitCount = bitCount + CountBits(CInt(maskParts(i)))
    Next i

    ' CIDR�\�L�ŕԂ�
    GetCIDR = IP & "/" & bitCount
End Function

' 8�r�b�g�̐����l�̒��Ɋ܂܂��1�̃r�b�g�����J�E���g����֐�
Function CountBits(n As Integer) As Integer
    Dim count As Integer
    count = 0
    While n > 0
        count = count + (n And 1)
        n = n \ 2 ' �E�V�t�g (�������Z)
    Wend
    CountBits = count
End Function

