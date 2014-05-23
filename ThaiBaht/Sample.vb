Public Class Sample

    Public Shared Function ThaiBaht(ByVal pAmount As Double) As String
        If pAmount = 0 Then
            Return "�ٹ��ҷ��ǹ"
        End If

        Dim _integerValue As String ' �ӹǹ���
        Dim _decimalValue As String ' �ȹ���
        Dim _integerTranslatedText As String ' �ӹǹ��� ������
        Dim _decimalTranslatedText As String ' �ȹ���������

        _integerValue = Format(pAmount, "####.00") ' �Ѵ Format ����Թ�繵���Ţ 2 ��ѡ
        _decimalValue = Mid(_integerValue, Len(_integerValue) - 1, 2) ' �ȹ���
        _integerValue = Mid(_integerValue, 1, Len(_integerValue) - 3) ' �ӹǹ���

        ' �ŧ �ӹǹ��� �� ������
        _integerTranslatedText = NumberToText(CDbl(_integerValue))

        ' �ŧ �ȹ��� �� ������
        If CDbl(_decimalValue) = 0 Then
            _decimalTranslatedText = NumberToText(CDbl(_decimalValue))
        Else
            _decimalTranslatedText = ""
        End If

        ' �������շȹ��
        If _decimalTranslatedText.Trim = "" Then
            _integerTranslatedText += "�ҷ��ǹ"
        Else
            _integerTranslatedText += "�ҷ" & _decimalTranslatedText & "ʵҧ��"
        End If

        Return _integerTranslatedText
    End Function


    Private Shared Function NumberToText(ByVal pAmount As Double) As String
        ' ����ѡ��
        Dim _numberText() As String = {"", "˹��", "�ͧ", "���", "���", "���", "ˡ", "��", "Ỵ", "���", "�Ժ"}

        ' ��ѡ ˹��� �Ժ ���� �ѹ ...
        Dim _digit() As String = {"", "�Ժ", "����", "�ѹ", "����", "�ʹ", "��ҹ"}
        Dim _value As String, _aWord As String, _text As String
        Dim _numberTranslatedText As String = ""
        Dim _length, _digitPosition As Integer

        _value = pAmount.ToString
        _length = Len(_value) ' ��Ҵ�ͧ �����ŷ���ͧ����ŧ �� 122200 �բ�Ҵ ��ҡѺ 6

        For i As Integer = 0 To _length - 1 ' ǹ�ٻ ������ҡ 0 ���֧ (��Ҵ - 1)
            ' ���˹觢ͧ ��ѡ (digit) �ͧ����Ţ
            ' ��
            ' ���˹���ѡ���0 (��ѡ˹���)
            ' ���˹���ѡ���1 (��ѡ�Ժ)
            ' ���˹���ѡ���2 (��ѡ����)
            ' ����繢����� i = 7 ���˹���ѡ����ҡѺ 1 (��ѡ�Ժ)
            ' ����繢����� i = 9 ���˹���ѡ����ҡѺ 3 (��ѡ�ѹ)
            ' ����繢����� i = 13 ���˹���ѡ����ҡѺ 1 (��ѡ�Ժ)
            _digitPosition = i - (6 * ((i - 1) \ 6))
            _aWord = Mid(_value, Len(_value) - i, 1)
            _text = ""
            Select Case _digitPosition
                Case 0 ' ��ѡ˹���
                    If _aWord = "1" And _length > 1 Then
                        ' ������Ţ 1 ����բ�Ҵ�ҡ���� 1 ����դ����ҡѺ "���"
                        _text = "���"
                    ElseIf _aWord <> "0" Then
                        ' ���������Ţ 0 ����� ����ѡ�� � _numberText()
                        _text = _numberText(CInt(_aWord))
                    End If
                Case 1 ' ��ѡ�Ժ
                    If _aWord = "1" Then
                        ' ������Ţ 1 ����ͧ�� ����ѡ�� ����ա �͡�ҡ����� "�Ժ"
                        '_numberTranslatedText = "�Ժ" + _numberTranslatedText
                        _text = _digit(_digitPosition)
                    ElseIf _aWord = "2" Then
                        ' ������Ţ 2 ������ѡ�ä�� "����Ժ"
                        _text = "���" + _digit(_digitPosition)
                    ElseIf _aWord <> "0" Then
                        ' ���������Ţ 0 ����� ����ѡ�� � _numberText() �������ѡ(digit) � _digit()
                        _text = _numberText(CInt(_aWord)) + _digit(_digitPosition)
                    End If
                Case 2, 3, 4, 5 ' ��ѡ���� �֧ �ʹ
                    If _aWord <> "0" Then
                        _text = _numberText(CInt(_aWord)) + _digit(_digitPosition)
                    End If
                Case 6 ' ��ѡ��ҹ
                    If _aWord = "0" Then
                        _text = "��ҹ"
                    ElseIf _aWord = "1" And _length - 1 > i Then
                        _text = "�����ҹ"
                    Else
                        _text = _numberText(CInt(_aWord)) + _digit(_digitPosition)
                    End If
            End Select
            _numberTranslatedText = _text + _numberTranslatedText
        Next

        Return _numberTranslatedText
    End Function
End Class
