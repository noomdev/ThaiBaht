Public Class Sample

    Public Shared Function ThaiBaht(ByVal pAmount As Double) As String
        If pAmount = 0 Then
            Return "ศูนย์บาทถ้วน"
        End If

        Dim _integerValue As String ' จำนวนเต็ม
        Dim _decimalValue As String ' ทศนิยม
        Dim _integerTranslatedText As String ' จำนวนเต็ม ภาษาไทย
        Dim _decimalTranslatedText As String ' ทศนิยมภาษาไทย

        _integerValue = Format(pAmount, "####.00") ' จัด Format ค่าเงินเป็นตัวเลข 2 หลัก
        _decimalValue = Mid(_integerValue, Len(_integerValue) - 1, 2) ' ทศนิยม
        _integerValue = Mid(_integerValue, 1, Len(_integerValue) - 3) ' จำนวนเต็ม

        ' แปลง จำนวนเต็ม เป็น ภาษาไทย
        _integerTranslatedText = NumberToText(CDbl(_integerValue))

        ' แปลง ทศนิยม เป็น ภาษาไทย
        If CDbl(_decimalValue) = 0 Then
            _decimalTranslatedText = NumberToText(CDbl(_decimalValue))
        Else
            _decimalTranslatedText = ""
        End If

        ' ถ้าไม่มีทศนิม
        If _decimalTranslatedText.Trim = "" Then
            _integerTranslatedText += "บาทถ้วน"
        Else
            _integerTranslatedText += "บาท" & _decimalTranslatedText & "สตางค์"
        End If

        Return _integerTranslatedText
    End Function


    Private Shared Function NumberToText(ByVal pAmount As Double) As String
        ' ตัวอักษร
        Dim _numberText() As String = {"", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า", "สิบ"}

        ' หลัก หน่วย สิบ ร้อย พัน ...
        Dim _digit() As String = {"", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน", "ล้าน"}
        Dim _value As String, _aWord As String, _text As String
        Dim _numberTranslatedText As String = ""
        Dim _length, _digitPosition As Integer

        _value = pAmount.ToString
        _length = Len(_value) ' ขนาดของ ข้อมูลที่ต้องการแปลง เช่น 122200 มีขนาด เท่ากับ 6

        For i As Integer = 0 To _length - 1 ' วนลูป เริ่มจาก 0 จนถึง (ขนาด - 1)
            ' ตำแหน่งของ หลัก (digit) ของตัวเลข
            ' เช่น
            ' ตำแหน่งหลักที่0 (หลักหน่วย)
            ' ตำแหน่งหลักที่1 (หลักสิบ)
            ' ตำแหน่งหลักที่2 (หลักร้อย)
            ' ถ้าเป็นข้อมูล i = 7 ตำแหน่งหลักจะเท่ากับ 1 (หลักสิบ)
            ' ถ้าเป็นข้อมูล i = 9 ตำแหน่งหลักจะเท่ากับ 3 (หลักพัน)
            ' ถ้าเป็นข้อมูล i = 13 ตำแหน่งหลักจะเท่ากับ 1 (หลักสิบ)
            _digitPosition = i - (6 * ((i - 1) \ 6))
            _aWord = Mid(_value, Len(_value) - i, 1)
            _text = ""
            Select Case _digitPosition
                Case 0 ' หลักหน่วย
                    If _aWord = "1" And _length > 1 Then
                        ' ถ้าเป็นเลข 1 และมีขนาดมากกว่า 1 ให้มีค่าเท่ากับ "เอ็ด"
                        _text = "เอ็ด"
                    ElseIf _aWord <> "0" Then
                        ' ถ้าไม่ใช่เลข 0 ให้หา ตัวอักษร ใน _numberText()
                        _text = _numberText(CInt(_aWord))
                    End If
                Case 1 ' หลักสิบ
                    If _aWord = "1" Then
                        ' ถ้าเป็นเลข 1 ไม่ต้องมี ตัวอักษร อื่นอีก นอกจากคำว่า "สิบ"
                        '_numberTranslatedText = "สิบ" + _numberTranslatedText
                        _text = _digit(_digitPosition)
                    ElseIf _aWord = "2" Then
                        ' ถ้าเป็นเลข 2 ให้ตัวอักษรคือ "ยี่สิบ"
                        _text = "ยี่" + _digit(_digitPosition)
                    ElseIf _aWord <> "0" Then
                        ' ถ้าไม่ใช่เลข 0 ให้หา ตัวอักษร ใน _numberText() และหาหลัก(digit) ใน _digit()
                        _text = _numberText(CInt(_aWord)) + _digit(_digitPosition)
                    End If
                Case 2, 3, 4, 5 ' หลักร้อย ถึง แสน
                    If _aWord <> "0" Then
                        _text = _numberText(CInt(_aWord)) + _digit(_digitPosition)
                    End If
                Case 6 ' หลักล้าน
                    If _aWord = "0" Then
                        _text = "ล้าน"
                    ElseIf _aWord = "1" And _length - 1 > i Then
                        _text = "เอ็ดล้าน"
                    Else
                        _text = _numberText(CInt(_aWord)) + _digit(_digitPosition)
                    End If
            End Select
            _numberTranslatedText = _text + _numberTranslatedText
        Next

        Return _numberTranslatedText
    End Function
End Class
