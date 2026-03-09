Attribute VB_Name = "iFunction"
Option Compare Text
Option Explicit

Public Function DativeCase(ByVal sSurName As String, Optional ByVal sName As String = vbNullString, Optional ByVal sPatronymic As String = vbNullString) As String
    Dim parts As Variant
    Dim surnameParts As Variant
    Dim surnamePart As String
    Dim resultPart As String
    Dim resultValue As String
    Dim nameException As String
    Dim isMale As Boolean
    Dim index As Long

    Application.Volatile True

    sSurName = Replace$(sSurName, " - ", "-")
    sSurName = Replace$(Replace$(sSurName, " -", "-"), "- ", "-")

    If Len(sName) = 0 And Len(sPatronymic) = 0 Then
        parts = Split(Application.Trim(sSurName))

        If UBound(parts) >= 0 Then sSurName = parts(0)
        If UBound(parts) >= 1 Then sName = parts(1)
        If UBound(parts) >= 2 Then sPatronymic = Replace$(parts(2), ".", vbNullString)
    End If

    isMale = Not (Right$(sPatronymic, 2) = "на" Or Right$(sPatronymic, 4) = "кызы")

    If Len(sSurName) > 0 Then
        surnameParts = Split(sSurName, "-")

        For index = LBound(surnameParts) To UBound(surnameParts)
            surnamePart = CStr(surnameParts(index))
            resultPart = vbNullString

            If isMale Then
                Select Case Right$(surnamePart, 1)
                    Case "о", "и", "ы", "у", "э", "е", "ю"
                        resultPart = surnamePart
                    Case "ь", "й"
                        resultPart = Left$(surnamePart, Len(surnamePart) - 1) & "ю"
                    Case "€", "а"
                        resultPart = Left$(surnamePart, Len(surnamePart) - 1) & "е"
                        If UBound(surnameParts) > 0 And index = 0 Then resultPart = surnamePart
                    Case Else
                        resultPart = surnamePart & "у"
                End Select

                Select Case Right$(surnamePart, 2)
                    Case "ец"
                        resultPart = Left$(surnamePart, Len(surnamePart) - 2) & "цу"
                        If LCase$(surnamePart) Like "*[уеыаоэ€июЄ]ец" Then resultPart = Left$(surnamePart, Len(surnamePart) - 1) & "цу"
                        If LCase$(surnamePart) Like "*[!уеыаоэ€июЄ][!уеыаоэ€июЄ]ец" Then resultPart = surnamePart & "у"
                    Case "зе", "их", "ых"
                        resultPart = surnamePart
                    Case "ый"
                        resultPart = Left$(surnamePart, Len(surnamePart) - 2) & "ому"
                    Case "ий", "ой"
                        resultPart = Left$(surnamePart, Len(surnamePart) - 2) & "ому"
                        If Len(surnamePart) <= 4 Then resultPart = Left$(surnamePart, Len(surnamePart) - 1) & "ю"
                        If Right$(surnamePart, 3) = "чий" Then resultPart = Left$(surnamePart, Len(surnamePart) - 2) & "ему"
                    Case "уй"
                        resultPart = Left$(surnamePart, Len(surnamePart) - 2) & "ую"
                End Select
            Else
                Select Case Right$(surnamePart, 1)
                    Case "о", "е", "э", "и", "ы", "у", "ю", "б", "в", "г", "д", "ж", "з", "к", "л", "м", "н", "п", _
                         "р", "с", "т", "ф", "х", "ц", "ч", "ш", "щ", "ь", "й"
                        resultPart = surnamePart
                    Case "€"
                        resultPart = Left$(surnamePart, Len(surnamePart) - 2) & "ой"
                    Case Else
                        resultPart = Left$(surnamePart, Len(surnamePart) - 1) & "ой"
                End Select

                Select Case Right$(surnamePart, 2)
                    Case "ха", "ла", "ее"
                        resultPart = Left$(surnamePart, Len(surnamePart) - 1) & "е"
                End Select
            End If

            If LCase$(surnamePart) Like "*[уеыаоэ€июЄ]а" Then resultPart = surnamePart

            surnameParts(index) = resultPart
        Next index

        resultValue = Join(surnameParts, "-") & " "
    End If

    If Len(sName) > 0 Then
        nameException = GetDativeException(sName)

        If Len(nameException) > 0 Then
            resultValue = resultValue & nameException
        ElseIf isMale Then
            Select Case Right$(sName, 1)
                Case "й", "ь"
                    resultValue = resultValue & Left$(sName, Len(sName) - 1) & "ю"
                Case "€", "а"
                    resultValue = resultValue & Left$(sName, Len(sName) - 1) & "е"
                Case "о"
                    resultValue = resultValue & sName
                Case Else
                    resultValue = resultValue & sName & "у"
            End Select
        Else
            Select Case Right$(sName, 1)
                Case "а", "€"
                    If Mid$(sName, Len(sName) - 1, 1) = "и" Then
                        resultValue = resultValue & Left$(sName, Len(sName) - 1) & "и"
                    Else
                        resultValue = resultValue & Left$(sName, Len(sName) - 1) & "е"
                    End If
                Case "ь"
                    resultValue = resultValue & Left$(sName, Len(sName) - 1) & "и"
                Case Else
                    resultValue = resultValue & sName
            End Select
        End If

        resultValue = resultValue & " "
    End If

    If Len(sPatronymic) > 0 Then
        If Right$(sPatronymic, 4) = "оглы" Or Right$(sPatronymic, 4) = "кызы" Then
            resultValue = resultValue & sPatronymic
        ElseIf isMale Then
            resultValue = resultValue & sPatronymic & "у"
        Else
            resultValue = resultValue & Left$(sPatronymic, Len(sPatronymic) - 1) & "е"
        End If
    End If

    resultValue = Replace$(resultValue, "-", "- ")
    resultValue = StrConv(resultValue, vbProperCase)
    DativeCase = Replace$(resultValue, "- ", "-")
End Function

Public Function GetDativeException(ByVal txt As String) As String
    Select Case txt
        Case "ѕавел"
            GetDativeException = "ѕавлу"
        Case "Ћев"
            GetDativeException = "Ћьву"
        Case "ѕЄтр"
            GetDativeException = "ѕетру"
        Case "јли", "Ѕали"
            GetDativeException = txt
    End Select
End Function

Public Function ExtractNumbers(ByVal inputString As String) As String
    Dim outputString As String
    Dim index As Long

    For index = 1 To Len(inputString)
        If IsNumeric(Mid$(inputString, index, 1)) Then
            outputString = outputString & Mid$(inputString, index, 1)
        End If
    Next index

    ExtractNumbers = outputString
End Function
