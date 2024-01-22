$Console:Only
$Debug

Open "jsonparser.txt" For Input As #1
a$ = Input$(LOF(1), 1)
Close #1


Do
    Line Input "K>"; keyString$
    Line Input "V>"; ValString$
    c$ = qJSONc(keyString$, ValString$)
    Print qJSONc("thisisKEY1", c$)
    'Input ""; c
Loop

' q


Function qJSON$ (DictiJSONstring As String, DictiKey As String)
    chr13Value = 1

    For i = 1 To Len(DictiJSONstring)

        If Mid$(DictiJSONstring, i, 1) = Chr$(34) Then

            If chr13Value Then
                chr13Value = 0
            Else
                chr13Value = 1
            End If

        End If

        If chr13Value Then
            If Mid$(DictiJSONstring, i, 1) = Chr$(32) Then _Continue
            If Mid$(DictiJSONstring, i, 1) = Chr$(13) Then _Continue
            If Mid$(DictiJSONstring, i, 1) = Chr$(10) Then _Continue
            If Mid$(DictiJSONstring, i, 1) = Chr$(9) Then _Continue
        End If

        upJSONstring$ = upJSONstring$ + Mid$(DictiJSONstring, i, 1)

    Next i

    colonNcount = 0: lastColonPosi = 0
    If InStr(DictiKey, ":") Then
        originalDictiKey$ = DictiKey
        Do
            colonNcount = colonNcount + 1
            DictiKey = Mid$(originalDictiKey$, lastColonPosi, InStr(originalDictiKey$, ":"))
            GoSub FindingKey
            lastColonPosi = InStr(originalDictiKey$, ":") + 1
        Loop While InStr(lastColonPosi, originalDictiKey$, ":")

        DictiKey = Mid$(originalDictiKey$, lastColonPosi)
        colonNcount = 0
        'For i = 0 To colonNcount
        '    dictkey = mid$(
        '    GoSub FindingKey

        'Next i

    End If

    FindingKey:
    Do

        upJSONstring$ = Mid$(upJSONstring$, InStr(upJSONstring$, Chr$(34)) + 1)
        currentExistingDictiKey$ = Left$(upJSONstring$, InStr(upJSONstring$, Chr$(34)) - 1)

        upJSONstring$ = Mid$(upJSONstring$, InStr(upJSONstring$, Chr$(34)) + 1)

        If InStr(upJSONstring$, ",") < InStr(upJSONstring$, Chr$(34)) And InStr(upJSONstring$, ",") Then

            'upJSONstring$ = Mid$(upJSONstring$, InStr(upJSONstring$, ",") + 1)
            upJSONstring = InStr(upJSONstring$, ",")
            upJSONstring1 = 2

        ElseIf Left$(upJSONstring$, 2) = ":" + Chr$(34) Then

            'upJSONstring$ = Mid$(upJSONstring$, InStr(3, upJSONstring$, Chr$(34)))
            upJSONstring = InStr(3, upJSONstring$, Chr$(34))
            upJSONstring1 = 3


        ElseIf Left$(upJSONstring$, 2) = ":{" Then

            upJSONstring = 1
            upJSONstring1 = 2
            Do
                upJSONstring = InStr(upJSONstring, upJSONstring$, "}")
                If upJSONstring = 0 Then
                    Print ("Err")
                    Exit Do
                End If

                numOFopeningCurly = cntNstrinstr(upJSONstring$, "{", "", upJSONstring)
                numOFclosingCurly = cntNstrinstr(upJSONstring$, "}", "", upJSONstring) + 1

                If numOFopeningCurly = numOFclosingCurly Then
                    upJSONstring = upJSONstring + 1
                    Exit Do
                End If


            Loop

            ' upJSONstring$ = Mid$(upJSONstring$, upJSONstring + 1)
        ElseIf Left$(upJSONstring$, 1) = ":" And InStr(upJSONstring$, "}") Then

            upJSONstring = InStr(upJSONstring$, "}")
            upJSONstring1 = 2

        End If

        If currentExistingDictiKey$ = DictiKey Then
            upJSONstring$ = Mid$(upJSONstring$, upJSONstring1, upJSONstring - upJSONstring1)
            Exit Do
        Else
            upJSONstring$ = Mid$(upJSONstring$, upJSONstring + 1)
        End If

        If upJSONstring$ = "" Then Exit Do

    Loop

    If colonNcount Then Return

    qJSON = upJSONstring$
End Function

Function cntNstrinstr (cntNstrinstr1 As String, cntNstrinstr2 As String, cntNstrinstr3 As String, cntNstrpos)
    cntNstrinstr4 = 0: cntNstrinstrPos = 0

    If cntNstrpos Then

        Do While (InStr(cntNstrinstrPos, cntNstrinstr1, cntNstrinstr2) < cntNstrpos And InStr(cntNstrinstrPos, cntNstrinstr1, cntNstrinstr2))
            cntNstrinstr4 = cntNstrinstr4 + 1
            cntNstrinstrPos = InStr(cntNstrinstrPos, cntNstrinstr1, cntNstrinstr2) + 1
        Loop

    Else
        Do While (InStr(cntNstrinstrPos, cntNstrinstr1, cntNstrinstr2) < InStr(cntNstrinstrPos, cntNstrinstr1, cntNstrinstr3) And InStr(cntNstrinstrPos, cntNstrinstr1, cntNstrinstr2))
            cntNstrinstr4 = cntNstrinstr4 + 1
            cntNstrinstrPos = InStr(cntNstrinstrPos, cntNstrinstr1, cntNstrinstr2) + 1
        Loop


    End If

    cntNstrinstr = cntNstrinstr4
End Function

Function qJSONc$ (KeyString As String, ValueString As String)
    queryString$ = KeyString: GoSub count
    keyStringColon = count
    queryString$ = ValueString: GoSub count

    If keyStringColon <> count Then
        Print "err"
        Exit Function
    End If

    cr8JSON$ = "{": lKColonNpos = 0: lVColonNpos = 0
    For i = 0 To count
        __AddedQuote$ = Chr$(34)

        If InStr(lKColonNpos, KeyString, ":") Then
            ithKey$ = _Trim$(Mid$(KeyString, lKColonNpos, InStr(lKColonNpos, KeyString, ":") - lKColonNpos))
            ithVal$ = _Trim$(Mid$(ValueString, lVColonNpos, InStr(lVColonNpos, ValueString, ":") - lVColonNpos))


            lKColonNpos = InStr(lKColonNpos, KeyString, ":") + 1
            lVColonNpos = InStr(lVColonNpos, ValueString, ":") + 1

            __AddedComma$ = ", "
        Else

            ithKey$ = _Trim$(Mid$(KeyString, lKColonNpos))
            ithVal$ = _Trim$(Mid$(ValueString, lVColonNpos))
            __AddedComma$ = "}"
        End If

        If Left$(ithVal$, 1) = "{" Then __AddedQuote$ = ""


        cr8JSON$ = cr8JSON$ + Chr$(34) + ithKey$ + Chr$(34) + ": " + __AddedQuote$ + ithVal$ + __AddedQuote$ + __AddedComma$
    Next i

    qJSONc = cr8JSON$

    Exit Function

    count:
    count = 0: lColonNpos = 0

    Do While InStr(lColonNpos, queryString$, ":")
        lColonNpos = InStr(lColonNpos, queryString$, ":") + 1
        numOFopeningCurly = cntNstrinstr(queryString$, "{", "", lColonNpos)
        numOFclosingCurly = cntNstrinstr(queryString$, "}", "", lColonNpos)

        If numOFopeningCurly = numOFclosingCurly Then
            count = count + 1
        End If
    Loop
    Return
End Function
