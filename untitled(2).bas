$Console:Only
$Debug

Open "jsonparser.txt" For Input As #1
a$ = Input$(LOF(1), 1)
Close #1

Do
    Line Input ">"; c$
    Print qJSON(a$, c$)
    'Input ""; c
Loop

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


        'If DictiKey = Left$(upJSONstring$, InStr(upJSONstring$, Chr$(34)) - 1) Then



        'Else

        '    upJSONstring$ = Mid$(upJSONstring$, InStr(upJSONstring$, Chr$(34)) + 1)

        '    If InStr(upJSONstring$, ",") < InStr(upJSONstring$, Chr$(34)) Then

        '        upJSONstring$ = Mid$(upJSONstring$, InStr(upJSONstring$, ",") + 1)

        '    ElseIf Left$(upJSONstring$, 2) = ":" + Chr$(34) Then

        '        upJSONstring$ = Mid$(upJSONstring$, InStr(3, upJSONstring$, Chr$(34)))


        '    ElseIf Left$(upJSONstring$, 2) = ":{" Then

        '        i = 1
        '        Do
        '            i = InStr(i, upJSONstring$, "}")
        '            If i = 0 Then
        '                Print ("Err")
        '                Exit Do
        '            End If

        '            numOFopeningCurly = cntNstrinstr(upJSONstring$, "{", "", i)
        '            numOFclosingCurly = cntNstrinstr(upJSONstring$, "}", "", i) + 1

        '            If numOFopeningCurly = numOFclosingCurly Then
        '                Exit Do
        '            End If

        '        Loop

        '        upJSONstring$ = Mid$(upJSONstring$, i + 1)

        '    End If

        'End If
    Loop

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
