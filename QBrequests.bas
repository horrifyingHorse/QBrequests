$Console:Only

Do
    Line Input ""; link$
    gda$ = qGet$(link$) '"https://www.google.com")
    Print (qStatusCode(gda$))
    ' Print gda$
    Print qContentLen(gda$)
    'Input c
    Print (qUrl$(gda$))
    'Print (qText(gda$))
    Print (qHeaders(gda$))
    Print (qEncoding(gda$))
Loop
End

Function qGet$ (url As String)
    'curl -vs "https://www.example.com" > output.txt 2>&1
    'Shell "echo no"

    Shell "cmd /c curl -vs " + Chr$(34) + url + Chr$(34) + " 2> req.get.tanu > req.get.con.tanu"

    Open "req.get.tanu" For Input As #1
    qGetResponseOutput$ = Input$(LOF(1), # 1)
    Close #1
    'Kill "req.get.tanu"

    qGetResponseOutput$ = qGetResponseOutput$ + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)

    Open "req.get.con.tanu" For Input As #1
    qGet$ = qGetResponseOutput$ + Input$(LOF(1), # 1)
    Close #1
    'Kill "req.get.con.tanu"
End Function

Function qContentLen (qObject As String)
    If InStr(UCase$(qObject), "CONTENT-LENGTH:") Then

        qObjectDataBytes$ = Mid$(qObject, InStr(UCase$(qObject), "CONTENT-LENGTH:") + 15)
        qObjectDataBytes$ = Left$(qObjectDataBytes$, InStr(qObjectDataBytes$, Chr$(13) + Chr$(10)) - 1)
        qContentLen = Val(qObjectDataBytes$)

        'ElseIf InStr(qObject, "bytes data]") Then
        '    qObjectDataBytes$ = Left$(qObject, InStr(qObject, "bytes data]") - 1)
        '    qObjectDataBytes$ = Mid$(qObjectDataBytes$, _InStrRev(qObjectDataBytes$, "[") + 1)
        '    qContentLen = Val(qObjectDataBytes$)

    Else

        qContentLen = 0

    End If

End Function

Function qText$ (qObject As String)
    If InStr(qObject, Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)) Then
        'qObjectDataBytes$ = Left$(qObject, InStr(qObject, "bytes data]") - 1)
        'qObjectDataBytes$ = Mid$(qObjectDataBytes$, _InStrRev(qObjectDataBytes$, "[") + 1)
        qText$ = Mid$(qObject, InStr(qObject, Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)) + 7) 'qContentLen(qobject)
        'qText$ = Mid$(qObjectDataBytes$, 0, qContentLen(qObject))

    Else

        qText$ = ""

    End If

End Function

Function qStatusCode (qObject As String)
    If InStr(UCase$(qObject), "< HTTP/1.1") Then

        qObjectgetStatusCode$ = Mid$(qObject, InStr(UCase$(qObject), "< HTTP/1.1") + 10)
        qStatusCode = Val(Left$(qObjectgetStatusCode$, InStr(qObjectgetStatusCode$, Chr$(13) + Chr$(10)) - 1))

    Else
        qStatusCode = 0

    End If

End Function

Function qUrl$ (qObject As String) 'www.example.com
    '
    If InStr(UCase$(qObject), "> GET") Then

        qObjectgetURL$ = Mid$(qObject, InStr(UCase$(qObject), "> GET") + 5)
        qRequestURLpath$ = _Trim$(Left$(qObjectgetURL$, InStr(qObjectgetURL$, "HTTP/1.1") - 1))
        'Print "p="; qRequestURLpath$
        'Print "-------------->"; qObjectgetURL$
        qObjectgetURL$ = Mid$(qObjectgetURL$, InStr(qObjectgetURL$, "> Host: ") + 8)
        'Print "-------------->"; qObjectgetURL$, InStr(qObjectgetURL$, Chr$(13) + Chr$(10))
        qObjectgetURL$ = _Trim$(Left$(qObjectgetURL$, InStr(qObjectgetURL$, Chr$(13)) - 1))
        'Print qObjectgetURL$, qRequestURLpath$, Asc(Right$(qObjectgetURL$, 1))
        qObjectgetURL$ = qObjectgetURL$ + qRequestURLpath$
        'Print (qObjectgetURL$)
        qUrl$ = qObjectgetURL$

    Else

        qUrl$ = "huh?"
        'Print (qObject)

    End If

End Function

Function qHeaders$ (qObject As String)
    If InStr(UCase$(qObject), "< HTTP/1.1") Then

        qObjectHeaderData$ = ""
        qObjectHeader$ = Mid$(qObject, InStr(UCase$(qObject), "< HTTP/1.1"))

        Do
            temporaryHeaderData$ = _Trim$(Mid$(qObjectHeader$, InStr(qObjectHeader$, "< ") + 1, InStr(qObjectHeader$, Chr$(13)) - 2))
            qObjectHeaderData$ = qObjectHeaderData$ + Chr$(13) + Chr$(10) + temporaryHeaderData$

            If temporaryHeaderData$ = Chr$(13) Then Exit Do
            If temporaryHeaderData$ = "" Then Exit Do

            qObjectHeader$ = Mid$(qObjectHeader$, InStr(qObjectHeader$, Chr$(13)) + 2)

        Loop

        qHeaders$ = qObjectHeaderData$

    Else

        qHeaders$ = "huh?"

    End If

End Function

Function qEncoding$ (qObject As String)
    If InStr(LCase$(qObject), "charset=") Then

        qEncoding$ = _Trim$(mid$(qobject, InStr(LCase$(qObject), "charset=") + 8, _
                        InStr(InStr(LCase$(qObject), "charset=") + 8, qObject, chr$(13)) - InStr(LCase$(qObject), "charset=") - 8))
        'Print (InStr(InStr(LCase$(qObject), "charset=") + 8, qObject, Chr$(13)) - InStr(LCase$(qObject), "charset=") - 8)

    Else

        qEncoding$ = ""

    End If
End Function

Function qGenURL$ ()
End Function



'$Console:Only
'Shell ("curl https://www.example.com/path/to/resource")

'client = _OpenClient("TCP/IP:443:www.w3schools.com")
'e$ = Chr$(13) + Chr$(10)
''https://www.w3schools.com/python/demopage.htm
'' Send an HTTPS request with the correct host and port
'x$ = "GET /python/demopage.htm HTTP/1.1" + e$
'x$ = x$ + "Host: www.w3schools.com" + e$
'x$ = x$ + "User-Agent: req" + e$
'x$ = x$ + "Accept: */*" + e$
'x$ = x$ + "Connection: keep-alive" + e$ + e$

'Print x$
'Put #client, , x$

'timelimit = 10

't! = Timer
'Do
'    _Delay 0.05 ' 50ms delay (20 checks per second)
'    Get #client, , a2$

'    If Len(a2$) Then
'        Print a2$;
'    End If

'Loop Until Timer > t! + timelimit

