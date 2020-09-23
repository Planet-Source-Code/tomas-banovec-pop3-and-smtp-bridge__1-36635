Attribute VB_Name = "modMain"

'thanx www.vbip.com for code of this encode function
 Function Base64_Encode(strSource) As String
    '
    Const BASE64_TABLE As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    '
    Dim strTempLine As String
    Dim j As Integer
    '
    For j = 1 To (Len(strSource) - Len(strSource) Mod 3) Step 3
        'Breake each 3 (8-bits) bytes to 4 (6-bits) bytes
        '
        '1 byte
        strTempLine = strTempLine + Mid(BASE64_TABLE, (Asc(Mid(strSource, j, 1)) \ 4) + 1, 1)
        '2 byte
        strTempLine = strTempLine + Mid(BASE64_TABLE, ((Asc(Mid(strSource, j, 1)) Mod 4) * 16 _
                       + Asc(Mid(strSource, j + 1, 1)) \ 16) + 1, 1)
        '3 byte
        strTempLine = strTempLine + Mid(BASE64_TABLE, ((Asc(Mid(strSource, j + 1, 1)) Mod 16) * 4 _
                       + Asc(Mid(strSource, j + 2, 1)) \ 64) + 1, 1)
        '4 byte
        strTempLine = strTempLine + Mid(BASE64_TABLE, (Asc(Mid(strSource, j + 2, 1)) Mod 64) + 1, 1)
    Next j
    '
    If Not (Len(strSource) Mod 3) = 0 Then
        '
        If (Len(strSource) Mod 3) = 2 Then
            '
            strTempLine = strTempLine + Mid(BASE64_TABLE, (Asc(Mid(strSource, j, 1)) \ 4) + 1, 1)
            '
            strTempLine = strTempLine + Mid(BASE64_TABLE, (Asc(Mid(strSource, j, 1)) Mod 4) * 16 _
                       + Asc(Mid(strSource, j + 1, 1)) \ 16 + 1, 1)
            '
            strTempLine = strTempLine + Mid(BASE64_TABLE, (Asc(Mid(strSource, j + 1, 1)) Mod 16) * 4 + 1, 1)
            '
            strTempLine = strTempLine & "="
            '
        ElseIf (Len(strSource) Mod 3) = 1 Then
            '
            '
            strTempLine = strTempLine + Mid(BASE64_TABLE, Asc(Mid(strSource, j, 1)) \ 4 + 1, 1)
            '
            strTempLine = strTempLine + Mid(BASE64_TABLE, (Asc(Mid(strSource, j, 1)) Mod 4) * 16 + 1, 1)
            '
            strTempLine = strTempLine & "=="
            '
        End If
        '
    End If
    '
    Base64_Encode = strTempLine
    '
End Function




