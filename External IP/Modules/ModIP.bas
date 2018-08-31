Attribute VB_Name = "ModIP"
'Get External IP Address
'Author: Danny Elkins/DigiRev
'Created: May 19th, 2007
'Last updated: May 19th, 2007

Option Explicit

'ModIP.bas
'Simply add this module to your project and you're ready to go!

'Check if a given string is an IP address.
Private Function IsIP(ByRef Text As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim strBuff() As String, intLoop As Integer
    Dim bytOctect As Byte
    
    strBuff() = Split(Text, ".")
    
    If UBound(strBuff()) = 3 Then
        bytOctect = CByte(strBuff(0))
        bytOctect = CByte(strBuff(1))
        bytOctect = CByte(strBuff(2))
        bytOctect = CByte(strBuff(3))
        
        IsIP = True
    End If
    
    Erase strBuff()
    
    Exit Function
    
ErrorHandler:
    
End Function

'Find next non-numeric character going backwords.
Private Function NextNonNumBack(ByRef Text As String, ByVal Start As Long) As Long
    Dim lonLoop As Long, strCur As String
    
    For lonLoop = Start To 1 Step -1
        strCur = Mid$(Text, lonLoop, 1)
        If Not IsNumeric(strCur) And Not strCur = "." Then
            NextNonNumBack = lonLoop
            Exit For
        End If
    Next lonLoop
    
End Function

'Find next non-numeric character going forwards (excludes periods (.)).
Private Function NextNonNumFor(ByRef Text As String, ByVal Start As Long) As Long
    Dim lonLoop As Long, lonLen As Long
    Dim strCur As String
    
    lonLen = Len(Text)
    
    For lonLoop = Start To lonLen
        strCur = Mid$(Text, lonLoop, 1)
        
        If Not IsNumeric(strCur) And Not strCur = "." Then
            NextNonNumFor = lonLoop
            Exit For
        End If
    Next lonLoop
    
End Function

Public Function ExtractIPs(ByRef HTML As String, _
    Optional ByVal ReturnAll As Boolean = False) As Collection
    
    'Tested for errors before this.
    'This is just used to prevent duplicates.
    On Error GoTo ErrorHandler
    
    Dim lonLoop As Long, lonLen As Long
    Dim lonStart As Long, lonEnd As Long
    Dim strTemp As String, lonIPStart As Long
    Dim colRet As Collection
    
    'Find .
    lonStart = InStr(1, HTML, ".")
    Set colRet = New Collection
    
    Do While lonStart > 0
        lonIPStart = NextNonNumBack(HTML, lonStart)
        
        If lonIPStart > 0 Then
            lonIPStart = lonIPStart + 1
            '123.123.123.123
            lonEnd = NextNonNumFor(HTML, lonStart)
            
            If lonEnd > 0 Then
                strTemp = Mid$(HTML, lonIPStart, lonEnd - lonIPStart)
                
                If Len(strTemp) > 0 Then
                    If IsIP(strTemp) Then
                        colRet.Add strTemp, "IP:" & strTemp
                    
                        If Not ReturnAll Then Exit Do
                    End If
                End If
                
            Else
                Exit Do
            End If
        
        Else
            Exit Do
        End If
        
        lonStart = InStr(lonEnd, HTML, ".")
    Loop
    
    Set ExtractIPs = colRet
    Set colRet = Nothing
    
    Exit Function
    
ErrorHandler:
    Resume Next
End Function
