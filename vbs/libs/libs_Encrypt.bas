Attribute VB_Name = "libs_Encrypt"
Private Const clOneMask = 16515072          '000000 111111 111111 111111
Private Const clTwoMask = 258048            '111111 000000 111111 111111
Private Const clThreeMask = 4032            '111111 111111 000000 111111
Private Const clFourMask = 63               '111111 111111 111111 000000

Private Const clHighMask = 16711680         '11111111 00000000 00000000
Private Const clMidMask = 65280             '00000000 11111111 00000000
Private Const clLowMask = 255               '00000000 00000000 11111111

Private Const cl2Exp18 = 262144             '2 to the 18th power
Private Const cl2Exp12 = 4096               '2 to the 12th
Private Const cl2Exp6 = 64                  '2 to the 6th
Private Const cl2Exp8 = 256                 '2 to the 8th
Private Const cl2Exp16 = 65536              '2 to the 16th



Public Function EncryptSHA256(ByVal sTextToHash As String, ByVal sSharedSecretKey As String)
    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA256")

    TextToHash = asc.Getbytes_4(sTextToHash)
    SharedSecretKey = asc.Getbytes_4(sSharedSecretKey)
    enc.key = SharedSecretKey

    Dim bytes() As Byte
    bytes = enc.ComputeHash_2((TextToHash))
    EncryptSHA256 = EncodeBase64(bytes)
    Set asc = Nothing
    Set enc = Nothing
End Function

Private Function EncodeBase64(ByRef arrData() As Byte) As String
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement

    Set objXML = New MSXML2.DOMDocument60

    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.Text

    Set objNode = Nothing
    Set objXML = Nothing

End Function


'mdlforencryptionanddecryption
Public Function xDecrypt(CodeKey As String, DataIn As String) As String
Dim arkdata1 As Long
Dim strDataOut As String
Dim intXOrValue1 As Integer
Dim intXOrValue2 As Integer
For arkdata1 = 1 To (Len(DataIn) / 2)

intXOrValue1 = Val("&H" & (Mid$(DataIn, (2 * arkdata1) - 1, 2)))

intXOrValue2 = asc(Mid$(CodeKey, ((arkdata1 Mod Len(CodeKey)) + 1), 1))
strDataOut = strDataOut + Chr(intXOrValue1 Xor intXOrValue2)
Next arkdata1
xDecrypt = strDataOut
End Function

Public Function xEncrypt(CodeKey As String, DataIn As String) As String
Dim arkdata1 As Long
Dim strDataOut As String
Dim temp As Integer
Dim tempstring As String
Dim intXOrValue1 As Integer
Dim intXOrValue2 As Integer
For arkdata1 = 1 To Len(DataIn)

intXOrValue1 = asc(Mid$(DataIn, arkdata1, 1))

intXOrValue2 = asc(Mid$(CodeKey, ((arkdata1 Mod Len(CodeKey)) + 1), 1))
temp = (intXOrValue1 Xor intXOrValue2)
tempstring = Hex(temp)
If Len(tempstring) = 1 Then tempstring = "0" & tempstring
strDataOut = strDataOut + tempstring
Next arkdata1
xEncrypt = strDataOut
End Function



Public Function Encode64(ByVal sString As String) As String
    If sString = "" Then
        Encode64 = ""
        Exit Function
    End If
    Dim bTrans(63) As Byte, lPowers8(255) As Long, lPowers16(255) As Long, bOut() As Byte, bIn() As Byte
    Dim lChar As Long, lTrip As Long, iPad As Integer, lLen As Long, lTemp As Long, lPos As Long, lOutSize As Long

    For lTemp = 0 To 63                                 'Fill the translation table.
        Select Case lTemp
            Case 0 To 25
                bTrans(lTemp) = 65 + lTemp              'A - Z
            Case 26 To 51
                bTrans(lTemp) = 71 + lTemp              'a - z
            Case 52 To 61
                bTrans(lTemp) = lTemp - 4               '1 - 0
            Case 62
                bTrans(lTemp) = 43                      'Chr(43) = "+"
            Case 63
                bTrans(lTemp) = 47                      'Chr(47) = "/"
        End Select
    Next lTemp

    For lTemp = 0 To 255                                'Fill the 2^8 and 2^16 lookup tables.
        lPowers8(lTemp) = lTemp * cl2Exp8
        lPowers16(lTemp) = lTemp * cl2Exp16
    Next lTemp

    iPad = Len(sString) Mod 3                           'See if the length is divisible by 3
    If iPad Then                                        'If not, figure out the end pad and resize the input.
        iPad = 3 - iPad
        sString = sString & String(iPad, Chr(0))
    End If

    bIn = StrConv(sString, vbFromUnicode)               'Load the input string.
    lLen = ((UBound(bIn) + 1) \ 3) * 4                  'Length of resulting string.
    lTemp = lLen \ 72                                   'Added space for vbCrLfs.
    lOutSize = ((lTemp * 2) + lLen) - 1                 'Calculate the size of the output buffer.
    ReDim bOut(lOutSize)                                'Make the output buffer.

    lLen = 0                                            'Reusing this one, so reset it.

    For lChar = LBound(bIn) To UBound(bIn) Step 3
        lTrip = lPowers16(bIn(lChar)) + lPowers8(bIn(lChar + 1)) + bIn(lChar + 2)    'Combine the 3 bytes
        lTemp = lTrip And clOneMask                     'Mask for the first 6 bits
        bOut(lPos) = bTrans(lTemp \ cl2Exp18)           'Shift it down to the low 6 bits and get the value
        lTemp = lTrip And clTwoMask                     'Mask for the second set.
        bOut(lPos + 1) = bTrans(lTemp \ cl2Exp12)       'Shift it down and translate.
        lTemp = lTrip And clThreeMask                   'Mask for the third set.
        bOut(lPos + 2) = bTrans(lTemp \ cl2Exp6)        'Shift it down and translate.
        bOut(lPos + 3) = bTrans(lTrip And clFourMask)   'Mask for the low set.
        If lLen = 68 Then                               'Ready for a newline
            bOut(lPos + 4) = 13                         'Chr(13) = vbCr
            bOut(lPos + 5) = 10                         'Chr(10) = vbLf
            lLen = 0                                    'Reset the counter
            lPos = lPos + 6
        Else
            lLen = lLen + 4
            lPos = lPos + 4
        End If
    Next lChar

    If bOut(lOutSize) = 10 Then lOutSize = lOutSize - 2 'Shift the padding chars down if it ends with CrLf.

    If iPad = 1 Then                                    'Add the padding chars if any.
        bOut(lOutSize) = 61                             'Chr(61) = "="
    ElseIf iPad = 2 Then
        bOut(lOutSize) = 61
        bOut(lOutSize - 1) = 61
    End If

    Encode64 = StrConv(bOut, vbUnicode)                 'Convert back to a string and return it.

End Function

Public Function Decode64(ByVal sString As String) As String
    If sString = "" Then
        Decode64 = ""
        Exit Function
    End If
    Dim bOut() As Byte, bIn() As Byte, bTrans(255) As Byte, lPowers6(63) As Long, lPowers12(63) As Long
    Dim lPowers18(63) As Long, lQuad As Long, iPad As Integer, lChar As Long, lPos As Long, sOut As String
    Dim lTemp As Long

    sString = Replace(sString, vbCr, vbNullString)      'Get rid of the vbCrLfs.  These could be in...
    sString = Replace(sString, vbLf, vbNullString)      'either order.

    lTemp = Len(sString) Mod 4                          'Test for valid input.
    If lTemp Then
        Call Err.Raise(vbObjectError, "MyDecode", "Input string is not valid Base64.")
    End If

    If InStrRev(sString, "==") Then                     'InStrRev is faster when you know it's at the end.
        iPad = 2                                        'Note:  These translate to 0, so you can leave them...
    ElseIf InStrRev(sString, "=") Then                  'in the string and just resize the output.
        iPad = 1
    End If

    For lTemp = 0 To 255                                'Fill the translation table.
        Select Case lTemp
            Case 65 To 90
                bTrans(lTemp) = lTemp - 65              'A - Z
            Case 97 To 122
                bTrans(lTemp) = lTemp - 71              'a - z
            Case 48 To 57
                bTrans(lTemp) = lTemp + 4               '1 - 0
            Case 43
                bTrans(lTemp) = 62                      'Chr(43) = "+"
            Case 47
                bTrans(lTemp) = 63                      'Chr(47) = "/"
        End Select
    Next lTemp

    For lTemp = 0 To 63                                 'Fill the 2^6, 2^12, and 2^18 lookup tables.
        lPowers6(lTemp) = lTemp * cl2Exp6
        lPowers12(lTemp) = lTemp * cl2Exp12
        lPowers18(lTemp) = lTemp * cl2Exp18
    Next lTemp

    bIn = StrConv(sString, vbFromUnicode)               'Load the input byte array.
    ReDim bOut((((UBound(bIn) + 1) \ 4) * 3) - 1)       'Prepare the output buffer.

    For lChar = 0 To UBound(bIn) Step 4
        lQuad = lPowers18(bTrans(bIn(lChar))) + lPowers12(bTrans(bIn(lChar + 1))) + _
                lPowers6(bTrans(bIn(lChar + 2))) + bTrans(bIn(lChar + 3))           'Rebuild the bits.
        lTemp = lQuad And clHighMask                    'Mask for the first byte
        bOut(lPos) = lTemp \ cl2Exp16                   'Shift it down
        lTemp = lQuad And clMidMask                     'Mask for the second byte
        bOut(lPos + 1) = lTemp \ cl2Exp8                'Shift it down
        bOut(lPos + 2) = lQuad And clLowMask            'Mask for the third byte
        lPos = lPos + 3
    Next lChar

    sOut = StrConv(bOut, vbUnicode)                     'Convert back to a string.
    If iPad Then sOut = Left$(sOut, Len(sOut) - iPad)   'Chop off any extra bytes.
    Decode64 = sOut

End Function

