Attribute VB_Name = "libs_Dict2JsonFunc"
Function Dict2Json(ByVal dict As Object) As String
    Dim key As Variant, result As String, value As String

    result = "{"
    For Each key In dict.Keys
        result = result & IIf(Len(result) > 1, ",", "") & vbLf

        If TypeName(dict(key)) = "Dictionary" Then
            value = Dict2Json(dict(key))
            Dict2Json = value
        Else
            value = """" & dict(key) & """"
        End If

        result = result & "     " & """" & key & """:" & value & ""
    Next key
    result = result & vbLf & "}"

    Dict2Json = result
End Function
