Attribute VB_Name = "gitCompiler_gitConf"
Public gitCompiler_conf As Dictionary

Public Function gitCompiler_getConf() As Dictionary
    Dim Ret As Dictionary
    Set fso = New Scripting.FileSystemObject
    gitCompiler_Path = ActiveWorkbook.Path & "\"
    Set Ret = Json2Dict(fso.OpenTextFile(gitCompiler_Path & ".conf", ForReading).ReadAll)
    
    For Each Keys In Ret.Keys
        prevVal = Ret(Keys)
        Ret.Remove Keys
        Ret(Replace(Keys, "obj.", "")) = prevVal
    Next
    Set gitCompiler_getConf = Ret
End Function

Public Function gitCompiler_setConf(ByVal dict As Dictionary)
    Set fso = New Scripting.FileSystemObject
    gitCompiler_Path = ActiveWorkbook.Path & "\"
    fso.OpenTextFile(gitCompiler_Path & ".conf", ForWriting).Write Dict2Json(dict)
End Function

