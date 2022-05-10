VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} gitCompiler_Creator 
   Caption         =   "gitCompiler"
   ClientHeight    =   5970
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7050
   OleObjectBlob   =   "gitCompiler_Creator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "gitCompiler_Creator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Created As Boolean
Private Sub btnCreate_Click()
    Set uForm_RetDict = New Dictionary
    If Me.txbProjectName.value = "" Then
        MsgBox "Nome do projeto não pode ser vazio"
        Exit Sub
    End If
    uForm_RetDict("projectName") = Me.txbProjectName.value
    uForm_RetDict("projectDesc") = Me.txbProjectDescription.value
    If Me.txbProjectPassword.value <> "" Then
        uForm_RetDict("projectPass") = EncryptSHA256(Me.txbProjectPassword.value, gitCompilerPrivateKey)
    Else
        uForm_RetDict("projectPass") = ""
    End If
    Created = True
    Unload gitCompiler_Creator
End Sub

Private Sub UserForm_Terminate()
    If Created = False Then Set uForm_RetDict = New Dictionary
End Sub
