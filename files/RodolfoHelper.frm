VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RodolfoHelper 
   Caption         =   "RodolfoHelper v2.0"
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6960
   OleObjectBlob   =   "RodolfoHelper.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RodolfoHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnChooseFile_Click()
    Dim checkIfOpen
    myFile = Application.GetOpenFilename()
    lblUrl.Caption = myFile

    checkIfOpen = IsWorkBookOpen(myFile)

    If checkIfOpen = False Then
        Workbooks.Open myFile
    End If
    
    If myFile = "" Then
        lblUrl.Caption = "Nenhum Ficheiro Selecionado."
    Else
        lblUrl.Caption = myFile
        btnUnite.Visible = True
        btnExtract.Visible = True
    End If
End Sub

Private Sub btnExtract_Click()
    RodolfoHelper.Hide
    wordExtract.Show
End Sub


Private Sub btnUnite_Click()
    RodolfoHelper.Hide
    CombineWords.Show
End Sub


Private Sub UserForm_Initialize()
    lblWelcome.Caption = "Bem-vindo " & Application.UserName
    If myFile = "" Then
        lblUrl.Caption = "Nenhum Ficheiro Selecionado."
        btnUnite.Visible = False
        btnExtract.Visible = False
    Else
        lblUrl.Caption = myFile
        btnUnite.Visible = True
        btnExtract.Visible = True
    End If
End Sub


Function IsWorkBookOpen(FileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function
