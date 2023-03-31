VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CombineWords 
   Caption         =   "RodolfoHelper v2.0 - Juntar Palavras"
   ClientHeight    =   6765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6810
   OleObjectBlob   =   "CombineWords.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CombineWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnBack_Click()
    txtFirst.Value = ""
    txtSecond.Value = ""
    txtRows.Value = ""
    txtFileName.Value = ""
    CombineWords.Hide
    RodolfoHelper.Show
End Sub

Private Sub btnRun_Click()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Selecionar pasta para guardar"
        If .Show = -1 Then
            If txtFirst.Value = "" Or txtSecond.Value = "" Or txtRows = "" Then
            MsgBox "ERRO: Campos não Preenchidos.", Title:="Erro"
        Else
    
        If txtFileName.Value = "" Then
            myFile = .SelectedItems(1) & "\noname.csv"
        Else
            myFile = .SelectedItems(1) & "\" & txtFileName.Value & ".csv"
        End If
  
        Dim r1 As String
        Dim r2 As String

        Dim cell_1 As Range
        Dim cell_2 As Range

        Dim result As Variant

        Set cell_1 = Range(txtSecond.Value)
        Set cell_2 = Range(txtFirst.Value)

        Open myFile For Output As #1
        For K = 1 To txtRows.Value
            r1 = cell_1(K)
            r2 = cell_2(K)
            result = r2 & r1
            Print #1, result
        Next K
        Close #1
        MsgBox "Ficheiro " & txtFileName.Value & ".csv guardado em: " & vbCrLf & myFile, Title:="Ficheiro Guardado com Sucesso"
            txtFirst.Value = ""
        txtSecond.Value = ""
        txtRows.Value = ""
        txtFileName.Value = ""
        CombineWords.Hide
        RodolfoHelper.Show
    End If
        Else
            MsgBox "Cancelado: Escolher pasta para guardar ficheiro.", Title:="Cancelado"
        End If
    End With
   
End Sub

Private Sub cbLocation_Change()
    If cbLocation.Value = "Desktop" Then
        lblLocation.Caption = CreateObject("WScript.Shell").specialfolders("Desktop")
    Else
        lblLocation.Caption = Application.DefaultFilePath
    End If
End Sub


Private Sub UserForm_Initialize()
    lblFileName = myFile
End Sub
