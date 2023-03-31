VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WordExtract 
   Caption         =   "RodolfoHelper v2.0 - Extrair Palavras"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6570
   OleObjectBlob   =   "WordExtract.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "WordExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub btnBack_Click()
    txtStart.Value = ""
    txtRows.Value = ""
    txtPos.Value = ""
    txtSize.Value = ""
    txtFileName.Value = ""
    wordExtract.Hide
    RodolfoHelper.Show
End Sub

Private Sub btnRun_Click()
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Selecionar pasta para guardar"
        If .Show = -1 Then
            If txtStart.Value = "" Or txtRows.Value = "" Or txtPos.Value = "" Or txtSize = "" Then
            MsgBox "ERRO: Campos não Preenchidos.", Title:="Erro"
        Else
    
       If txtFileName.Value = "" Then
            myFile = .SelectedItems(1) & "\noname.csv"
        Else
            myFile = .SelectedItems(1) & "\" & txtFileName.Value & ".csv"
        End If
    
        Dim cell_1 As Range
        Dim value_1 As Variant
        Set cell_1 = Range(txtStart.Value)
        
    
        Open myFile For Output As #1
        For K = 1 To txtRows.Value
            value_1 = Left(Mid(cell_1(K), txtPos.Value), txtSize.Value)
            Print #1, value_1
        Next K
        Close #1
        MsgBox "Ficheiro " & txtFileName.Value & ".csv guardado em: " & vbCrLf & myFile, Title:="Ficheiro Guardado com Sucesso"
        txtStart.Value = ""
        txtRows.Value = ""
        txtPos.Value = ""
        txtSize.Value = ""
        txtFileName.Value = ""
        wordExtract.Hide
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
