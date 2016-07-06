VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5880
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim x As String

ChDir "D:\" 'Ваш путь
x = Application.GetSaveAsFilename("e", "PDF (*.pdf), *.pdf")
If x = "False" Then
'Действия если юзер нажал "Отмена"
Else
If OptionButton1 = True Then

ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=(x), _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End If
If OptionButton2 = True Then

 ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=(x), _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
End If

ThisWorkbook.Close False
End If
End Sub








