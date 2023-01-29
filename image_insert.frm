VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} image_insert 
   Caption         =   "画像挿入パレット"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4800
   OleObjectBlob   =   "image_insert.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "image_insert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const folder_directory As String = "C:\temp\insert_pic"

Private Sub CB_folder_Change()
    Dim buf As String, cnt As Long
    Const Path As String = folder_directory + "\"
    buf = Dir(Path & CB_folder.Value & "\" & "*.jpg")
    
    LB_image.Clear
    
    Do While buf <> ""
        LB_image.AddItem (buf)
        buf = Dir()
    Loop
    

End Sub

Private Sub CB_insert_Click()
    If Not LB_image.Value = "" Then
        Dim tmp_s
        tmp_s = folder_directory & "\" & CB_folder & "\" & LB_image.Value
        ActiveSheet.Pictures.Insert (tmp_s)
    End If
    

End Sub

Private Sub LB_image_Click()
    Dim tmp_s
    tmp_s = folder_directory & "\" & CB_folder & "\" & LB_image.Value
    image_preview.Picture = LoadPicture(tmp_s)
End Sub

Private Sub UserForm_Initialize()
    
    'Conbo box initialize
    Dim FSO As Object
    Dim TARGET As Object
    Dim folder_name As Variant
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set TARGET = FSO.GetFolder(folder_directory).SubFolders
    
    Dim tmp_str() As String
    
    
    For Each folder_name In TARGET
        tmp_str = Split(folder_name, "\")
        CB_folder.AddItem (tmp_str(UBound(tmp_str)))
    Next folder_name

End Sub
