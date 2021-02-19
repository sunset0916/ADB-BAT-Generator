Attribute VB_Name = "ADB_BAT_Generator"
Option Explicit

Sub batoutput()
    
    Dim fno As Integer
    Dim r As Long
    Dim i As Integer
    Dim sfile As String
    Dim u As String
    Dim e As String
    
    sfile = Application.GetSaveAsFilename(fileFilter:="バッチファイル(*.bat),*bat")
    If sfile = "False" Then
        Exit Sub
    End If
    
    fno = FreeFile
    
    r = Range("a1").CurrentRegion.Rows.Count
    
    e = "@echo off"
    
    Open sfile For Output As #fno
    
    Print #fno, e
    
    For i = 1 To r
        u = ""
        If Cells(i, 1) = "0" Then
            u = "adb shell pm uninstall -k --user 0 "
        ElseIf Cells(i, 1) = "1" Then
            u = "adb uninstall "
        Else
            MsgBox (i & "行目でエラーが発生しました。修正してからやり直してください。")
            Close #fno
            Exit Sub
        End If
        Print #fno, u & Cells(i, 2)
    Next
    
    Close #fno
    
End Sub
