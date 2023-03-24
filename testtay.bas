'Lam test tay cho themis 
'Written by Tran Huu Nam - huunam0@gmail.com
Attribute VB_Name = "testtay"
Option Explicit

Sub ghitep(Nd As String, tep As String)
    Open tep For Output As #1
    Print #1, Nd
    Close 1
End Sub
Function xoadb(s As String) As String
    Dim p As String, i As Integer
    p = ""
    For i = Len(s) To 1 Step -1
   
        If Asc(Mid(s, i, 1)) > 32 Then Exit For
    Next i
    If i <= 0 Then
        xoadb = ""
    Else
        xoadb = Left(s, i)
    End If
End Function

Sub saveTest(Optional themis As Boolean = False)

Dim t As Table

Dim c As Integer, d As Integer, r As Integer

Dim ten As String, ra As String, thu As String, test As String, thumuc As String, tep As String

thumuc = ActiveDocument.Path
If thumuc = "" Then
    MsgBox "You must save the file"
    Exit Sub
End If
ChDrive (Left(thumuc, 1))
ChDir (thumuc)

For Each t In Selection.Tables
    d = t.Rows.Count
    ten = xoadb(t.Rows(1).Cells(1).Range.Text) ' vd bai1.inp
    ra = xoadb(t.Rows(1).Cells(2).Range.Text) 'neu co thi zip
    
    ten = ten & "."
    thu = Left(ten, InStr(ten, ".") - 1) ' vd bai1
    If Dir(thu, vbDirectory) = "" Then MkDir thu
    
    For r = 2 To d
        If r <= 10 Then
            tep = thu & "\0" + Trim(Str(r - 1))
        Else
            tep = thu & "\" & Trim(Str(r - 1))
        End If
        If themis Then
            MkDir tep
            tep = tep & "\" & thu
        End If
        'MkDir test
        ghitep xoadb(t.Rows(r).Cells(1).Range.Text), tep & ".inp"
        ghitep xoadb(t.Rows(r).Cells(2).Range.Text), tep & ".out"
        
    Next r
    If Len(ra) > 2 Then Shell "zip.exe " & thu & ".zip -r " & thu & "\*.*"
Next t
End Sub

Sub saveTestAsTHN()
    saveTest False
End Sub

Sub saveTestAsThemis()
    saveTest True
End Sub

