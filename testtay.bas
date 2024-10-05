Attribute VB_Name = "testtay"
'Lam test tay cho themis
'Written by Tran Huu Nam - huunam0@gmail.com
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

Dim C As Integer, d As Integer, r As Integer

Dim ten As String, ra As String, thu As String, test As String, thumuc As String, tep As String
Dim vidu As String, intxt As String, outtxt As String
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
    ten = LCase(ten)
    'ten = Replace(ten, ".inp", "")
    ra = ten 'xoadb(t.Rows(1).Cells(2).Range.Text) 'neu co thi zip
    
    ten = ten & "."
    thu = Left(ten, InStr(ten, ".") - 1) ' vd bai1
    If Dir(thumuc + "\" + thu, vbDirectory) = "" Then MkDir thumuc + "\" + thu
    
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
        intxt = t.Rows(r).Cells(1).Range.Text
        outtxt = t.Rows(r).Cells(2).Range.Text
        intxt = xoadb(intxt)
        outtxt = xoadb(outtxt)
        If Len(intxt) = 0 Or Len(outtxt) = 0 Then Exit For
        ghitep intxt, tep & ".inp"
        ghitep outtxt, tep & ".out"
        'If r <= 4 Then vidu = vidu & "## VD" & (r - 1) & vbCrLf & "### Input" & vbCrLf & "```" & vbCrLf & intxt & vbCrLf & "```" & vbCrLf & "### Output" & vbCrLf & "```" & vbCrLf & outtxt & vbCrLf & "```" & vbCrLf
        If r <= 4 Then
            vidu = vidu & "## Ví d" & ChrW(7909)
            If r > 2 Then vidu = vidu & " " & (r - 1)
            vidu = vidu & ":" & vbCrLf & "### Input" & vbCrLf & "```" & vbCrLf & intxt & vbCrLf & "```" & vbCrLf & "### Output" & vbCrLf & "```" & vbCrLf & outtxt & vbCrLf & "```" & vbCrLf
        End If
    Next r
    If Len(ra) > 2 Then Shell "tar -a -c -f " & thu & ".zip  " & thu & "\*.*"
    'Rem zip -r ..\%tm%.zip *.*
    'tar -a -c -f ..\%tm%.zip *.*
    t.Rows(1).Cells(2).Range.Text = thumuc + "\" + thu + ".zip"
Next t
Selection.EndKey Unit:=wdStory
Selection.InsertAfter vidu
End Sub

Sub saveTestAsTHN()
    saveTest False
End Sub

Sub saveTestAsThemis()
    saveTest True
End Sub

Function doctep(tep As String) As String
    Dim s As String
    Open tep For Input As #1
    s = Input(LOF(1), 1)
    Close 1
    While Right(s, 2) = vbCrLf
        s = Left(s, Len(s) - 2)
    Wend
    doctep = s
End Function
Function SelectFolder()
    Dim FldrPicker As FileDialog
    Dim myFolder As String
    
    'Have User Select Folder to Save to with Dialog Box
      Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)
    
      With FldrPicker
        .Title = "Select A Target Folder"
        .AllowMultiSelect = False
        If .Show <> -1 Then myFolder = "" 'Check if user clicked cancel button
        myFolder = .SelectedItems(1) & "\"
      End With
      
    'Carry out rest of your code here....
    SelectFolder = myFolder
End Function
Function GetFilesIn(Folder As String) As Collection
  Dim F As String
  Set GetFilesIn = New Collection
  F = Dir(Folder & "\*")
  Do While F <> ""
    GetFilesIn.Add F
    F = Dir
  Loop
End Function
Function GetFoldersIn(Folder As String) As Collection
  Dim F As String
  Set GetFoldersIn = New Collection
  F = Dir(Folder & "\*", vbDirectory)
  Do While F <> ""
    If GetAttr(Folder & "\" & F) And vbDirectory Then
        If Left(F, 1) <> "." Then GetFoldersIn.Add F
    End If
    F = Dir
  Loop
End Function
Function GetRightFolder(fname) As String
    Dim a
    If Right(fname, 1) <> "\" Then fname = fname + "\"
    a = Split(fname, "\")
    GetRightFolder = a(UBound(a) - 1)
End Function
Sub openThemisTest()
    Dim t As Table
    Dim tm As String, bai As String
    Dim C As Collection, F
    tm = SelectFolder
    If tm = "" Then Exit Sub
    If Right(tm, 1) <> "\" Then tm = tm + "\"
    bai = GetRightFolder(tm)
    Set C = GetFoldersIn(tm)
    
    If ActiveDocument.Tables.Count = 0 Then ActiveDocument.Tables.Add Selection.Range, 2, 2
    Set t = ActiveDocument.Tables(1)
    t.Rows(1).Cells(1).Range.Text = bai
    Dim tt As Integer
    tt = 2
    For Each F In C
        If t.Rows.Count < tt Then t.Rows.Add
        t.Rows(tt).Cells(1).Range.Text = doctep(tm + F + "\" + bai + ".inp")
        t.Rows(tt).Cells(2).Range.Text = doctep(tm + F + "\" + bai + ".out")
        tt = tt + 1
    Next F
End Sub
Sub test()
    'MsgBox GetRightFolder("D:\SINHTEST\BANGTAY\flow")
    Dim s As String
    s = doctep("D:\SINHTEST\BANGTAY\flow\01\flow.inp")
    s = Left(s, Len(s) - 2)
    MsgBox Asc(Right(s, 1))
End Sub
