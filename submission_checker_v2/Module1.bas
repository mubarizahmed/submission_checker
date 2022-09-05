Attribute VB_Name = "Module1"
Dim Resume_flag As Boolean

Sub Import_Click()
    Dim path As String
    Dim sheet As Worksheet
    Dim textLine As String
    Dim fields() As String
    Dim startRow, startCol, i, Length  As Integer
    
    path = ActiveWorkbook.path
    path = path + "\Report.csv"
    Set sheet = ActiveSheet
    startRow = 3
    startCol = 3
    
    Do Until IsEmpty(sheet.Cells(startRow, startCol))
        startRow = startRow + 1
    Loop
    

    
    Open path For Input As #1
    Line Input #1, textLine ' Read line into variable.
    
    Do While Not EOF(1) ' Loop until end of file.
        Line Input #1, textLine ' Read line into variable.
        fields = Split(textLine, ",") 'split line into fields
        'put fields into sheet
        Length = UBound(fields, 1)
        For i = 0 To UBound(fields, 1)
            sheet.Cells(startRow, startCol + i) = fields(i)
        Next i
        startRow = startRow + 1
    Loop
    Close #1
End Sub
Sub Hide_Click()
    Dim start As Range
    Dim i As Integer
    
    Set col = Range("G:J")
    
    If col.EntireColumn.Hidden = False Then
        For i = 1 To 13
            col.EntireColumn.Hidden = True
            If col.Cells(3, 2).Value = 0 Then
                col.Offset(0, 1).EntireColumn.Hidden = True
                col.Offset(0, 2).EntireColumn.Hidden = True
            End If
            Set col = col.Offset(0, 6)
        Next

    Else
        For i = 1 To 13
            col.EntireColumn.Hidden = False
            col.Offset(0, 1).EntireColumn.Hidden = False
            col.Offset(0, 2).EntireColumn.Hidden = False
            'col.Offset(0, 3).EntireColumn.Hidden = False
            Set col = col.Offset(0, 6)
        Next
    End If
    
End Sub

Sub Check()
    Dim FileSystem As Object
    Dim HostFolder As String

    HostFolder = ActiveWorkbook.path
    'MsgBox (HostFolder)
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    DoFolder FileSystem.GetFolder(HostFolder)
End Sub

Sub DoFolder(Folder)
    
    Dim Subfolder
    Dim File
    Dim Foldername() As String
    Dim Studentname As String
    Dim HW_No As Integer
    
    HW_No = Worksheets("Grading").Range("C1").Value

    'MsgBox (HW_No)
    
    For Each Subfolder In Folder.SubFolders
        ' match student data
        'MsgBox (Subfolder)
        Foldername = Split(Split(Subfolder, "\")(UBound(Split(Subfolder, "\"))), "_")
        'MsgBox (Foldername(0))
        Studentname = Replace(Foldername(0), "-", " ")
        MsgBox (Studentname)
        
        Dim File_No As Integer
        Dim File_Found As Boolean
        Dim File_HWnum As Integer
        Dim File_Stuname As String
        Dim File_StuID As Integer
        
        File_No = 0
        File_Found = False
        
        'Error flags
        Dim Nonmandatory As Boolean
        Dim Invalid_Name As Boolean
        Dim Invalid_ID As Boolean
        
        Nonmandatory = False
        Invalid_Name = False
        Invalid_ID = False
        
        Resume_flag = False
        
        For Each File In Subfolder.Files

            'MsgBox (File)
            
            File_No = File_No + 1
            
            Dim line As String
            Dim linearr() As String
            Dim line2 As Variant
            Dim Endflag As Integer
            Endflag = 0
            Const adReadLine = -2&

            With CreateObject("ADODB.Stream")
                .Open
                If IsUnicodeFile(File) Then
                    .Charset = "utf-8"
                End If
                .LoadFromFile File

                    
                line = .ReadText()
                'line = Module2.GetFileText(File, utf16)
                'line = .ReadAllText(adReadLine)

                
                
                MsgBox line
                If InStr(1, line, vbCrLf, vbTextCompare) <> 0 Then
                    linearr = Split(line, vbCrLf)
                ElseIf InStr(1, line, vbCr, vbTextCompare) <> 0 Then
                    linearr = Split(line, vbCr)
                ElseIf InStr(1, line, vbLf, vbTextCompare) <> 0 Then
                    linearr = Split(line, vbLf)
                End If
                '.Close
                'Get header info
                For Each line2 In linearr
                    'MsgBox (line2)
                    If InStr(1, line2, "Homework", vbTextCompare) <> 0 Then
                        File_HWnum = CInt(Trim(Mid(line2, InStr(1, line2, "Homework", vbTextCompare) + 9, 1)))
                        Endflag = Endflag + 1
                        'MsgBox (File_HWnum)
                    End If
                    If InStr(1, line2, "Name:", vbTextCompare) <> 0 Then
                        File_Stuname = Trim(Mid(line2, InStr(1, line2, "Name:", vbTextCompare) + 5))
                        Endflag = Endflag + 1
                        'MsgBox (File_Stuname)
                    End If
                    If InStr(1, line2, "Matriculation number:", vbTextCompare) <> 0 Then
                        File_StuID = CInt(Trim(Mid(line2, InStr(1, line2, "Matriculation number:", vbTextCompare) + 22)))
                        'MsgBox (File_StuID)
                        Endflag = Endflag + 1
                    End If
                    If Endflag > 2 Then Exit For
                
                Next line2
                
                'Name & ID validation
                If Not Namecomp(Studentname, File_Stuname) Then
                    Invalid_Name = True
                End If
                If Sheets("Students").Range("b1", Sheets("Students").Range("b1").End(xlDown)).Find(Studentname) Is Nothing Then
                    'MsgBox (Worksheets("Students").Range("b1", Worksheets("Students").Range("b1").End(xlDown)).Find(File_StuID).Address)
                    If Sheets("Students").Range("a1", Sheets("Students").Range("a1").End(xlDown)).Find(File_StuID) Is Nothing Then
                        'MsgBox (Sheets("Students").Range("b1").End(xlDown).Value)
                        Sheets("Students").Range("b1").End(xlDown).Offset(1, 0).Value = Studentname
                        Sheets("Students").Range("b1").End(xlDown).Offset(0, -1).Value = File_StuID
                    Else
                        Invalid_ID = True
                    End If
                Else
                    MsgBox (Sheets("Students").Range("b1", Sheets("Students").Range("b1").End(xlDown)).Find(Studentname).Address)
                    If File_StuID <> Worksheets("Students").Range("b1", Worksheets("Students").Range("b1").End(xlDown)).Find(Studentname).Offset(0, -1).Value Then
                        Invalid_ID = True
                    End If
                    
                End If
                '
                'load relevant tasks
                Dim tasks() As String
                Dim task As Variant
                Dim No_Tasks As Integer
                No_Tasks = 0
                
                tasks = Split(line, "Task")
                For i = 1 To UBound(tasks)
                    tasks(i - 1) = tasks(i)
                Next i
                'ReDim Preserve tasks(UBound(tasks) - 1)
                
                For Each task In tasks
                    Dim task_no As Integer
                    
                    task = Trim(task)
                    'MsgBox (Left(task, InStr(1, task, ":") - 1))
                    task_no = CInt(Left(task, InStr(1, task, ":") - 1))
                    'MsgBox (task_no)
                    
                    If Mandatory(HW_No, task_no) Then
                        No_Tasks = No_Tasks + 1
                        'MsgBox Sheets("Grading").Controls("Text1")
                        Sheets("Grading").OLEObjects("Text" & No_Tasks).Object.Text = task
                        Sheets("Grading").Cells(10 + (20 * (No_Tasks - 1)), 3).Value = task_no
                        
                    Else
                        Nonmandatory = True
                    End If
                    
                Next
                '
                'load errors and names
                Sheets("Grading").Cells(3, 3).Value = Studentname
                Sheets("Grading").Cells(4, 3).Value = File_StuID
                Dim pos As String
                pos = 1
                If Invalid_Name Then
                    Sheets("Grading").Cells(4 + pos, 6).Value = "Please put your name in the file header!"
                    Sheets("Grading").Cells(3, 4).Value = 0
                    pos = pos + 1
                End If
                If Invalid_ID Then
                    Sheets("Grading").Cells(4 + pos, 6).Value = "Please put your matriculation number in the header!"
                    Sheets("Grading").Cells(4, 4).Value = 0
                    pos = pos + 1
                End If
                If Nonmandatory Then
                    Sheets("Grading").Cells(4 + pos, 6).Value = "Please ONLY submit the mandatory tasks!"
                    pos = pos + 1
                End If
                '
                
                
            End With

                
            'stop looping through files
            If File_Found Then Exit For
            
            
            
        Next
        Pause_loop
        'write to sheet
        
    Next
End Sub

Function Namecomp(str1 As String, str2 As String) As Boolean
    Dim strarr1() As String
    Dim strarr2() As String
    
    strarr1 = Split(Trim(Replace(str1, "  ", " ")), " ")
    strarr2 = Split(Trim(Replace(str2, "  ", " ")), " ")
    Namecomp = False
    For i = LBound(strarr1) To UBound(strarr1)
        If InStr(1, str2, strarr1(i), vbTextCompare) Then
            Namecomp = True
            Exit For
        End If
    Next
    
End Function

Function Mandatory(hw As Integer, task As Integer) As Boolean
    
    If Sheets("HW").Range("a1", Sheets("HW").Range("a1").End(xlDown)).Find(hw).Offset(0, task).Value = 1 Then
        Mandatory = True
    Else
        Mandatory = False
    End If
    'MsgBox Mandatory
End Function

Sub Pause_loop()
    Dim counter As Integer
    Do Until Resume_flag
        If (counter < 755) Then
            counter = counter + 1
        Else
            counter = 0
            DoEvents
        End If
    Loop
End Sub

Sub Resume_loop()
    Resume_flag = True
End Sub
Public Function IsUnicodeFile(FilePath)
    Dim objFSO
    Dim objStream

    Dim intAsc1Chr
    Dim intAsc2Chr


    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If (objFSO.FileExists(FilePath) = False) Then
        IsUnicodeFile = False
        Exit Function
    End If

    ' 1=Read-only, False==do not create if not exist, -1=Unicode 0=ASCII
    Set objStream = objFSO.OpenTextFile(FilePath, 1, False, 0)
    intAsc1Chr = Asc(objStream.Read(1))
    intAsc2Chr = Asc(objStream.Read(1))
    objStream.Close

    If (intAsc1Chr = 255) And (intAsc2Chr = 254) Then
        IsUnicodeFile = True
    Else
        IsUnicodeFile = False
    End If

    Set objStream = Nothing
    Set objFSO = Nothing
End Function
