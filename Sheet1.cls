VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'************************************************************************************************
'@Author Carlos Marten 2015
' This code creates two command buttons on the first two right adjacent cells.
' These dynamically created buttons are linked to blank Word Documents
' Requirements:
' - Enable developer mode
' - Enable macros
' - Add VBA Extensibility and Word references
' - Windows 7 or newer
' - insert this code on the first sheet
' - Have a 'tools' sheet with two specific paths in the cells [B3] and [B4] and listening _
' cells range in cell [B6] with specific format [CR:CR]
' * C = Column, R = Row
'************************************************************************************************

Sub Worksheet_Change(ByVal Target As Range)
    
    Dim FilePath1 As String
    Dim FilePath2 As String
    Dim KeyCells As Range
    
    If Target.Cells.Count > 1 Then Exit Sub

    ' CELLS THAT HAVE THE LISTENERS
    Set KeyCells = Range(ThisWorkbook.Sheets("tools").Range("B6").Value)
    
    If Not Application.Intersect(KeyCells, Range(Target.Address)) Is Nothing Then
            If IsEmpty(Target.Value) Then
                ' DELETES BUTTON, VALUES AND PROCEDURES
                If (Not IsEmpty(Target.Offset(0, 1).Value)) Then
                    ThisWorkbook.Sheets(1).Shapes(Target.Offset(0, 1).Value).Delete
                    procName = (Target.Offset(0, 1).Value & "_Click")
                    DeleteProcedureCode ThisWorkbook, ThisWorkbook.Sheets(1).CodeName, procName
                    
                    ThisWorkbook.Sheets(1).Shapes(Target.Offset(0, 2).Value).Delete
                    procName = (Target.Offset(0, 2).Value & "_Click")
                    DeleteProcedureCode ThisWorkbook, ThisWorkbook.Sheets(1).CodeName, procName
                    
                    ' SET PATHS
                    FilePath1 = ThisWorkbook.Sheets("tools").Range("B3").Value & Left(Target.Offset(0, 1).Value, Len(Target.Offset(0, 1).Value) - 2) & ".docx"
                    FilePath2 = ThisWorkbook.Sheets("tools").Range("B4").Value & Left(Target.Offset(0, 1).Value, Len(Target.Offset(0, 1).Value) - 2) & ".docx"
                    
                    ' DELETE FILES
                    DeleteFile FilePath1
                    DeleteFile FilePath2
                    
                    Target.Offset(0, 1).Value = ""
                    Target.Offset(0, 2).Value = ""
                End If
            Else
                ' SET PATHS
                FilePath1 = ThisWorkbook.Sheets("tools").Range("B3").Value & Target.Value & ".docx"
                FilePath2 = ThisWorkbook.Sheets("tools").Range("B4").Value & Target.Value & ".docx"

                
                ' CHECK IF FILES EXIST
                If (Len(Dir(FilePath1)) <> 0) And (Len(Dir(FilePath2)) <> 0) Then
                    MsgBox "Filename """ & Target.Value & """ already in use"
                    Target.Value = ""
                Else
                    ' CREATE TWO BLANK WORD DOCUMENTS
                    CreateNewWordDoc FilePath1
                    CreateNewWordDoc FilePath2
                    ' IF BUTTONS CREATED, DELETE BUTTONS, VALUES AND PROCEDURES
                    If Not IsEmpty(Target.Offset(0, 1).Value) Then
                        ThisWorkbook.Sheets(1).Shapes(Target.Offset(0, 1).Value).Delete
                        procName = (Target.Offset(0, 1).Value & "_Click")
                        DeleteProcedureCode ThisWorkbook, ThisWorkbook.Sheets(1).CodeName, procName
                        Target.Offset(0, 1).Value = ""
                        
                        ThisWorkbook.Sheets(1).Shapes(Target.Offset(0, 2).Value).Delete
                        procName = (Target.Offset(0, 2).Value & "_Click")
                        DeleteProcedureCode ThisWorkbook, ThisWorkbook.Sheets(1).CodeName, procName
                        Target.Offset(0, 2).Value = ""
                    End If
                    
                    ' CREATE DESCRIPTION COMMAND BUTTON
                    Dim button_1 As OLEObject, N%
                    Dim cell_1 As Range
                    
                    
                    Set cell_1 = ThisWorkbook.Sheets(1).Range(Target.Offset(0, 1).Address)  'target address +1 column
                    Set button_1 = ThisWorkbook.Sheets(1).OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False)
                    
                    button_1.Object.Caption = "Open File"
                    button_1.Name = Target.Value & "_1"
                    button_1.Top = cell_1.Top
                    button_1.Left = cell_1.Left
                    button_1.Width = cell_1.Width
                    button_1.Height = cell_1.RowHeight
                    button_1.Placement = xlMoveAndSize
                    
                    FilePath1 = Replace(FilePath1, Environ$("UserName"), """&Environ$(""UserName"")&""")
                    
                    With ThisWorkbook.VBProject.VBComponents(ThisWorkbook.Sheets(1).CodeName).CodeModule
                             N = .CountOfLines
                            .InsertLines N + 1, "Private Sub " & Target.Value & "_1" & "_Click()"
                            .InsertLines N + 2, vbTab & "Set wordapp = CreateObject(" & """" & "word.Application" & """" & ")"
                            .InsertLines N + 3, vbTab & "If Not FileLocked(" & """" & FilePath1 & """" & ") Then"
                            .InsertLines N + 4, vbTab & vbTab & "wordapp.documents.Open " & """" & "file://\\" & FilePath1 & """"
                            .InsertLines N + 5, vbTab & vbTab & "wordapp.Visible = True"
                            .InsertLines N + 6, vbTab & vbTab & "AppActivate (wordapp.Windows(1).Caption)"
                            .InsertLines N + 7, vbTab & "End If"
                            .InsertLines N + 8, "End Sub"
                    End With
                    
                    Target.Offset(0, 1).Value = Target.Value & "_1"
                    
                    ' CREATE COVER LETTER COMMAND BUTTON
                    Dim button_2 As OLEObject, M%
                    Dim cell_2 As Range
                    
                    Set cell_2 = ThisWorkbook.Sheets(1).Range(Target.Offset(0, 2).Address)  'target address +2 column
                    Set button_2 = ThisWorkbook.Sheets(1).OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False)
                    
                    button_2.Object.Caption = "Open File"
                    button_2.Name = Target.Value & "_2"
                    button_2.Top = cell_2.Top
                    button_2.Left = cell_2.Left
                    button_2.Width = cell_2.Width
                    button_2.Height = cell_2.RowHeight
                    button_2.Placement = xlMoveAndSize
                    
                    FilePath2 = Replace(FilePath2, Environ$("UserName"), """&Environ$(""UserName"")&""")
                    
                    With ThisWorkbook.VBProject.VBComponents(ThisWorkbook.Sheets(1).CodeName).CodeModule
                             M = .CountOfLines
                            .InsertLines M + 1, "Private Sub " & Target.Value & "_2" & "_Click()"
                            .InsertLines M + 2, vbTab & "Set wordapp = CreateObject(" & """" & "word.Application" & """" & ")"
                            .InsertLines M + 3, vbTab & "If Not FileLocked(" & """" & FilePath2 & """" & ") Then"
                            .InsertLines M + 4, vbTab & vbTab & "wordapp.documents.Open " & """" & "file://\\" & FilePath2 & """"
                            .InsertLines M + 5, vbTab & vbTab & "wordapp.Visible = True"
                            .InsertLines M + 6, vbTab & vbTab & "AppActivate (wordapp.Windows(1).Caption)"
                            .InsertLines M + 7, vbTab & "End If"
                            .InsertLines M + 8, "End Sub"
                    End With
                   
                    Target.Offset(0, 2).Value = Target.Value & "_2"
                End If
            End If
    End If
End Sub

Sub CreateNewWordDoc(ByVal filePath As String)
    Set wrdApp = CreateObject("Word.Application")
    Set wrdDoc = wrdApp.Documents.Add
    With wrdDoc
        .SaveAs (filePath)
        .Close ' close the document
    End With
    wrdApp.Quit ' close the Word application
    Set wrdDoc = Nothing
    Set wrdApp = Nothing
End Sub

Sub DeleteFile(ByVal FileToDelete As String)
   If Len(Dir(FileToDelete)) <> 0 Then
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Sub

Sub DeleteProcedureCode(ByVal wb As Workbook, _
    ByVal DeleteFromModuleName As String, ByVal ProcedureName As String)
Dim VBCM As CodeModule, ProcStartLine As Long, ProcLineCount As Long
    On Error Resume Next
    Set VBCM = wb.VBProject.VBComponents(DeleteFromModuleName).CodeModule
    If Not VBCM Is Nothing Then
        Dim OrocStartLine As Integer
        ProcStartLine = 0
        ProcStartLine = VBCM.ProcStartLine(ProcedureName, vbext_pk_Proc)
        If ProcStartLine > 0 Then
            ProcLineCount = VBCM.ProcCountLines(ProcedureName, vbext_pk_Proc)
            VBCM.DeleteLines ProcStartLine, ProcLineCount
        End If
        Set VBCM = Nothing
    End If
    On Error GoTo 0
End Sub

' CHECKS WEATHER A SPECIFIC FILE IS ALREADY OPEN BY ANOTHER PROCESS
Function FileLocked(strFileName As String) As Boolean
   On Error Resume Next
   Open strFileName For Binary Access Read Write Lock Read Write As #1
   Close #1
   If Err.Number <> 0 Then
      MsgBox "Word document may already be open or not available"
      FileLocked = True
      Err.Clear
   End If
End Function