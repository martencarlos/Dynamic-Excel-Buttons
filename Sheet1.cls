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
'@Author Carlos Marten 2014/15
' This code creates two command buttons on the first two right adjacent cells.
' The buttons specific word documents
' Requirements:
' - insert this code on the first sheet
' - Have a 'tools' sheet with two specific paths in the cells [B3] and [B4] and listening _
' cells range in cell [B6] with specific format [CR:CR]
' * C = Column, R = Row
'************************************************************************************************

Sub Worksheet_Change(ByVal Target As Range)
    
    Dim Desc_path As String
    Dim Cov_letter_path As String
    Dim KeyCells As Range
    
    ' CELLS THAT HAVE THE LISTENERS
    Set KeyCells = Range(ThisWorkbook.Sheets("tools").Range("B6").Value)
    
    If Not Application.Intersect(KeyCells, Range(Target.Address)) Is Nothing Then
            ' DELETES BUTTON, VALUES AND PROCEDURES IF NAME IS EMPTY
            If IsEmpty(Target.Value) Then
                If (Not IsEmpty(Target.Offset(0, 1).Value)) Then
                    ThisWorkbook.Sheets(1).Shapes(Target.Offset(0, 1).Value).Delete
                    procName = (Target.Offset(0, 1).Value & "_Click")
                    DeleteProcedureCode ThisWorkbook, ThisWorkbook.Sheets(1).CodeName, procName
                    Target.Offset(0, 1).Value = ""
                    
                    ThisWorkbook.Sheets(1).Shapes(Target.Offset(0, 2).Value).Delete
                    procName = (Target.Offset(0, 2).Value & "_Click")
                    DeleteProcedureCode ThisWorkbook, ThisWorkbook.Sheets(1).CodeName, procName
                    Target.Offset(0, 2).Value = ""
                End If
            Else
                ' SET PATHS
                Desc_path = ThisWorkbook.Sheets("tools").Range("B3").Value & Target.Value & ".docx"
                Cov_letter_path = ThisWorkbook.Sheets("tools").Range("B4").Value & Target.Value & ".docx"

                
                ' CHECK IF FILES EXIST
                If Len(Dir(Desc_path)) = 0 Then
                    MsgBox "Full Description File does not exist"
                ElseIf Len(Dir(Cov_letter_path)) = 0 Then
                    MsgBox "Cover letter File does not exist"
                Else
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
                    Dim button_D As OLEObject, N%
                    Dim cell_D As Range
                    
                    
                    Set cell_D = ThisWorkbook.Sheets(1).Range(Target.Offset(0, 1).Address)  'target address +1 column
                    Set button_D = ThisWorkbook.Sheets(1).OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False)
                    
                    button_D.Object.Caption = "Open File"
                    button_D.Name = Target.Value & "_D"
                    button_D.Top = cell_D.Top
                    button_D.Left = cell_D.Left
                    button_D.Width = cell_D.Width
                    button_D.Height = cell_D.RowHeight
                    button_D.Placement = xlMoveAndSize
                    
                    Desc_path = Replace(Desc_path, Environ$("UserName"), """&Environ$(""UserName"")&""")
                    
                    With ThisWorkbook.VBProject.VBComponents(ThisWorkbook.Sheets(1).CodeName).CodeModule
                             N = .CountOfLines
                            .InsertLines N + 1, "Private Sub " & Target.Value & "_D" & "_Click()"
                            .InsertLines N + 2, vbTab & "Set wordapp = CreateObject(" & """" & "word.Application" & """" & ")"
                            .InsertLines N + 3, vbTab & "If Not FileLocked(" & """" & Desc_path & """" & ") Then"
                            .InsertLines N + 4, vbTab & vbTab & "wordapp.documents.Open " & """" & "file://\\" & Desc_path & """"
                            .InsertLines N + 5, vbTab & vbTab & "wordapp.Visible = True"
                            .InsertLines N + 6, vbTab & vbTab & "AppActivate (wordapp.Windows(1).Caption)"
                            .InsertLines N + 7, vbTab & "End If"
                            .InsertLines N + 8, "End Sub"
                    End With
                    
                    Target.Offset(0, 1).Value = Target.Value & "_D"
                    
                    ' CREATE COVER LETTER COMMAND BUTTON
                    Dim button_CL As OLEObject, M%
                    Dim cell_CL As Range
                    
                    Set cell_CL = ThisWorkbook.Sheets(1).Range(Target.Offset(0, 2).Address)  'target address +2 column
                    Set button_CL = ThisWorkbook.Sheets(1).OLEObjects.Add(ClassType:="Forms.CommandButton.1", Link:=False, DisplayAsIcon:=False)
                    
                    button_CL.Object.Caption = "Open File"
                    button_CL.Name = Target.Value & "_CL"
                    button_CL.Top = cell_CL.Top
                    button_CL.Left = cell_CL.Left
                    button_CL.Width = cell_CL.Width
                    button_CL.Height = cell_CL.RowHeight
                    button_CL.Placement = xlMoveAndSize
                    
                    Cov_letter_path = Replace(Cov_letter_path, Environ$("UserName"), """&Environ$(""UserName"")&""")
                    
                    With ThisWorkbook.VBProject.VBComponents(ThisWorkbook.Sheets(1).CodeName).CodeModule
                             M = .CountOfLines
                            .InsertLines M + 1, "Private Sub " & Target.Value & "_CL" & "_Click()"
                            .InsertLines M + 2, vbTab & "Set wordapp = CreateObject(" & """" & "word.Application" & """" & ")"
                            .InsertLines M + 3, vbTab & "If Not FileLocked(" & """" & Cov_letter_path & """" & ") Then"
                            .InsertLines M + 4, vbTab & vbTab & "wordapp.documents.Open " & """" & "file://\\" & Cov_letter_path & """"
                            .InsertLines M + 5, vbTab & vbTab & "wordapp.Visible = True"
                            .InsertLines M + 6, vbTab & vbTab & "AppActivate (wordapp.Windows(1).Caption)"
                            .InsertLines M + 7, vbTab & "End If"
                            .InsertLines M + 8, "End Sub"
                    End With
                   
                    Target.Offset(0, 2).Value = Target.Value & "_CL"
                End If
            End If
    End If
End Sub

' DELETES THE SPECIFIED PROCEDURE FROM THE SPECIFIED WORKBOOK AND MODULE/SHEET
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
