VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTaskList 
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   10770
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lstTasks 
      Height          =   3750
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   6615
      SortKey         =   1
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Task"
         Object.Width           =   5365
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Added"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Completed"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmTaskList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents evtVBProjects As VBProjectsEvents
Attribute evtVBProjects.VB_VarHelpID = -1
Public WithEvents evtVBFiles As FileControlEvents
Attribute evtVBFiles.VB_VarHelpID = -1

Private mintCurrSelect As Integer

Private Sub evtVBFiles_AfterWriteFile(ByVal VBProject As VBIDE.VBProject, ByVal FileType As VBIDE.vbext_FileType, ByVal FileName As String, ByVal Result As Integer)
   
   Dim strPath As String
   
   On Error GoTo evtVBFiles_AfterWriteFile_Error
   
   If FileType = vbext_ft_Project Then
      If Len(gstrTaskFile) > 0 Then
         WriteTasks
      Else
         GetTaskFile FileName
         WriteTasks
      End If
   End If

Exit Sub

evtVBFiles_AfterWriteFile_Error:
   DoError "frmTaskList", "evtVBFiles_AfterWriteFile", Err
    
End Sub

Private Sub evtVBProjects_ItemActivated(ByVal VBProject As VBIDE.VBProject)

   Dim strPath As String
   
   '--Save the current task list
   WriteTasks
   
   '--Retrieve the tasks for the newly active project
   strPath = gobjVBInstance.ActiveVBProject.FileName
   GetTaskFile strPath
   
End Sub

Private Sub evtVBProjects_ItemAdded(ByVal VBProject As VBIDE.VBProject)
   
   Dim strPath As String
   
   strPath = VBProject.FileName
   
   If Len(strPath) = 0 Then
      Exit Sub
   End If
   
   If Len(gstrTaskFile) > 0 Then
      WriteTasks
   Else
      GetTaskFile strPath
   End If

End Sub

Private Sub evtVBProjects_ItemRemoved(ByVal VBProject As VBIDE.VBProject)

   '--Save the current tasks and clear the list
   WriteTasks
   gstrTaskFile = vbNullString
   lstTasks.ListItems.Clear

End Sub

Private Sub lstTasks_Click()
   
   On Error GoTo lstTasks_Click_Error

   If gblMouseClick = vbRightButton Then
      Form.PopupMenu mnuMenu
   End If

   Exit Sub
   
lstTasks_Click_Error:
   DoError "frmTaskList", "lstTasks_Click", Err
    
End Sub

Private Sub lstTasks_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

   lstTasks.SortKey = ColumnHeader.Index - 1
   lstTasks.Sorted = True
   If lstTasks.SortOrder = lvwDescending Then
      lstTasks.SortOrder = lvwAscending
   Else
      lstTasks.SortOrder = lvwDescending
   End If
   lstTasks.Refresh

End Sub

Private Sub lstTasks_DblClick()
   
   On Error GoTo lstTasks_DblClick_Error
   
   lstTasks.ListItems.Add
   lstTasks.StartLabelEdit
   lstTasks.ListItems.Item(1).SubItems(1) = Format(Now, "mm/dd/yyyy")

lstTasks_DblClick_Error:
   DoError "frmTaskList", "lstTasks_DblClick", Err
   
End Sub

Private Sub lstTasks_ItemCheck(ByVal Item As MSComctlLib.ListItem)

   On Error GoTo lstTasks_ItemCheck_Error
   
   If Item.Checked = True Then
      Item.ForeColor = vbGrayText
      Item.SubItems(2) = Format(Now, "mm/dd/yyyy")
   Else
      Item.ForeColor = vbBlack
      Item.SubItems(2) = ""
   End If
    
lstTasks_ItemCheck_Error:
    DoError "frmTaskList", "lstTasks_ItemCheck", Err
    
End Sub

Private Sub lstTasks_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
   mintCurrSelect = Item.Index

End Sub

Private Sub lstTasks_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   On Error GoTo lstTasks_MouseDown_Error

   gblMouseClick = Button
    
   Exit Sub
   
lstTasks_MouseDown_Error:
   DoError "frmTaskList", "lstTasks_MouseDown", Err

End Sub

Private Sub mnuView_Click()
   
   Load frmView
   
   With frmView
      .InitForm lstTasks, gobjVBInstance.ActiveVBProject.Description
      .Show
   End With

End Sub

Private Sub mnuAbout_Click()
    
   frmAbout.Show

End Sub

Private Sub mnuNew_Click()

   Dim strTask As String
   
   On Error GoTo mnuNew_Click_Error

   strTask = InputBox("Enter new Task:", "New Task")
   lstTasks.ListItems.Add , , strTask
   lstTasks.ListItems.Item(1).SubItems(1) = Format(Now, "mm/dd/yyyy")
    
   Exit Sub
   
mnuNew_Click_Error:
   DoError "frmTaskList", "mnuNew_Click", Err
   
End Sub

Private Sub Form_Initialize()
   
   Dim strPath As String
   
   On Error GoTo Form_Initialize_Error

   Set Me.evtVBProjects = gobjVBInstance.Events.VBProjectsEvents
   Set Me.evtVBFiles = gobjVBInstance.Events.FileControlEvents(Nothing)
   
   If Not (gobjVBInstance Is Nothing) Then
      If gobjVBInstance.VBProjects.Count = 0 Then
         Exit Sub
      Else
         strPath = gobjVBInstance.ActiveVBProject.FileName
         GetTaskFile strPath
      End If
   End If
   
   Exit Sub
   
Form_Initialize_Error:
    DoError "frmTaskList", "Form - Initialize", Err
    
End Sub

Private Sub mnuDelete_Click()

   On Error GoTo mnuDelete_Error
   
   If mintCurrSelect <> 0 Then
      If MsgBox("Delete task: " & lstTasks.ListItems(lstTasks.SelectedItem.Index).Text & "?", _
                vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
         lstTasks.ListItems.Remove (mintCurrSelect)
         WriteTasks
      End If
   End If
   
   Exit Sub
   
mnuDelete_Error:
   DoError "frmTaskList", "mnuDelete", Err
    
End Sub

Public Sub LoadTasks()
   
   Dim lngCount As Long
   Dim lngIndex As Long
   Dim lstItem As ListItem
   Dim strBuffer As String
   Dim strSDesc As String
   Dim strSAdd As String
   Dim strSEnd As String
   
   On Error GoTo LoadTasks_Error
   
   lstTasks.ListItems.Clear
   
   lngCount = CLng(GetFromIni("Tasks", "Count", gstrTaskFile))
   
   For lngIndex = 1 To lngCount
      strBuffer = GetFromIni("tasks", "task" & lngIndex, gstrTaskFile)
      Set lstItem = lstTasks.ListItems.Add
      If Mid(strBuffer, 1, 1) = "*" Then
         lstItem.Checked = True
         lstItem.ForeColor = vbGrayText
         strBuffer = Mid(strBuffer, 2)
         strSDesc = GetToken(strBuffer, "|")
         strSAdd = GetToken(strBuffer, "|")
         strSEnd = strBuffer
         lstItem.Text = strSDesc
         
         lstItem.SubItems(1) = strSAdd
         lstItem.SubItems(2) = strSEnd
      Else
         strSDesc = GetToken(strBuffer, "|")
         strSAdd = GetToken(strBuffer, "|")
         strSEnd = strBuffer
         lstItem.Text = strSDesc
         
         lstItem.SubItems(1) = strSAdd
         lstItem.SubItems(2) = strSEnd
      End If
   Next lngIndex
   
   lstTasks.SortOrder = lvwAscending
   lstTasks.SortKey = 1
   lstTasks.Sorted = True
   
LoadTasks_Error:
   DoError "frmTaskList", "LoadTasks", Err
   
End Sub

Function StripNameFromPath(ByVal pstrSearchstring As String) As String
   
   Dim strTest As String
   Dim intLastSlashPos As Integer
   Dim i As Integer
   Dim intMyPos As Integer
   Dim strSearchchar As String
   
   On Error GoTo StripNameFromPath_Error
   
   strTest = "NULL"
   strSearchchar = "\"
   
   For i = 1 To Len(pstrSearchstring)
      intMyPos = InStr(i, pstrSearchstring, strSearchchar, 1)
      If intMyPos = 0 Then
         If strTest = "NULL" Then
            strTest = Str(i)
         End If
      End If
   Next i
   
   intLastSlashPos = Val(strTest)
   StripNameFromPath = Mid(pstrSearchstring, 1, intLastSlashPos - 1)
   
   Exit Function
   
StripNameFromPath_Error:
   DoError "frmTaskList", "StripNameFromPath", Err
   
End Function


Private Sub Form_Resize()
   
   On Error Resume Next
   
   lstTasks.Height = Form.Height - 75
   lstTasks.Width = Form.Width - 90
   lstTasks.ColumnHeaders(1).Width = lstTasks.Width - 1440 - 1440

End Sub

Private Sub GetTaskFile(strProjectPath As String)
'--Loads the task file if it doesn't exits, otherwise
'  creates a new one.

   Dim intFileNum As Integer
   
   On Error GoTo GetTaskFile_Error
   
   strProjectPath = StripNameFromPath(strProjectPath)
   gstrTaskFile = strProjectPath & TASK_FILE_NAME
   
   If FileExists(gstrTaskFile) = False Then
      intFileNum = FreeFile
      Open gstrTaskFile For Output As #intFileNum
      Print #intFileNum, "[Tasks]"
      Print #intFileNum, "Count=0"
      Close #intFileNum
   Else
      LoadTasks
   End If
   
   Exit Sub
   
GetTaskFile_Error:
   DoError "frmTaskList", "GetTaskFile", Err
   
End Sub

Private Sub WriteTasks()
   
   Dim lngCount As Long
   Dim lngIndex As Long
   Dim strBuffer As String
   
   On Error GoTo WriteTasks_Error
   
   lngCount = lstTasks.ListItems.Count
   WriteToIni "Tasks", "Count", CStr(lngCount), gstrTaskFile
   For lngIndex = 1 To lngCount
      If lstTasks.ListItems.Item(lngIndex).Checked Then
         WriteToIni "Tasks", "Task" & lngIndex, "*" _
                  & lstTasks.ListItems.Item(lngIndex).Text _
                  & "|" & lstTasks.ListItems.Item(lngIndex).SubItems(1) _
                  & "|" & lstTasks.ListItems.Item(lngIndex).SubItems(2), _
                  gstrTaskFile
      Else
         WriteToIni "Tasks", "Task" & lngIndex, lstTasks.ListItems.Item(lngIndex).Text _
                  & "|" & lstTasks.ListItems.Item(lngIndex).SubItems(1) _
                  & "|" & lstTasks.ListItems.Item(lngIndex).SubItems(2), _
                  gstrTaskFile
      End If
   Next lngIndex
   
   Exit Sub
   
WriteTasks_Error:
   DoError "frmTaskList", "WriteTasks", Err
   
End Sub



