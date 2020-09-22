VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Tasks 
   ClientHeight    =   6225
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   11190
   _ExtentX        =   19738
   _ExtentY        =   10980
   _Version        =   393216
   Description     =   $"Tasks.dsx":0000
   DisplayName     =   "Task List for VB 6.0"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Command Line / Startup"
   LoadBehavior    =   5
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Tasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim mcbMenuCommandBar As Office.CommandBarControl
Dim mdocAddIn As docAddIn

Public WithEvents MenuHandler As CommandBarEvents          'command bar event handler
Attribute MenuHandler.VB_VarHelpID = -1

Const guidTasks$ = "AB3075C1-B54F-11d3-941A-00A0CC547B23"

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

   Set gobjVBInstance = Application
   Set mcbMenuCommandBar = AddToAddInCommandBar("Task List")
   Set Me.MenuHandler = gobjVBInstance.Events.CommandBarEvents(mcbMenuCommandBar)
   If gwinWindow Is Nothing Then
      Set gwinWindow = gobjVBInstance.Windows.CreateToolWindow(AddInInst, _
                                                              "TaskList.docAddIn", _
                                                              "Tasks", _
                                                              guidTasks, _
                                                              mdocAddIn)
   End If
   gwinWindow.Visible = True

End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)

   mcbMenuCommandBar.Delete

   Set mdocAddIn = Nothing

End Sub

Sub Show()

   gwinWindow.Visible = True

End Sub

Sub Hide()

   gwinWindow.Visible = False

End Sub

Function AddToAddInCommandBar(sCaption As String) As Office.CommandBarControl

   Dim cbMenuCommandBar As Office.CommandBarControl  'command bar object
   Dim cbMenu As Object

   On Error GoTo AddToAddInCommandBarErr

   Set cbMenu = gobjVBInstance.CommandBars("Tools")
   If cbMenu Is Nothing Then
      Exit Function
   End If

   Set cbMenuCommandBar = cbMenu.Controls.Add(1)
   cbMenuCommandBar.Caption = sCaption

   Set AddToAddInCommandBar = cbMenuCommandBar

   Exit Function

AddToAddInCommandBarErr:

End Function

Private Sub MenuHandler_Click(ByVal CommandBarControl As Object, Handled As Boolean, CancelDefault As Boolean)

   If CommandBarControl.Caption = "Task List" Then
      Me.Show
   End If

End Sub
