VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   7650
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   11520
   _ExtentX        =   20320
   _ExtentY        =   13494
   _Version        =   393216
   Description     =   "Add/Remove Error Handling"
   DisplayName     =   "vbAutoErrorHandler"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Written by: Jason Stracner (www.lr-tc.com, lrtc.blogspot.com)
Option Explicit
'Thanks to Roger Gilchrist for showing me how to add submenu items to the add-ins menu.
Public VBInstance                                     As VBIDE.VBE
Dim mcbMenuCommandBarAddErrorTraps                    As Office.CommandBarControl
Public WithEvents MenuHandlerAddErrorTraps            As CommandBarEvents          'command bar event handler
Attribute MenuHandlerAddErrorTraps.VB_VarHelpID = -1
Dim mcbMenuCommandBarRemoveErrorTraps                 As Office.CommandBarControl
Public WithEvents MenuHandlerRemoveErrorTraps         As CommandBarEvents
Attribute MenuHandlerRemoveErrorTraps.VB_VarHelpID = -1
Dim mcbMenuCommandBarAddLineNumbers                   As Office.CommandBarControl
Public WithEvents MenuHandlerAddLineNumbers           As CommandBarEvents
Attribute MenuHandlerAddLineNumbers.VB_VarHelpID = -1
Dim mcbMenuCommandBarRemoveLineNumbers                As Office.CommandBarControl
Public WithEvents MenuHandlerRemoveLineNumbers        As CommandBarEvents
Attribute MenuHandlerRemoveLineNumbers.VB_VarHelpID = -1
Dim mcbMenuCommandBarEditTemplate                     As Office.CommandBarControl
Public WithEvents MenuHandlerEditTemplate             As CommandBarEvents
Attribute MenuHandlerEditTemplate.VB_VarHelpID = -1
Dim mcbMenuCommandBarAdd2All                          As Office.CommandBarControl
Public WithEvents MenuHandlerAdd2All                  As CommandBarEvents
Attribute MenuHandlerAdd2All.VB_VarHelpID = -1
Dim mcbMenuCommandBarRemoveFromAll                    As Office.CommandBarControl
Public WithEvents MenuHandlerRemoveFromAll            As CommandBarEvents
Attribute MenuHandlerRemoveFromAll.VB_VarHelpID = -1
Private mcbMenuCommandBarTopMenu                      As Office.CommandBarPopup
Dim mfrmAddIn                                         As New frmAddIn

Public Function funcHandleError(ByVal sModule As String, ByVal sProcedure As String, ByVal oErr As ErrObject) As Long
  Dim Msg As String

  Msg = oErr.Source & " caused error """ & _
            oErr.Description & """ (" & _
            oErr.Number & ")" & vbCrLf & _
            "in module " & sModule & _
            " procedure " & sProcedure & _
            ", line " & Erl & "."
  funcHandleError = MsgBox( _
            Msg, vbAbortRetryIgnore + vbMsgBoxHelpButton + vbCritical, _
            "What to you want to do?", _
            oErr.HelpFile, _
            oErr.HelpContext)
End Function

Sub Hide()
  Unload mfrmAddIn
End Sub

Sub Show()
  On Error Resume Next
  
  If mfrmAddIn Is Nothing Then
    Set mfrmAddIn = New frmAddIn
  End If
  
  Load mfrmAddIn
  Set mfrmAddIn.VBInstance = VBInstance
  Set mfrmAddIn.Connect = Me
End Sub

'this method adds the Add-In to VB
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

  'this method adds the Add-In to VB
  'save the vb instance

  Set VBInstance = Application
  'this is a good place to set a breakpoint and
  'test various addin objects, properties and methods
  If ConnectMode = ext_cm_External Then
    'Used by the wizard toolbar to start this wizard
    Me.Show
   Else
    If AddInMenuAvailable Then
      Set mcbMenuCommandBarTopMenu = VBInstance.CommandBars("Add-Ins").Controls.Add(msoControlPopup)
      mcbMenuCommandBarTopMenu.Caption = "vbAutoErrorHandler"
      '
      CreateSubMenuItem mcbMenuCommandBarTopMenu, mcbMenuCommandBarAddErrorTraps, "Add Error Handling"
      'sink the event
      Set MenuHandlerAddErrorTraps = VBInstance.Events.CommandBarEvents(mcbMenuCommandBarAddErrorTraps)
      '
      CreateSubMenuItem mcbMenuCommandBarTopMenu, mcbMenuCommandBarAddLineNumbers, "Add Line Numbers"
      'sink the event
      Set MenuHandlerAddLineNumbers = VBInstance.Events.CommandBarEvents(mcbMenuCommandBarAddLineNumbers)
      '
      CreateSubMenuItem mcbMenuCommandBarTopMenu, mcbMenuCommandBarAdd2All, "Add Both To Entire Project"
      'sink the event
      Set MenuHandlerAdd2All = VBInstance.Events.CommandBarEvents(mcbMenuCommandBarAdd2All)
      '
      CreateSubMenuItem mcbMenuCommandBarTopMenu, mcbMenuCommandBarRemoveErrorTraps, "Remove Error Handling"
      'sink the event
      Set MenuHandlerRemoveErrorTraps = VBInstance.Events.CommandBarEvents(mcbMenuCommandBarRemoveErrorTraps)
      '
      CreateSubMenuItem mcbMenuCommandBarTopMenu, mcbMenuCommandBarRemoveLineNumbers, "Remove Line Numbers"
      'sink the event
      Set MenuHandlerRemoveLineNumbers = VBInstance.Events.CommandBarEvents(mcbMenuCommandBarRemoveLineNumbers)
      '
      CreateSubMenuItem mcbMenuCommandBarTopMenu, mcbMenuCommandBarRemoveFromAll, "Remove Both From Entire Project"
      'sink the event
      Set MenuHandlerRemoveFromAll = VBInstance.Events.CommandBarEvents(mcbMenuCommandBarRemoveFromAll)
      '
      CreateSubMenuItem mcbMenuCommandBarTopMenu, mcbMenuCommandBarEditTemplate, "Edit Template"
      'sink the event
      Set MenuHandlerEditTemplate = VBInstance.Events.CommandBarEvents(mcbMenuCommandBarEditTemplate)
    End If
  End If
End Sub

Private Sub CreateSubMenuItem(HeadMenu As Office.CommandBarPopup, SubMenuItem As Office.CommandBarControl, SMCaption As String)

  Set SubMenuItem = HeadMenu.Controls.Add(, , , , False)
  SubMenuItem.Caption = SMCaption

End Sub


Public Function AddInMenuAvailable() As Boolean

  AddInMenuAvailable = Not (VBInstance.CommandBars("Add-Ins") Is Nothing)
  If Not AddInMenuAvailable Then
    MsgBox "'Add-Ins' Menu is unavailable.", vbCritical
  End If

End Function

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
  On Error Resume Next
  
  'delete the command bar entry
  mcbMenuCommandBarAddErrorTraps.Delete
  mcbMenuCommandBarAddLineNumbers.Delete
  mcbMenuCommandBarRemoveLineNumbers.Delete
  mcbMenuCommandBarRemoveErrorTraps.Delete
  mcbMenuCommandBarEditTemplate.Delete
  mcbMenuCommandBarRemoveFromAll.Delete
  mcbMenuCommandBarAdd2All.Delete
  mcbMenuCommandBarTopMenu.Delete
  
  Unload mfrmAddIn
  Set mfrmAddIn = Nothing
End Sub

Private Sub MenuHandlerAdd2All_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Me.Show
  mfrmAddIn.funcAddErrorHandlersToAll
End Sub

'this event fires when the menu is clicked in the IDE
Private Sub MenuHandlerAddErrorTraps_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Me.Show
  mfrmAddIn.funcAddErrorHandlers
End Sub
Private Sub MenuHandlerAddLineNumbers_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Me.Show
  mfrmAddIn.funcAddLineNumbers
End Sub
Private Sub MenuHandlerRemoveErrorTraps_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Me.Show
  mfrmAddIn.funcRemoveErrorHandlers
End Sub

Private Sub MenuHandlerRemoveFromAll_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Me.Show
  mfrmAddIn.funcRemoveErrorHandlersFromAll
End Sub

Private Sub MenuHandlerRemoveLineNumbers_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Me.Show
  mfrmAddIn.funcRemoveLineNumbers
End Sub
Private Sub MenuHandleredittemplate_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
  Me.Show
  mfrmAddIn.funcEditTemplate
End Sub
