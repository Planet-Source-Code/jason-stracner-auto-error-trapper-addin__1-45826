VERSION 5.00
Begin VB.Form frmAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My Add In"
   ClientHeight    =   3195
   ClientLeft      =   2175
   ClientTop       =   1935
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by: Jason Stracner (www.lr-tc.com, lrtc.blogspot.com)
Option Explicit
Public VBInstance As VBIDE.VBE
Public Connect As Connect
Const sRemovalTag As String = "'< Added by vbAutoErrorHandler (do not remove this tag)>"
Dim sCodeToAddToTop As String
Dim sCodeToAddToBottom As String

Public Function funcHandleError(ByVal sModule As String, ByVal sProcedure As String, ByVal oErr As ErrObject, Optional sExtraInformation As String) As Long
  Dim Msg As String
  
  Msg = oErr.Source & " caused error '" & _
          oErr.Description & "' (" & _
          oErr.Number & ")" & vbCrLf & _
          "in module " & sModule & _
          " procedure " & sProcedure & _
          ", line " & Erl & "."
  If sExtraInformation <> "" Then
    Msg = Msg & vbNewLine & sExtraInformation
  End If
  funcHandleError = MsgBox( _
          Msg, vbAbortRetryIgnore + vbMsgBoxHelpButton + vbCritical, _
          "What to you want to do?", _
          oErr.HelpFile, _
          oErr.HelpContext)
  oErr.Clear
  Err.Clear
End Function

Private Sub CancelButton_Click()
  Connect.Hide
End Sub

Function funcLoadTemplate()
  On Error GoTo Error_handler_routine
  Dim iFreeFile As Long
  Dim sFileContent As String
  Dim sTopTag As String
  Dim sBottomTag As String
  Dim sPathToIniFile As String
  Dim sLinesOfFile() As String
  Dim iEachLine As Long
  
  Call subCheckIniFile
  iFreeFile = FreeFile

  sPathToIniFile = VBInstance.ActiveVBProject.FileName
  sPathToIniFile = Mid(sPathToIniFile, 1, InStrRev(sPathToIniFile, "\"))
  sPathToIniFile = sPathToIniFile & "vbAutoErrorHandler.ini"
      
  If FileExists(sPathToIniFile) Then
    Open sPathToIniFile For Input As #iFreeFile
    sFileContent = StrConv(InputB(LOF(iFreeFile), iFreeFile), vbUnicode)
    Close #iFreeFile
  Else
    MsgBox "Error opening " & sPathToIniFile, vbCritical, App.Title
  End If
  
  'remove comment lines
  sLinesOfFile = Split(sFileContent, vbNewLine)
  sFileContent = ""
  For iEachLine = LBound(sLinesOfFile) To UBound(sLinesOfFile)
    If Strings.Left(sLinesOfFile(iEachLine), 1) <> "#" Then
      sFileContent = sFileContent & sLinesOfFile(iEachLine) & vbNewLine
    End If
  Next iEachLine
  
  sTopTag = sFileContent
  sTopTag = Mid(sTopTag, InStr(sTopTag, "<add-to-top>") + Len("<add-to-top>"))
  sTopTag = Mid(sTopTag, 1, InStr(sTopTag, "</add-to-top>") - 1)
  sTopTag = funcRemoveNewlinesFromBeginingAndEnd(sTopTag)
  
  sBottomTag = sFileContent
  sBottomTag = Mid(sBottomTag, InStr(sBottomTag, "<add-to-bottom>") + Len("<add-to-bottom>"))
  sBottomTag = Mid(sBottomTag, 1, InStr(sBottomTag, "</add-to-bottom>") - 1)
  sBottomTag = funcRemoveNewlinesFromBeginingAndEnd(sBottomTag)
  
  'add removal tag to sTopTag
  sLinesOfFile = Split(sTopTag, vbNewLine)
  sTopTag = ""
  For iEachLine = LBound(sLinesOfFile) To UBound(sLinesOfFile)
    sTopTag = sTopTag & funcPadStringWithSpace(sLinesOfFile(iEachLine), 100) & " " & sRemovalTag & vbNewLine
  Next iEachLine
  sTopTag = funcRemoveNewlinesFromBeginingAndEnd(sTopTag)
  
  'add removal tag to sBottomTag
  sLinesOfFile = Split(sBottomTag, vbNewLine)
  sBottomTag = ""
  For iEachLine = LBound(sLinesOfFile) To UBound(sLinesOfFile)
    sBottomTag = sBottomTag & funcPadStringWithSpace(sLinesOfFile(iEachLine), 100) & " " & sRemovalTag & vbNewLine
  Next iEachLine
  sBottomTag = funcRemoveNewlinesFromBeginingAndEnd(sBottomTag)
  
  sCodeToAddToTop = sTopTag
  sCodeToAddToBottom = sBottomTag

  Err.Clear
Error_handler_routine:
  If Err <> 0 Then
    Select Case funcHandleError("frmAddIn", "funcLoadTemplate", Err)
      Case vbRetry: Resume
      Case vbIgnore: Resume Next
    End Select
  End If
End Function

Function funcRemoveNewlinesFromBeginingAndEnd(ByVal sInput As String) As String
  Dim sOutput As String
  
  sOutput = sInput
  Do
    If InStr(sOutput, vbNewLine) Then
      If Split(sOutput, vbNewLine)(0) = "" Then
        'remove first newline
        sOutput = Replace(sOutput, vbNewLine, "", 1, 1)
      Else
        Exit Do
      End If
    Else
      Exit Do
    End If
  Loop
  
  'take them off the end
  Do
    If InStr(sOutput, vbNewLine) Then
      If InStrRev(sOutput, vbNewLine) = Len(sOutput) - 1 Then
        'take one off the end
        sOutput = Mid(sOutput, 1, Len(sOutput) - 2)
      Else
        Exit Do
      End If
    Else
      Exit Do
    End If
  Loop
  
  funcRemoveNewlinesFromBeginingAndEnd = sOutput
End Function

Friend Function funcEditTemplate()
  Dim sPathToIniFile As String
  Dim sFileContent As String
  
  sPathToIniFile = VBInstance.ActiveVBProject.FileName
  sPathToIniFile = Mid(sPathToIniFile, 1, InStrRev(sPathToIniFile, "\"))
  sPathToIniFile = sPathToIniFile & "vbAutoErrorHandler.ini"
  
  If FileExists(sPathToIniFile) Then
    Shell "notepad " & """" & sPathToIniFile & """", vbNormalFocus
  Else
    Call subCheckIniFile
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    Shell "notepad " & """" & sPathToIniFile & """", vbNormalFocus
  End If
End Function

Sub subCheckIniFile()
  Dim sPathToIniFile As String
  Dim sFileContent As String
  
  sPathToIniFile = VBInstance.ActiveVBProject.FileName
  sPathToIniFile = Mid(sPathToIniFile, 1, InStrRev(sPathToIniFile, "\"))
  sPathToIniFile = sPathToIniFile & "vbAutoErrorHandler.ini"
  
  If Not FileExists(sPathToIniFile) Then
    sFileContent = ""
    sFileContent = sFileContent & "# Instructions:" & vbNewLine
    sFileContent = sFileContent & "#   You can edit the code inside the <add-to-top> and <add-to-bottom>" & vbNewLine
    sFileContent = sFileContent & "# tags to customize what code gets inserted into the top and bottom" & vbNewLine
    sFileContent = sFileContent & "# of your methods.  The tags <module-name> and <method-name> will " & vbNewLine
    sFileContent = sFileContent & "# automatically  be replaced when they are inserted into the code." & vbNewLine
    sFileContent = sFileContent & "# " & vbNewLine
    sFileContent = sFileContent & "#" & vbNewLine
    sFileContent = sFileContent & "# 'Here is a suggestion for the function you can use for" & vbNewLine
    sFileContent = sFileContent & "# 'the funcHandleError code.  You can put this code in " & vbNewLine
    sFileContent = sFileContent & "# 'a module so that all errors will get sent to this function." & vbNewLine
    sFileContent = sFileContent & "#" & vbNewLine
    sFileContent = sFileContent & "# Public Function funcHandleError(ByVal sModule As String, ByVal sProcedure As String, ByVal oErr As ErrObject) As Long" & vbNewLine
    sFileContent = sFileContent & "#   Dim Msg As String" & vbNewLine
    sFileContent = sFileContent & "#   " & vbNewLine
    sFileContent = sFileContent & "#   Msg = oErr.Source & "" caused error '"" & _" & vbNewLine
    sFileContent = sFileContent & "#           oErr.Description & ""' ("" & _" & vbNewLine
    sFileContent = sFileContent & "#           oErr.Number & "")"" & vbCrLf & _" & vbNewLine
    sFileContent = sFileContent & "#           ""in module "" & sModule & _" & vbNewLine
    sFileContent = sFileContent & "#           "" procedure "" & sProcedure & _" & vbNewLine
    sFileContent = sFileContent & "#           "", line "" & Erl & "".""" & vbNewLine
    sFileContent = sFileContent & "#   funcHandleError = MsgBox( _" & vbNewLine
    sFileContent = sFileContent & "#           Msg, vbAbortRetryIgnore + vbMsgBoxHelpButton + vbCritical, _" & vbNewLine
    sFileContent = sFileContent & "#           ""What to you want to do?"", _" & vbNewLine
    sFileContent = sFileContent & "#           oErr.HelpFile, _" & vbNewLine
    sFileContent = sFileContent & "#           oErr.HelpContext)" & vbNewLine
    sFileContent = sFileContent & "# End Function" & vbNewLine
    sFileContent = sFileContent & "" & vbNewLine
    sFileContent = sFileContent & "" & vbNewLine
    sFileContent = sFileContent & "<add-to-top>" & vbNewLine
    sFileContent = sFileContent & "  On Error GoTo Error_handler_routine_for_<method-name>_in_<module-name> " & vbNewLine
    sFileContent = sFileContent & "</add-to-top>" & vbNewLine
    sFileContent = sFileContent & "" & vbNewLine
    sFileContent = sFileContent & "<add-to-bottom>" & vbNewLine
    sFileContent = sFileContent & "  Err.Clear " & vbNewLine
    sFileContent = sFileContent & "Error_handler_routine_for_<method-name>_in_<module-name>: " & vbNewLine
    sFileContent = sFileContent & "  If Err <> 0 Then " & vbNewLine
    sFileContent = sFileContent & "    Select Case funcHandleError(""<module-name>"", ""<method-name>"", Err) " & vbNewLine
    sFileContent = sFileContent & "      Case vbRetry: Resume " & vbNewLine
    sFileContent = sFileContent & "      Case vbIgnore: Resume Next " & vbNewLine
    sFileContent = sFileContent & "    End Select " & vbNewLine
    sFileContent = sFileContent & "  End If" & vbNewLine
    sFileContent = sFileContent & "</add-to-bottom>" & vbNewLine
    Call subWriteIniFile(sFileContent, sPathToIniFile)
  End If
End Sub

Public Sub subWriteIniFile(ByVal sContent As String, ByVal sFilePath As String)
  Dim iFreeFile As Long
       
  On Error Resume Next
  iFreeFile = FreeFile
  Open sFilePath For Output As #iFreeFile
  Print #iFreeFile, sContent
  Close #iFreeFile
  If Err Then
    MsgBox "Error writting file " & sFilePath, vbCritical, App.Title
  Else
    Do
      DoEvents
    Loop Until FileExists(sFilePath)
  End If
End Sub

Public Function FileExists(ByVal strFileName As String) As Boolean
  Dim intLen As Integer
  On Error Resume Next

  If strFileName$ <> "" Then
    intLen% = Len(Dir$(strFileName$))
    If intLen = 0 Then
      intLen% = Len(Dir$(strFileName$, vbDirectory))
    End If
    FileExists = (Not Err And intLen% > 0)
  Else
    FileExists = False
  End If
End Function

Friend Function funcRemoveErrorHandlersFromAll()
  Dim oCodeModule As CodeModule
  Dim iEachCodeModule As Long
  
  For iEachCodeModule = 1 To VBInstance.ActiveVBProject.VBComponents.Count
    VBInstance.ActiveVBProject.VBComponents(iEachCodeModule).CodeModule.CodePane.Show
    Set oCodeModule = Nothing
    Set oCodeModule = VBInstance.CodePanes(iEachCodeModule).CodeModule
    Call funcRemoveLineNumbersFromCodeModule(oCodeModule)
    Call funcRemoveErrorHandlersFromCodeModule(oCodeModule)
    DoEvents
  Next iEachCodeModule
End Function

Function funcRemoveErrorHandlers()
  Dim oCodeModule As CodeModule
  
  Set oCodeModule = VBInstance.ActiveCodePane.CodeModule
  Call funcRemoveErrorHandlersFromCodeModule(oCodeModule)
End Function

Friend Function funcRemoveErrorHandlersFromCodeModule(oCodeMod As CodeModule)
      Dim iCurrentLine As Long, iStartLine As Long, iEndLine As Long
      Dim sProcName As String, fFindIt As Boolean
      Dim eProcKind As vbext_ProcKind
      Dim sExitStatement As String
      Dim sCurrentProc As String
      Dim iEachLine As Long
      Dim iEachMember As Long
      Dim iLastLineOfGeneralDeclarations As Long
      Dim iFirstLineOfThisFunction As Long
      Dim col_sMethodNames As New Collection
      Dim sTextOfLine As String

      If oCodeMod Is Nothing Then
          'they're not in the code pane.
          MsgBox "You must have a code window open to add error handling."
          Exit Function
      End If

      For iEachLine = oCodeMod.CountOfLines To 1 Step -1
          sTextOfLine = Trim(oCodeMod.Lines(iEachLine, 1))
          If Right(sTextOfLine, 1) = ">" Then
            If Len(sTextOfLine) >= Len(sRemovalTag) Then
              If Right(sTextOfLine, Len(sRemovalTag)) = sRemovalTag Then
                Call oCodeMod.DeleteLines(iEachLine)
              End If
            End If
          End If
      Next iEachLine
End Function

Friend Function funcAddErrorHandlersToAll()
  Dim oCodeModule As CodeModule
  Dim iEachCodeModule As Long
  
  Call funcLoadTemplate
  If Trim(sCodeToAddToTop) = "" Then
    MsgBox "Error in the ini file.  Couldn't find the <add-to-top> tags."
    Exit Function
  End If
  If Trim(sCodeToAddToBottom) = "" Then
    MsgBox "Error in the ini file.  Couldn't find the <add-to-bottom> tags."
    Exit Function
  End If
  For iEachCodeModule = 1 To VBInstance.ActiveVBProject.VBComponents.Count
    VBInstance.ActiveVBProject.VBComponents(iEachCodeModule).CodeModule.CodePane.Show
    Set oCodeModule = Nothing
    Set oCodeModule = VBInstance.ActiveVBProject.VBComponents(iEachCodeModule).CodeModule
    Call funcAddLineNumbersToCodeModule(oCodeModule)
    Call funcAddErrorHandlersToCodeModule(oCodeModule)
    DoEvents
  Next iEachCodeModule
End Function

'just current pane
Friend Function funcAddErrorHandlers()
  Dim oCodeModule As CodeModule
  
  Call funcLoadTemplate
  If Trim(sCodeToAddToTop) = "" Then
    MsgBox "Error in the ini file.  Couldn't find the <add-to-top> tags."
    Exit Function
  End If
  If Trim(sCodeToAddToBottom) = "" Then
    MsgBox "Error in the ini file.  Couldn't find the <add-to-bottom> tags."
    Exit Function
  End If
      
  Set oCodeModule = VBInstance.ActiveCodePane.CodeModule
  Call funcAddErrorHandlersToCodeModule(oCodeModule)
End Function

Function funcAddErrorHandlersToCodeModule(oCodeMod As CodeModule)
      On Error GoTo Error_handler_routine
      Dim iCurrentLine As Long
      Dim iStartLine As Long
      Dim iEndLine As Long
      Dim sProcName As String, fFindIt As Boolean
      Dim eProcKind As vbext_ProcKind
      Dim sExitStatement As String
      Dim sCurrentProc As String
      Dim iEachLine As Long
      Dim iEachMember As Long
      Dim iLastLineOfGeneralDeclarations As Long
      Dim iFirstLineOfThisFunction As Long
      Dim col_sMethodNames As New Collection
      Dim col_sMethodTypes As New Collection
      Dim sFirstLineToLookFor As String
      Dim sNewCodeToAddToTheTop As String
      Dim sNewCodeToAddToTheBottom As String
      
      If oCodeMod Is Nothing Then
          'they're not in the code pane.
          MsgBox "You must have a code window open to add error handling."
          Exit Function
      End If

      'find the last line of the declarations.
      iLastLineOfGeneralDeclarations = oCodeMod.CountOfLines
      For iEachLine = 1 To oCodeMod.CountOfLines
          sProcName = oCodeMod.ProcOfLine(iEachLine, eProcKind) '
          If Trim(sProcName) <> "" Then
            iEndLine = oCodeMod.ProcStartLine(sProcName, eProcKind)
            If iLastLineOfGeneralDeclarations > iEndLine Then
               iLastLineOfGeneralDeclarations = iEndLine
            End If
            If sProcName <> "funcHandleError" Then
              'don't add error handler code to our error handler function
              On Error Resume Next
              'gets an error when it tries to add a duplicate.
              col_sMethodNames.Add sProcName, sProcName '
              If Err Then
                Err.Clear
              Else
                col_sMethodTypes.Add eProcKind
              End If
              On Error GoTo Error_handler_routine
            End If
          End If
      Next iEachLine

      'add Option Explicit
      fFindIt = oCodeMod.Find("Option Explicit", 1, 1, iLastLineOfGeneralDeclarations, -1, True, False, False)
      If fFindIt = False Then
          oCodeMod.InsertLines 1, "Option Explicit"
      End If

      For iEachMember = 1 To col_sMethodNames.Count
          sProcName = col_sMethodNames(iEachMember)
          iStartLine = oCodeMod.ProcBodyLine(sProcName, col_sMethodTypes(iEachMember))
          iEndLine = oCodeMod.ProcCountLines(sProcName, col_sMethodTypes(iEachMember)) + iStartLine
          If InStr(sCodeToAddToTop, vbNewLine) Then
            sFirstLineToLookFor = Split(sCodeToAddToTop, vbNewLine)(0)
          Else
            sFirstLineToLookFor = sCodeToAddToTop
          End If
          sFirstLineToLookFor = Replace(sFirstLineToLookFor, sRemovalTag, "")
          sFirstLineToLookFor = Trim(sFirstLineToLookFor)
          'Don't insert code it the remove tag is persent or the first line of the <add-to-top> text is there.
          fFindIt = oCodeMod.Find(sFirstLineToLookFor, iStartLine, 1, iEndLine, -1, True, False, False) Or _
                    oCodeMod.Find(sRemovalTag, iStartLine, 1, iEndLine, -1, True, False, False)
          If fFindIt = False Then
            Do While Right(Trim(oCodeMod.Lines(iStartLine, 1)), 2) = " _"
                iStartLine = iStartLine + 1
            Loop
            'do tag replacements
            sNewCodeToAddToTheTop = sCodeToAddToTop
            sNewCodeToAddToTheTop = Replace(sNewCodeToAddToTheTop, "<module-name>", oCodeMod.Parent.Name)
            sNewCodeToAddToTheTop = Replace(sNewCodeToAddToTheTop, "<method-name>", sProcName)
            'Add "on error goto Error_handler_routine"
            oCodeMod.InsertLines iStartLine + 1, sNewCodeToAddToTheTop
            iStartLine = oCodeMod.ProcStartLine(sProcName, col_sMethodTypes(iEachMember))
            iStartLine = iStartLine + oCodeMod.ProcCountLines(sProcName, col_sMethodTypes(iEachMember))
            Do
              If iStartLine <= 0 Then
                MsgBox "Error: Couldn't find end of method named " & sProcName & " in module " & oCodeMod.Parent.Name & ".", vbCritical, App.Title
                Exit Function
              End If
              If funcDetectEndOfMethod(Trim(oCodeMod.Lines(iStartLine, 1))) Then
                Exit Do
              Else
                iStartLine = iStartLine - 1
              End If
            Loop
            'do tag replacements
            sNewCodeToAddToTheBottom = sCodeToAddToBottom
            sNewCodeToAddToTheBottom = Replace(sNewCodeToAddToTheBottom, "<module-name>", oCodeMod.Parent.Name)
            sNewCodeToAddToTheBottom = Replace(sNewCodeToAddToTheBottom, "<method-name>", sProcName)
            'Insert the line "Error_handler_routine: ... " at the bottom
            Call oCodeMod.InsertLines(iStartLine, sNewCodeToAddToTheBottom)
          End If
      Next iEachMember
    Err.Clear
Error_handler_routine:
    If Err <> 0 Then
      Select Case funcHandleError("frmAddIn", "funcAddErrorHandlers", Err)
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
      End Select
    End If
End Function

Function funcPadStringWithSpace(sInput As String, iFinishedLength As Long) As String
  Dim sOutput As String
  
  If Len(sInput) = iFinishedLength Then
    sOutput = sInput
  ElseIf Len(sInput) < iFinishedLength Then
    sOutput = sInput & Space(iFinishedLength - Len(sInput))
  ElseIf Len(sInput) > iFinishedLength Then
    sOutput = sInput
  End If
  
  funcPadStringWithSpace = sOutput
End Function

Function funcAddLineNumbers()
  Dim oCodeModule As CodeModule
  
  Set oCodeModule = VBInstance.ActiveCodePane.CodeModule
  Call funcAddLineNumbersToCodeModule(oCodeModule)
End Function

Friend Function funcAddLineNumbersToCodeModule(oCodeMod As CodeModule)
      Dim iStartLine As Long
      Dim iEndLine As Long
      Dim sProcName As String
      Dim eProcKind As vbext_ProcKind
      Dim sTextOfLine As String
      Dim iEachLine As Long
      Dim iEachMember As Long
      Dim col_sMethodNames As New Collection
      Dim col_sMethodTypes As New Collection
      Dim col_sMethodStartLine As New Collection
      Dim col_sMethodEndLine As New Collection
      Dim iEachLineOfMethod As Long
      Dim iLineNumber As Long
      Dim sFirstWordOfLine As String
      Dim sLastWordOfPreviousLine As String
      Dim sPreviousLine As String
    
      If oCodeMod Is Nothing Then
          'they're not in the code pane.
          MsgBox "You must have a code window open to add error handling."
          Exit Function
      End If

      For iEachLine = 1 To oCodeMod.CountOfLines
          sProcName = oCodeMod.ProcOfLine(iEachLine, eProcKind) '
          If sProcName <> "" Then
            iStartLine = 0
            iEndLine = 0
            iStartLine = oCodeMod.ProcBodyLine(sProcName, eProcKind)
            iEndLine = oCodeMod.ProcCountLines(sProcName, eProcKind) + oCodeMod.ProcStartLine(sProcName, eProcKind)
            col_sMethodNames.Add sProcName
            col_sMethodTypes.Add eProcKind
            col_sMethodStartLine.Add iStartLine
            col_sMethodEndLine.Add iEndLine
            iEachLine = iEndLine
          End If
      Next iEachLine

      For iEachMember = 1 To col_sMethodNames.Count
          For iEachLineOfMethod = col_sMethodStartLine(iEachMember) + 1 To col_sMethodEndLine(iEachMember) - 1
            sTextOfLine = oCodeMod.Lines(iEachLineOfMethod, 1)
            If iStartLine <> 1 Then
              sPreviousLine = oCodeMod.Lines(iEachLineOfMethod - 1, 1)
              If Trim(sPreviousLine) <> "" Then
                sLastWordOfPreviousLine = Split(sPreviousLine, " ")(UBound(Split(sPreviousLine, " ")))
              Else
                sLastWordOfPreviousLine = ""
              End If
            Else
              sLastWordOfPreviousLine = ""
            End If
            If Len(Trim(sTextOfLine)) > 0 Then
              sFirstWordOfLine = Split(Trim(sTextOfLine), " ")(0)
              If sFirstWordOfLine <> "Case" And _
                      sFirstWordOfLine <> "On" And _
                      sFirstWordOfLine <> "End" And _
                      IsNumeric(sFirstWordOfLine) = False And _
                      sLastWordOfPreviousLine <> "_" And _
                      funcDetectEndOfMethod(sTextOfLine) = False Then
                Call oCodeMod.ReplaceLine(iEachLineOfMethod, CStr(iEachLineOfMethod) & "  " & sTextOfLine)
              End If
            End If
          Next iEachLineOfMethod
      Next iEachMember
End Function

Function funcDetectEndOfMethod(sLine As String) As Boolean
  Dim sCheckForEnd As String
  Dim iEachNumber As Long
  
  sCheckForEnd = sLine
  For iEachNumber = 0 To 9
    'remove all numbers
    sCheckForEnd = Replace(sCheckForEnd, CStr(iEachNumber), "")
  Next iEachNumber
  sCheckForEnd = Trim(sCheckForEnd)
  
  If InStr(sCheckForEnd, "End Function") = 1 Or _
            InStr(sCheckForEnd, "End Property") = 1 Or _
            InStr(sCheckForEnd, "End Sub") = 1 Then
    funcDetectEndOfMethod = True
  Else
    funcDetectEndOfMethod = False
  End If
End Function

Function funcRemoveLineNumbers()
  Dim oCodeModule As CodeModule
  
  Set oCodeModule = VBInstance.ActiveCodePane.CodeModule
  Call funcRemoveLineNumbersFromCodeModule(oCodeModule)
End Function

Friend Function funcRemoveLineNumbersFromCodeModule(oCodeMod As CodeModule)
      Dim iStartLine As Long
      Dim iEndLine As Long
      Dim sProcName As String
      Dim eProcKind As vbext_ProcKind
      Dim sTextOfLine As String
      Dim iEachLine As Long
      Dim iEachMember As Long
      Dim col_sMethodNames As New Collection
      Dim col_sMethodTypes As New Collection
      Dim col_sMethodStartLine As New Collection
      Dim col_sMethodEndLine As New Collection
      Dim iEachLineOfMethod As Long
      Dim iLineNumber As Long
      Dim sFirstWordOfLine As String
      Dim sLastWordOfPreviousLine As String
      Dim sPreviousLine As String

      If oCodeMod Is Nothing Then
          'they're not in the code pane.
          MsgBox "You must have a code window open to add error handling."
          Exit Function
      End If

      For iEachLine = 1 To oCodeMod.CountOfLines
          sProcName = oCodeMod.ProcOfLine(iEachLine, eProcKind) '
          If sProcName <> "" Then
            iStartLine = 0
            iEndLine = 0
            iStartLine = oCodeMod.ProcBodyLine(sProcName, eProcKind)
            iEndLine = oCodeMod.ProcCountLines(sProcName, eProcKind) + iStartLine
            col_sMethodNames.Add sProcName
            col_sMethodTypes.Add eProcKind
            col_sMethodStartLine.Add iStartLine
            col_sMethodEndLine.Add iEndLine
            iEachLine = iEndLine
          End If
      Next iEachLine
        
      For iEachMember = 1 To col_sMethodNames.Count
          For iEachLineOfMethod = col_sMethodStartLine(iEachMember) + 1 To col_sMethodEndLine(iEachMember)
            sPreviousLine = ""
            sLastWordOfPreviousLine = ""
            sFirstWordOfLine = ""
            sTextOfLine = oCodeMod.Lines(iEachLineOfMethod, 1)
            If iStartLine <> 1 Then
              sPreviousLine = oCodeMod.Lines(iEachLineOfMethod - 1, 1)
              If Trim(sPreviousLine) <> "" Then
                sLastWordOfPreviousLine = Split(sPreviousLine, " ")(UBound(Split(sPreviousLine, " ")))
              End If
            End If
            If Len(Trim(sTextOfLine)) > 0 Then
              sFirstWordOfLine = Split(sTextOfLine, " ")(0)
              If sLastWordOfPreviousLine <> "_" And IsNumeric(sFirstWordOfLine) Then
                Call oCodeMod.ReplaceLine(iEachLineOfMethod, Replace(sTextOfLine, sFirstWordOfLine & "  ", "", 1, 1))
                If InStr(oCodeMod.Lines(iEachLineOfMethod, 1), sFirstWordOfLine & " ") Then
                  Call oCodeMod.ReplaceLine(iEachLineOfMethod, Replace(sTextOfLine, sFirstWordOfLine & " ", "", 1, 1))
                End If
                If IsNumeric(sFirstWordOfLine) And oCodeMod.Lines(iEachLineOfMethod, 1) = sFirstWordOfLine Then
                  Call oCodeMod.ReplaceLine(iEachLineOfMethod, Replace(sTextOfLine, sFirstWordOfLine, "", 1, 1))
                End If
              End If
            End If
          Next iEachLineOfMethod
      Next iEachMember
End Function


