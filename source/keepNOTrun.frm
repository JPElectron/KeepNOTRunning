VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Keep NOT Running"
   ClientHeight    =   615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2655
   Icon            =   "keepNOTrun.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   615
   ScaleWidth      =   2655
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Detect As String
Dim Elapse As Double
Dim Idle As Double
Dim Launch As String
Dim Running As Boolean
Dim StartTime As Double

Dim Tick

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetLastInputInfo Lib "user32" (plii As Any) As Long
Private Type LASTINPUTINFO
   cbSize As Long
   dwTime As Long
End Type

Private Sub CheckProcess()
'On Error Resume Next
  Dim cb As Long
  Dim cbNeeded As Long
  Dim NumElements As Long
  Dim ProcessIDs() As Long
  Dim cbNeeded2 As Long
  Dim NumElements2 As Long
  Dim Modules(1 To 200) As Long
  Dim lRet As Long
  Dim ModuleName As String
  Dim nSize As Long
  Dim hProcess As Long
  Dim i As Long
  Dim lii As LASTINPUTINFO
         
    'Get the array containing the process id's for each process object
    cb = 8
    cbNeeded = 96
    Do While cb <= cbNeeded
        cb = cb * 2
        ReDim ProcessIDs(cb / 4) As Long
        lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
    Loop
         
    NumElements = cbNeeded / 4
    foundit = False
    For i = 1 To NumElements
      'Get a handle to the Process
         hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(i))
      'Got a Process handle
         If hProcess <> 0 Then
           'Get an array of the module handles for the specified process
             lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)
             
           'If the Module Array is retrieved, Get the ModuleFileName
            If lRet <> 0 Then
                ModuleName = Space(MAX_PATH)
                nSize = 500
                lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
                 
                If CBool(InStr(1, (Left(ModuleName, lRet)), m_strFilter, vbTextCompare)) Then
                    'AddListItem Left(ModuleName, lRet), ProcessIDs(i)
                    If LCase(Left(ModuleName, lRet)) Like LCase(Detect) = True Then
                        foundit = True
                    End If
                    
                End If
            End If
        End If
               
        'Close the handle to the process
        lRet = CloseHandle(hProcess)
    Next
    If foundit = False Then
        Running = False
        Debug.Print "Not Found"
    End If
    
    lii.cbSize = Len(lii)
    Call GetLastInputInfo(lii)
    deltaM = (GetTickCount() - lii.dwTime) / 1000 / 60
    deltaM2 = (GetTickCount() - StartTime) / 1000 / 60
    Debug.Print deltaM, deltaM2
    
    If foundit = True Then
        If Running = False Then
            'the program was just started
            Debug.Print "Started!"
            Running = True
            StartTime = GetTickCount()
            
        ElseIf deltaM >= Idle And deltaM2 >= Idle Then
            Debug.Print "Shutting down"
            Shell Launch, vbNormalNoFocus
            Running = False
            Open "keepnotrun.log" For Append As #1
            Print #1, Now & " - " & Launch & " executed."
            Close
            
        ElseIf deltaM < Idle Or deltaM2 < Idle Then
            Debug.Print "Waiting for idle time..."

        End If
    End If
    
End Sub


Private Sub Form_Load()

Waiting = False
On Error Resume Next
Tick = 0
Delay = 0
inifound = Dir("keepnotrun.ini")
If inifound = "" Then
    Open "keepnotrun.ini" For Output As #1
    
        Print #1, "; Keep NOT Running v" & App.Major & "." & App.Minor & " Configuration File"
        Print #1, " "
        Print #1, "; full path of executable to check"
        Print #1, "Detect="
        Print #1, " "
        Print #1, "; check the running processes list every X seconds"
        Print #1, "Interval=0.5"
        Print #1, " "
        Print #1, "; if the machine is idle for X minutes"
        Print #1, "Idle=20"
        Print #1, " "
        Print #1, "; full path of executable to be launched after idle"
        Print #1, "Launch="

    Close #1
    MsgBox "Created the keepnotrun.ini file, please configure before use.", vbOKOnly, "Error"
    Unload Me
    End
Else
    Open "keepnotrun.ini" For Input As #1
    Do While Not EOF(1)
    Input #1, temp
    inidata = Split(temp, "=")
        If UBound(inidata) > 0 Then
            'inidata(0) contains the identifier
            'inidata(1) contains the value, plus a comment
            bestdata = Split(inidata(1), "vbtab")
            'bestdata(0) contains the right thing
            
            Select Case inidata(0)
                Case "Detect"
                    Detect = bestdata(0)
                
                Case "Interval"
                    Elapse = bestdata(0)
                
                Case "Launch"
                    Launch = bestdata(0)
                
                Case "Idle"
                    Idle = bestdata(0)
                    
            End Select
        End If
    Loop
    Close
End If

'check to make sure every thing is right before starting the timer
If Launch <> "" Then
    If Detect <> "" Then
        If Elapse >= 0.5 Then
            If Launch Like "*" & App.EXEName & "*" = False Then
                If Detect Like "*" & App.EXEName & "*" = False Then
                    Timer1.Enabled = True
                Else
                    MsgBox "Launch= or Detect= cannot be set to " & App.EXEName, vbOKOnly, "Error in INI"
                    Unload Me
                End If
            Else
                MsgBox "Launch= or Detect= cannot be set to " & App.EXEName, vbOKOnly, "Error in INI"
                Unload Me
            End If
        Else
            MsgBox "Interval= is too short, it must be 0.5 seconds or greater.", vbOKOnly, "Error in INI"
            Unload Me
        End If
    Else
        MsgBox "Detect= must be set to the full path of executable to be checked.", vbOKOnly, "Error in INI"
        Unload Me
    End If
Else
    MsgBox "Launch= must be set to the full path of executable to be re-launched.", vbOKOnly, "Error in INI"
    Unload Me
End If
End Sub

Private Sub Timer1_Timer()
Tick = Tick + 0.1
'Debug.Print CurrentIdleTime
If Tick > Elapse Then
    Tick = 0
    CheckProcess
End If
End Sub


