VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Class by Hustle 2002, I probably owe some inspiration to
'the AllAPI Network team as their tools have help me delve deeper
'into API programming!!
Private sIEVer          As String
Private sOSVer          As String
Private lTotalMEM       As Long
Private sScreenSze      As String
Private lHardDriveSze   As Long
Private aDiskSizes      As Variant
Private iErrNumber      As Integer
Private iPlatFormID     As Integer
Private aProcesses()    As Variant
Private iActiveProcs    As Integer

Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const MAX_PATH = 260
Private Const WIN9X_SYS_FOUND = 1
Private Const WINNT_SYS_FOUND = 2


Private Type VersionInfo
   cbSize               As Long
   dwMajorVersion       As Long
   dwMinorVersion       As Long
   dwBuildNumber        As Long
   dwPlatformID         As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformID        As Long
    szCSDVersion        As String * 128
End Type

Private Type MEMORYSTATUS
    dwLength            As Long
    dwMemoryLoad        As Long
    dwTotalPhys         As Long
    dwAvailPhys         As Long
    dwTotalPageFile     As Long
    dwAvailPageFile     As Long
    dwTotalVirtual      As Long
    dwAvailVirtual      As Long
End Type

Private Type RECT
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type


Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long ' This process
    th32DefaultHeapID As Long
    th32ModuleID As Long ' Associated exe
    cntThreads As Long
    th32ParentProcessID As Long ' This process's parent process
    pcPriClassBase As Long ' Base priority of process threads
    dwFlags As Long
    szExeFile As String * 260 ' MAX_PATH
End Type

Public Enum enumProcessInfo
    eProcName = 1
    eProcMemLoad = 2
End Enum

Public Enum enumDiskSpace
    eTotalDiskSpace = 1
    eFreeDiskSpace = 2
    eUsedDiskSpace = 3
End Enum

Private Declare Function DllGetVersion Lib "Shlwapi.dll" (dwVersion As VersionInfo) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetProcessMemoryInfo Lib "PSAPI.DLL" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcesses Lib "PSAPI.DLL" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Function LoadProcessArray()
Dim lRet                As Long
Dim lCB(2)              As Long
Dim lCBNeeded           As Long
Dim lProcIDs()          As Long
Dim lProcNum            As Long
Dim l                   As Long
Dim lProcHnd           As Long
Dim sModName            As String
Dim lRSize              As Long
Dim lModules(1 To 200)  As Long
Dim sProcessName        As String
Dim iDeliminator        As Integer
Dim thisPMC             As PROCESS_MEMORY_COUNTERS

    'Check platform type as no built in support yet for
    'using a non NT system, currently developing this functionality
    Select Case iPlatFormID
        Case WIN9X_SYS_FOUND
            MsgBox "This application will not run on a non Windows NT machine!" _
            & vbCrLf & "This program will now close!", vbCritical, "Error"
            End
        Case WINNT_SYS_FOUND
            
            'This should be 16, but the next loop increments it before
            'getting the proccess IDs
            lCB(1) = 8
            lCBNeeded = 96
            
            Do While lCB(1) <= lCBNeeded
                DoEvents
                lCB(1) = lCB(1) * 2
                ReDim lProcIDs(lCB(1) / 4) As Long
                'Get Process the IDs
                lRet = EnumProcesses(lProcIDs(1), lCB(1), lCBNeeded)
            Loop
            
            'Count the number of proccesses returned
            lProcNum = lCBNeeded / 4
            
            ReDim aProcesses(lProcNum, 2)
            iActiveProcs = 1
            For l = 1 To lProcNum
                
                'Retrieve the process handle
                lProcHnd = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lProcIDs(l))
                
                If lProcHnd <> 0 Then
                    lRet = EnumProcessModules(lProcHnd, lModules(1), 200, lCB(2))
                    
                    If lRet <> 0 Then
                        sModName = Space(MAX_PATH)
                        lRSize = 500
                        
                        'Get an array of the module handles for this process
                        'using the retrieved handle
                        lRet = GetModuleFileNameExA(lProcHnd, lModules(1), sModName, lRSize)
                        sProcessName = Left(sModName, lRet)
                        'Strip the returned application path string to
                        'the specific file in the process list
                        iDeliminator = InStrRev(sProcessName, "\")
                        sProcessName = Right(sProcessName, Len(sProcessName) - iDeliminator)
                        Debug.Print sProcessName
                        'Add the process name to the holding array object
                        aProcesses(iActiveProcs, 1) = sProcessName
                            
                        'Then we get the required site of memory allocated to this process
                        thisPMC.cb = LenB(thisPMC)
                        lRet = GetProcessMemoryInfo(lProcHnd, thisPMC, thisPMC.cb)
                        'assign the memory allocation to the holding array
                        aProcesses(iActiveProcs, 2) = CLng(thisPMC.WorkingSetSize / 1024)
                        'Count the proccesses added to the holding array
                        iActiveProcs = iActiveProcs + 1
                    End If
                End If
                'Close the handle to this process
                lRet = CloseHandle(lProcHnd)
            Next l
            Debug.Print "End Group"
    End Select
End Function
Public Function Initialise() As Boolean
Dim myVInfo     As VersionInfo
Dim sIEVerTemp  As String
Dim myOSInfo    As OSVERSIONINFO
Dim lRet        As Long
Dim myMEMStatus As MEMORYSTATUS
Dim rScreen     As RECT

    'Reset Error Handler
    iErrNumber = 0
    
    'Here we Load the fixed drive sizes into an array
    sLoadDriveSpaceArray
    
    'Here we record the desktopsize
    rScreen = DesktopSize()
    sScreenSze = rScreen.Right & "x" & rScreen.Bottom
    
    'Here we record the total memory size
    GlobalMemoryStatus myMEMStatus
    lTotalMEM = myMEMStatus.dwTotalPhys / 1024
    
    'Here we gather information on the OS
    myOSInfo.dwOSVersionInfoSize = Len(myOSInfo)
    lRet = GetVersionEx(myOSInfo)
    If lRet = 0 Then GoTo badlRet
    sOSVer = WhatOS(myOSInfo)
    
    'Here we gather the iE Version on this computer
    myVInfo.cbSize = Len(myVInfo)
    lRet = DllGetVersion(myVInfo)
    If lRet <> 0 Then GoTo badlRet
    
    sIEVerTemp = "Internet Explorer " & myVInfo.dwMajorVersion & "." & myVInfo.dwMinorVersion & "." & myVInfo.dwBuildNumber
    sIEVer = sIEVerTemp
    
    'Load Process array with active processes
    LoadProcessArray
    
    Initialise = True
    Exit Function

badlRet:
Initialise = False
iErrNumber = 2
Exit Function

End Function
Private Function WhatOS(objSysInfo As OSVERSIONINFO) As String
Dim iVer_Major  As Integer
Dim iVer_Minor  As Integer
Dim lBuild      As Long
Dim sOSVersion  As String

    iVer_Major = objSysInfo.dwMajorVersion
    iVer_Minor = objSysInfo.dwMinorVersion
    iPlatFormID = objSysInfo.dwPlatformID
    lBuild = objSysInfo.dwBuildNumber
    
    Select Case iPlatFormID
            Case WIN9X_SYS_FOUND 'This means it is a non-NT OS
            Select Case iVer_Major
                Case 4
                    Select Case iVer_Minor
                        Case 0
                            If lBuild = 950 Then
                                sOSVersion = "Windows 95 "
                            ElseIf lBuild > 950 And lBuild < 1080 Then
                                sOSVersion = "Windows 95 SP1 "
                            End If
                        Case Is < 10
                            If lBuild > 1080 Then sOSVersion = "Windows 95 SR2 "
                        Case 10
                            If lBuild = 1998 Then
                                sOSVersion = "Windows 98 "
                            ElseIf lBuild > 1998 And lBuild < 2183 Then
                                sOSVersion = "Windows 98 SP1 "
                            ElseIf lBuild >= 2183 Then
                                sOSVersion = "Windows 98 SE "
                            End If
                        Case 90
                            sOSVersion = "Windows ME "
                    End Select
                Case Else
                    'unknown windows Version
                    sOSVersion = "Windows Unknown Version"
            End Select
            Case WINNT_SYS_FOUND 'NT File System
                Select Case iVer_Major
                    Case 3
                        sOSVersion = "Windows NT 3.51 "
                    Case 4
                        sOSVersion = "Windows NT 4 "
                    Case 5
                        If iVer_Minor = 0 Then
                            sOSVersion = "Windows 2000 "
                        ElseIf iVer_Minor = 1 Then
                            sOSVersion = "Windows XP "
                        End If
                    Case Else
                            sOSVersion = "Windows Unknown Version"
                End Select
    End Select
    
    WhatOS = sOSVersion & " " & iVer_Major & "." & iVer_Minor & " Build (" & lBuild & ")"

End Function
Public Property Get sIEVersion() As String
    sIEVersion = sIEVer
End Property
Public Property Get sOSVersion() As String
    sOSVersion = sOSVer
End Property
Public Property Get lTotalMemory() As Long
    lTotalMemory = lTotalMEM
End Property
Public Property Get sDeskTopSize() As String
    sDeskTopSize = sScreenSze
End Property
Public Property Get lDiskSize() As Long
    lDiskSize = lHardDriveSze
End Property
Public Property Get iErrNo() As Integer
    iErrNo = iErrNumber
End Property
Public Property Get iProcessCount() As Integer
    iProcessCount = iActiveProcs
End Property
Public Function sProcessInfo(ProcIndex As Integer, ProcInfo As enumProcessInfo) As String
    sProcessInfo = aProcesses(ProcIndex, ProcInfo)
End Function
Public Property Get sErrDesc() As String
Dim strDesc     As String

    Select Case iErrNo
        Case 1
            strDesc = "Invalid Drive Letter Assignment"
        Case 2
            strDesc = "Class Failed to initialise"
        Case Else
            strDesc = "Unspecified Error has occured"
    End Select
    
    sErrDesc = strDesc

End Property
Private Function DesktopSize() As RECT
Dim DesktopHwnd As Long
Dim lReturn     As Long

    DesktopHwnd = GetDesktopWindow()
    lReturn = GetWindowRect(DesktopHwnd, DesktopSize)
    
End Function
Private Sub sLoadDriveSpaceArray()
Dim lFBytes     As Currency
Dim lTBytes     As Currency
Dim ltnbytes    As Currency
Dim aBytest(24, 3) As Variant
Dim lRet        As Long
Dim i           As Integer

    'Here we cycle through all of the drive letter assignments
    For i = 99 To 122
        lRet = GetDiskFreeSpaceEx(Chr(i) & ":\", lFBytes, lTBytes, ltnbytes)
        'If no drive is found we skip that drive
        If lRet <> 0 Then
            aBytest(i - 98, 1) = CLng((lTBytes / 1048576) * 10000)
            aBytest(i - 98, 2) = CLng((ltnbytes / 1048576) * 10000)
            aBytest(i - 98, 3) = CLng(aBytest(i - 98, 1) - aBytest(i - 98, 2))
        End If
    Next i
    
    aDiskSizes = aBytest

End Sub
Public Function GetDriveByLetter(sDrive As String, eSizeOf As enumDiskSpace) As Long
Dim iChar       As Integer

    iChar = Asc(LCase(sDrive))
    
    If iChar < 99 Or iChar > 122 Then
        iErrNumber = 1
        Exit Function
    End If
    
    GetDriveByLetter = aDiskSizes(iChar - 98, eSizeOf)

End Function
Function fMemoryInUse() As Long
Dim myMEMStatus As MEMORYSTATUS
    
    'Retrieve the memory in use from the API object
    GlobalMemoryStatus myMEMStatus
    fMemoryInUse = myMEMStatus.dwMemoryLoad

End Function

