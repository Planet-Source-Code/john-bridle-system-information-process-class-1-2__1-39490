VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmInfo 
   Caption         =   "System Information"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4770
   Icon            =   "frmInfo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrOne 
      Interval        =   500
      Left            =   120
      Top             =   4680
   End
   Begin VB.Frame frDriveInfo 
      Caption         =   "Hard Drive Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   4320
      Width           =   4575
      Begin VB.Frame frDrive 
         Caption         =   "Drive Letter"
         Height          =   735
         Left            =   3120
         TabIndex        =   10
         Top             =   120
         Width           =   1335
         Begin VB.ComboBox cboDrive 
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Label lblUsedSpace 
         Height          =   255
         Left            =   1680
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblUsedDisk 
         Caption         =   "Used Disk Space:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblFreeSpace 
         Height          =   255
         Left            =   1560
         TabIndex        =   15
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblFreeDisk 
         Caption         =   "Free Disk Space:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblTotalDisk 
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label LlblDiskTotal 
         Caption         =   "Total Size of Drive :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame frSystem 
      Caption         =   "System Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin MSComctlLib.ListView lstProcesses 
         Height          =   1815
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   3201
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Process Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Process Memory"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "CPU %"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ProgressBar pbMemLoad 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProcessNo 
         Height          =   255
         Left            =   2640
         TabIndex        =   23
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblProcesses 
         Caption         =   "Number of active Processes:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   2535
      End
      Begin VB.Label lblPercentage 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   21
         Top             =   1560
         Width           =   435
      End
      Begin VB.Label lblMemLoad 
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblMemLoadDesc 
         Caption         =   "Memory Load:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblMemory 
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblMemDesc 
         Caption         =   "System Memory:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblIEVersion 
         Height          =   255
         Left            =   2400
         TabIndex        =   6
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label lblIEDesc 
         Caption         =   "Internet Explorer Version:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lblDesktop 
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblDesktopDesc 
         Caption         =   "Desktop Resolution:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblOS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblOSDesc 
         Caption         =   "Operating System:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objSysInfo As clsSysInfo

Private Sub cboDrive_Click()
    
    'Update the Disk Values
    With Me
        .lblTotalDisk.Caption = objSysInfo.GetDriveByLetter(.cboDrive.Text, eTotalDiskSpace) & " Megabytes"
        .lblFreeSpace.Caption = objSysInfo.GetDriveByLetter(.cboDrive.Text, eFreeDiskSpace) & " Megabytes"
        .lblUsedSpace.Caption = objSysInfo.GetDriveByLetter(.cboDrive.Text, eUsedDiskSpace) & " Megabytes"
    End With
    
    'This handles the error property of the system information class
    If objSysInfo.iErrNo > 0 Then
        MsgBox objSysInfo.sErrDesc
    End If

End Sub

Private Sub Form_Load()
Dim i       As Integer
On Error GoTo err_handler
'Create a new instance of the System Information Class
Set objSysInfo = New clsSysInfo

    'This Process Loads the system information into the Class object
    If objSysInfo.Initialise Then
    
        With Me
            .lblIEVersion.Caption = objSysInfo.sIEVersion
            .lblOS.Caption = objSysInfo.sOSVersion
            .lblMemory.Caption = objSysInfo.lTotalMemory
            .lblDesktop.Caption = objSysInfo.sDeskTopSize
            'Here we populate our drive drop down list by
            'checking each individual drive in the class
            'has a size of greater than 0
            For i = 99 To 122
                If objSysInfo.GetDriveByLetter(Chr(i), eTotalDiskSpace) > 0 Then
                    .cboDrive.AddItem UCase(Chr(i))
                End If
            Next i
            'Select the first drive in the list
            .cboDrive.ListIndex = 0
            'Load Disk Values for the first selected drive in the list
            .lblTotalDisk.Caption = objSysInfo.GetDriveByLetter(.cboDrive.Text, eTotalDiskSpace) & " Megabytes"
            .lblFreeSpace.Caption = objSysInfo.GetDriveByLetter(.cboDrive.Text, eFreeDiskSpace) & " Megabytes"
            .lblUsedSpace.Caption = objSysInfo.GetDriveByLetter(.cboDrive.Text, eUsedDiskSpace) & " Megabytes"
            .lblProcessNo.Caption = objSysInfo.iProcessCount
            'Call the function to load the list view object with
            'the process holding array
            LoadProccessInfo
        End With
        
        'This handles the error property of the system information class
        If objSysInfo.iErrNo > 0 Then
            GoTo err_handler
        End If
    Else
        GoTo err_handler
    End If

Exit Sub

err_handler:
If objSysInfo.iErrNo > 0 Then
    MsgBox objSysInfo.sErrDesc, vbCritical, "Application Error"
Else
    MsgBox Err.Number & " " & Err.Description, vbCritical, "Application Error"
End If

Exit Sub

End Sub

Private Sub Form_Terminate()
    'Destroy the class object before closing
    Set objSysInfo = Nothing
End Sub

Private Sub tmrOne_Timer()
Dim iTemp   As Integer
Dim lTemp   As Long

'The timer is used to display updates in the memory load
'and New or destroyed processes
With Me
    iTemp = objSysInfo.fMemoryInUse
    lTemp = objSysInfo.lTotalMemory / 100
    .lblMemLoad.Caption = lTemp * iTemp
    .pbMemLoad.Value = iTemp
    .lblPercentage.Caption = CStr(iTemp) & "%"
    .lblProcessNo.Caption = objSysInfo.iProcessCount
    LoadProccessInfo
End With

End Sub
Private Sub LoadProccessInfo()
Dim i               As Integer
Dim x               As Integer
Dim isInProcess     As Boolean
Dim sTemp           As String
On Error Resume Next

With Me
    'Reload the process array
    objSysInfo.LoadProcessArray
    'If the list is empty skip this next section
    'this is used the first time the form is loaded
    If .lstProcesses.ListItems.Count > 0 Then
        'Here we cycle backwards through the list view object removing processes
        'that have since closed or have been destroyed
        For x = .lstProcesses.ListItems.Count To 1 Step -1
            DoEvents
            For i = 1 To objSysInfo.iProcessCount
                sTemp = objSysInfo.sProcessInfo(i, eProcName)
                If sTemp = .lstProcesses.ListItems(x) Then
                    isInProcess = True
                End If
            Next i
            If isInProcess = False Then
                Debug.Print .lstProcesses.ListItems(x)
                .lstProcesses.ListItems.Remove (x)
            End If
            isInProcess = False
        Next x

        isInProcess = False
        'Now we cycle forwards through the list to check to see
        'if any new processes appear in our class object
        For i = 1 To objSysInfo.iProcessCount - 1
            For x = 1 To .lstProcesses.ListItems.Count
                 sTemp = objSysInfo.sProcessInfo(i, eProcName)
                If sTemp = .lstProcesses.ListItems(x) Then
                    isInProcess = True
                End If
            Next x
            If isInProcess = False Then
                Debug.Print sTemp
                .lstProcesses.ListItems.Add .lstProcesses.ListItems.Count + 1, , sTemp
                .lstProcesses.ListItems(.lstProcesses.ListItems.Count).ListSubItems.Add , , objSysInfo.sProcessInfo(i, eProcMemLoad) & "k"
                .lstProcesses.ListItems(.lstProcesses.ListItems.Count).ListSubItems.Add , , CInt(CLng(objSysInfo.sProcessInfo(i, eProcMemLoad)) / CLng(.lblMemory.Caption))
                'GoTo Restart
            End If
            isInProcess = False
        Next i
        
    Else
        'This loads the list view object when the application starts
        For i = 1 To objSysInfo.iProcessCount - 1
            .lstProcesses.ListItems.Add i, "Key" & CStr(i), objSysInfo.sProcessInfo(i, eProcName)
            .lstProcesses.ListItems(.lstProcesses.ListItems.Count).ListSubItems.Add , , objSysInfo.sProcessInfo(i, eProcMemLoad) & "k"
            .lstProcesses.ListItems(.lstProcesses.ListItems.Count).ListSubItems.Add , , CInt(CLng(objSysInfo.sProcessInfo(i, eProcMemLoad)) / CLng(.lblMemory.Caption))
        Next i
    End If
End With

End Sub
