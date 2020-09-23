VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multi-Ping"
   ClientHeight    =   4770
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7800
   Icon            =   "frmPing.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7800
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Results:"
      Height          =   3135
      Left            =   4080
      TabIndex        =   10
      Top             =   0
      Width           =   3495
      Begin VB.ListBox lstPing 
         Height          =   2595
         Left            =   120
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Hosts to Ping:"
      Height          =   3135
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3975
      Begin VB.CheckBox chkTo 
         Caption         =   "To:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "Clear List"
         Top             =   1560
         Width           =   615
      End
      Begin VB.CommandButton cmdDel 
         Appearance      =   0  'Flat
         Caption         =   "Rem"
         Height          =   375
         Left            =   120
         Picture         =   "frmPing.frx":0442
         TabIndex        =   15
         ToolTipText     =   "Remove Entry"
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         ToolTipText     =   "Add Entry"
         Top             =   480
         Width           =   495
      End
      Begin VB.ListBox lstAdr 
         Height          =   1425
         Left            =   840
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Go!"
         Default         =   -1  'True
         Height          =   375
         Left            =   2520
         Picture         =   "frmPing.frx":0544
         TabIndex        =   6
         Top             =   2640
         Width           =   975
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   840
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtLast 
         Height          =   285
         Left            =   840
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblNote 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Remember that from computer 0.0.0.1 to computer 0.0.1.1 it has go through 255 computers"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "From:"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   7575
      Begin MSComDlg.CommonDialog cdl 
         Left            =   360
         Top             =   360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ProgressBar Progress 
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   960
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Min             =   1e-4
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         Height          =   255
         Left            =   5160
         TabIndex        =   13
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "0%"
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Percentage done"
         Height          =   675
         Left            =   2280
         TabIndex        =   7
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSave 
         Caption         =   "Save Log As..."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Constants:
'Ignore them =)
Const IP_0_0_0_1 = 16777216
Const IP_0_0_1_0 = 65536
Const IP_0_1_0_0 = 256
Const IP_1_0_0_0 = 1

Private Sub chkTo_Click()
    'User enables Interval Search:
    If chkTo.Value = 0 Then
        txtLast.Enabled = False
    Else
        txtLast.Enabled = True
    End If
End Sub

Private Sub cmdAdd_Click()
    If IsAddress(txtIP) Then
        'Data is an IP adress:
        cmdAdd.Enabled = False
        
        If chkTo.Value = 1 Then
            
            If Trim(txtLast) = "" Then 'Trim is used to "clean" adjacent spaces
            'User forgot to insert an address let's just ignore the interval:
                If Not EntryExists(txtIP, lstAdr) Then
                    'Single address entry:
                    'There is no repeated entry, ok to proceed:
                    lstAdr.AddItem txtIP
                    'Check if Rem button is enabled
                    If cmdDel.Enabled = False Then cmdDel.Enabled = True
                End If
            
            ElseIf IsAddress(txtLast) Then 'User has inserted some value
                'Interval of Adresses search routine:
                Dim FirstAddr As Double, LastAddr As Double, IpAddr As String, CurAddr As Double
                'First and last address converted to long so they can be calculated later:
                FirstAddr = AddrToLong(txtIP)
                LastAddr = AddrToLong(txtLast)
                'Obviously first addr must be smaller
                If FirstAddr < LastAddr Then  'OK to proceed
                    For CurAddr = FirstAddr To LastAddr
                        'From first address to the last:
                        'Convert it to IP - string - so we can show it in list
                        IpAddr = LongToAddr(CurAddr)
                        If Not EntryExists(IpAddr, lstAdr) Then
                            'Assure there are no duplicates
                            lstAdr.AddItem IpAddr
                        End If
                        'we don't wan't to make our app to "freeze"
                        DoEvents
                    Next CurAddr
                End If
                'at least one entry was made, so let's check for Remove button:
                If cmdDel.Enabled = False Then cmdDel.Enabled = True
            End If
        
        'No interval search (not checked)
        ElseIf Not EntryExists(txtIP, lstAdr) Then
            'Single address entry:
            'No repeated entry, ok to proceed...
            lstAdr.AddItem txtIP
            'Check if Rem button is enabled
            If cmdDel.Enabled = False Then cmdDel.Enabled = True
        End If
    End If
    'Clean up process:
    txtIP = "": txtLast = ""
    cmdAdd.Enabled = True
End Sub

Private Sub cmdClear_Click()
    'Clear contents
    lstAdr.Clear
    'Check for delete button (there won't be nothin to delete)
    If cmdDel.Enabled = True Then cmdDel.Enabled = False
End Sub

Private Sub cmdDel_Click()
    With lstAdr
        'Assure error free:
        If .ListCount = 0 Then cmdDel.Enabled = False: Exit Sub
        'Assure there is somethin selected
        If .ListIndex > -1 Then
            If .Selected(.ListIndex) Then
                .RemoveItem .ListIndex
            End If
        End If
        'If for some reason it's still enabled and
        'there's nothin to delete, disable button:
        If .ListCount = 0 Then cmdDel.Enabled = False
    End With
End Sub

Private Sub cmdSearch_Click()
    Static Working As Boolean
    'Boolean flag to know if user canceled
    If Working Then Working = False: Exit Sub 'User pressed then stop working
    'Start working:
    Working = True
    'Enable cancel button
    cmdSearch.Caption = "Cancel"
    'Clear ping results
    lstPing.Clear
    Dim i As Long, Elapsed As Long, Remaining As Long, ret As Long
    'Initialize progress bar
    Progress.Min = 0
    Progress.Max = lstAdr.ListCount - 1
    
    'Start ping loop:
    For i = 0 To Progress.Max
        'Initialize clock:
        Elapsed = GetTickCount
        'Change progressbar:
        Progress = i
        'Check if user pressed cancel:
        If Not Working Then
            lstPing.AddItem "Operation aborted by user!"
            GoTo CleanUp
        End If
        'Status
        Label4 = "Pinging: " & lstAdr.List(i) & vbLf & "Hosts remaining: " & Progress.Max - Progress & vbLf & "Estimated time: " & Round((Progress.Max - Progress) * Remaining / 1000, 3) & " secs"
        'Refresh form, if ommited user can't cancel operation:
        DoEvents
        'ret->delay of ping
        'Ping host:
        ret = PingHostByAdress(lstAdr.List(i))
        If ret >= 0 Then
            lstPing.AddItem lstAdr.List(i) & ": " & ret & " ms"
            'DoEvents
        End If
        
        'This sets the delay of last action that will be used
        'for calculating remaining time:
        Remaining = GetTickCount - Elapsed
Next i
CleanUp:
    'Safest way of cleaning progress bar:
    Progress = Progress.Min
    'Reset button caption:
    cmdSearch.Caption = "Go!"
    'Work is done
    Working = False
    'Clean status label:
    Label4 = vbLf & vbLf & "Percentage Done"
End Sub

Private Sub Form_Load()
    'Check if Del button is going to be enabled:
    If lstAdr.ListCount = 0 Then cmdDel.Enabled = False
    'Clean status label:
    Label4 = vbLf & vbLf & "Percentage Done"
    'Check if last txtbox is going to be enabled:
    If chkTo.Value = 0 Then txtLast.Enabled = False
    'Refresh tooltip text:
    If lstAdr.ListCount = 0 Then lstAdr.ToolTipText = "No adresses to ping"
End Sub

Private Sub lstAdr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'To spare some resources let's do this only when it's needed:
    Static LastCount As Long
    'Check if value has changed so we can or not refresh tooltip text:
    If LastCount <> lstAdr.ListCount Then
        If lstAdr.ListCount = 1 Then
            'Only one address to ping:
            lstAdr.ToolTipText = lstAdr.ListCount & " address to ping"
        ElseIf lstAdr.ListCount = 0 Then
            'No addresses to ping:
            lstAdr.ToolTipText = "No addresses to ping"
        Else
            'More than one address to ping:
            lstAdr.ToolTipText = lstAdr.ListCount & " addresses to ping"
        End If
        'Refresh last value:
        LastCount = lstAdr.ListCount
    End If
End Sub

Private Sub lstAdr_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Enable drag and drop support:
    Dim Addr As String
    'Check the kind of incoming value
    Addr = data.GetData(vbCFText)
    'Check if is an address
    If IsAddress(Addr) Then
        'Check if it exists in list:
        If Not EntryExists(Addr, lstAddr) Then
            lstAdr.AddItem Addr
            If cmdDel.Enabled = False Then cmdDel.Enabled = True
        End If
    End If
End Sub

Function EntryExists(Entry As String, List As ListBox) As Boolean
    With List
        'Simple loop in matching each value of each list value with the disired
        Dim i As Integer, Last As Integer
        EntryExists = False
        Last = .ListCount - 1
        If Last < 0 Then Exit Function
        For i = 0 To Last
            If .List(i) = Entry Then
                EntryExists = True
                Exit Function
            End If
        Next i
    End With
End Function

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuSave_Click()
    On Error GoTo ErrHandler
    Dim file As Integer
    With cdl
        .DialogTitle = "Save Log As..."
        .Filter = "Log File (*.log) |*.log"
        .ShowSave
        If lstPing.ListCount = 0 Then GoTo ErrHandler
        file = FreeFile
        Open .FileName For Output As #file
            'put's the contents of listbox "txtPing" in disired file:
            Dim i As Integer
            For i = 0 To lstPing.ListCount - 1
                Print #file, lstPing.List(i)
            Next i
        Close #file
    End With
ErrHandler:
'user has pressed cancel
End Sub

Private Sub txtIP_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Enable dragdrop in textbox:
    Dim Addr As String
    Addr = data.GetData(vbCFText)
    If IsAddress(Addr) Then
        txtIP = Addr
    End If
End Sub

Private Sub txtLast_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Enable dragdrop in textbox:
    Dim Addr As String
    Addr = data.GetData(vbCFText)
    If IsAddress(Addr) Then
        txtLast = Addr
    End If
End Sub
