VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text 2 Forlders"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5520
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   5520
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chkSaveLog 
      Caption         =   "Save log"
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin MSComctlLib.StatusBar stbStatus 
      Height          =   270
      Left            =   0
      TabIndex        =   8
      Top             =   1560
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   476
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create Folders"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cdlSelect 
      Left            =   360
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkShowFolders 
      Caption         =   "Show folder after run"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton cmdTextFile 
      Caption         =   "..."
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox txtTextFile 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton cmdRootFolder 
      Caption         =   "..."
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtRootFolder 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Text File:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Root Folder:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCreate_Click()   '新建文件夹
    
    Dim strRootFolder, strFolderName, strFolderPath, strFileName As String
    Dim strLog As String
    Dim lngCount, lngSum, lngLine As Long
    
    lngLine = 0 '计数行号
    lngCount = 0    '计数建立成功的文件夹
    lngSum = 0  '计数文件夹总数
    
    stbStatus.SimpleText = "Ready"
    
    '容错处理
    If txtRootFolder.Text = "" Then
        stbStatus.SimpleText = "Please select root folder!"
        Exit Sub
    End If
    
    If txtTextFile.Text = "" Then
        stbStatus.SimpleText = "Please select text file!"
        Exit Sub
    End If
    
    '保存日志信息
    strLog = "Text 2 Folders Log File" & vbCrLf
    strLog = strLog & "Start time:" & Now & vbCrLf & "=========================================================" & vbCrLf

    strRootFolder = txtRootFolder.Text
    
    '按文本文档新建文件夹
    Open txtTextFile.Text For Input As #1
        Do While Not EOF(1)
            Line Input #1, strFolderName
            
            lngLine = lngLine + 1
            '去除文件名两侧空格
            strFolderName = Trim(strFolderName)
            
            '排除空白文件名
            If strFolderName <> "" Then
                lngSum = lngSum + 1
                
                '排除非法字符，容错
                If InStr(1, strFolderName, "\") = 0 And InStr(1, strFolderName, "/") = 0 And _
                    InStr(1, strFolderName, "?") = 0 And InStr(1, strFolderName, "*") = 0 And _
                    InStr(1, strFolderName, """") = 0 And InStr(1, strFolderName, ":") = 0 And _
                    InStr(1, strFolderName, "<") = 0 And InStr(1, strFolderName, ">") = 0 And _
                    InStr(1, strFolderName, "|") = 0 Then
                    
                    strFolderPath = strRootFolder & "\" & strFolderName
                    
                    If Len(strFolderPath) <= 137 Then   'VB创建文件夹路径有长度限制
                        If Dir(strFolderPath, vbDirectory) = "" Then
                            stbStatus.SimpleText = "Creating " & strFolderPath & " now..."
                            MkDir strRootFolder & "\" & strFolderName
                            lngCount = lngCount + 1
                        Else
                            stbStatus.SimpleText = strFolderPath & " is already exist..."
                            strLog = strLog & "Line:" & lngLine & " " & stbStatus.SimpleText & vbCrLf
                        End If
                    Else
                        stbStatus.SimpleText = strFolderPath & " is too long..."
                        strLog = strLog & "Line:" & lngLine & " " & stbStatus.SimpleText & vbCrLf
                    End If
                Else
                    stbStatus.SimpleText = strFolderName & " invalid folder name..."
                    strLog = strLog & "Line:" & lngLine & " " & stbStatus.SimpleText & vbCrLf
                End If
            Else
                strLog = strLog & "Line:" & lngLine & " " & "Null folder name..." & vbCrLf
            End If
        Loop
        
        stbStatus.SimpleText = lngCount & " folders create success and " & lngSum - lngCount & " folders create failed!"
        
        '保存日志信息
        strLog = strLog & stbStatus.SimpleText & vbCrLf
        strLog = strLog & "=========================================================" & vbCrLf & "End time:" & Now
            
        '保存日志文件
        If chkSaveLog.Value = 1 Then
            strFileName = Replace(Replace(Replace(Now, "/", ""), " ", ""), ":", "")
            Open strRootFolder & "\" & strFileName & ".log" For Output As #2
                Print #2, strLog
            Close #2
        End If
        
    Close #1
    
    '打开文件夹
    If chkShowFolders.Value = 1 Then
        Shell "explorer " & strRootFolder, 1
    End If
End Sub

Private Sub cmdRootFolder_Click()
    frmOpenFolder.Show 1
End Sub

Private Sub cmdTextFile_Click()
    cdlSelect.Filter = "文本文档(*.txt)|*.txt"
    cdlSelect.ShowOpen
    If cdlSelect.FileName <> "" Then
        txtTextFile.Text = cdlSelect.FileName
    End If
End Sub

Private Sub Form_Load()
    stbStatus.SimpleText = "Ready"
End Sub
