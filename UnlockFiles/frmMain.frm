VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB File Unlocker"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   120
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2775
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Process Path"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Locked File"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "PID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Handle"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Select Locked File"
      Height          =   375
      Left            =   2940
      TabIndex        =   0
      Top             =   3000
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
    On Error GoTo ErrorOut
    
    With Dlg
        .Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNFileMustExist + cdlOFNPathMustExist
        .Filter = "All Files (*.*)|*.*"
        .DialogTitle = "Select locked file"
        .CancelError = True
        .ShowOpen
        If .FileName <> "" Then
            
            DoEvents '// Give the CommonDialog window time to disappear
            lvwItems.ListItems.Clear '// Clear Listview
            strFile = .FileName '// Put the file path in a public variable
            
            If Not UnLockFile(.FileName, lvwItems) Then
                '// Something went wrong
                MsgBox "Something went wrong. The file may still be locked.  ", vbExclamation, "VB File Unlocker"
            Else
                '// Everything went OK, but no processes were closed,
                '// which means the file wasn't locked
                If lvwItems.ListItems.Count = 0 Then MsgBox "This file is not locked.  ", vbInformation, "VB File Unlocker"
            End If

        End If
    End With

Exit Sub
ErrorOut:
End Sub
