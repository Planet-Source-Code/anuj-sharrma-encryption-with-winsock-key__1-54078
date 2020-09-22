VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00D5DDDD&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEncryptSomeText 
      BackColor       =   &H00D5DDDD&
      Caption         =   "Encrypt Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6405
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2565
      Width           =   1455
   End
   Begin VB.CommandButton cmdEncryptaFile 
      BackColor       =   &H00D5DDDD&
      Caption         =   "Encrypt a File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D5DDDD&
      Height          =   1785
      Left            =   120
      TabIndex        =   11
      Top             =   645
      Width           =   7770
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00D5DDDD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   360
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1245
         Width           =   2250
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Close"
         Height          =   405
         Left            =   5415
         TabIndex        =   18
         Top             =   1230
         Width           =   1080
      End
      Begin VB.CommandButton cmdTextDecrypt 
         Caption         =   "&Decrypt"
         Height          =   405
         Left            =   6570
         TabIndex        =   17
         Top             =   780
         Width           =   1080
      End
      Begin VB.CommandButton cmdTextEncrypt 
         Caption         =   "&Encrypt"
         Height          =   405
         Left            =   5445
         TabIndex        =   16
         Top             =   315
         Width           =   1080
      End
      Begin VB.TextBox txtResult 
         BackColor       =   &H00D5DDDD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   360
         Left            =   1245
         TabIndex        =   13
         Top             =   795
         Width           =   4125
      End
      Begin VB.TextBox txtSourceText 
         BackColor       =   &H00D5DDDD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   360
         Left            =   1245
         TabIndex        =   12
         Top             =   315
         Width           =   4125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "System IP address (Password)"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   270
         Left            =   255
         TabIndex        =   20
         Top             =   1260
         Width           =   2700
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Result"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   270
         Left            =   495
         TabIndex        =   15
         Top             =   840
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Text"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   270
         Left            =   150
         TabIndex        =   14
         Top             =   360
         Width           =   915
      End
   End
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   495
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D5DDDD&
      Height          =   1785
      Left            =   120
      TabIndex        =   1
      Top             =   645
      Width           =   7770
      Begin VB.CommandButton cmdEncrypt 
         BackColor       =   &H00D5DDDD&
         Caption         =   "Encrypt"
         Height          =   420
         Left            =   4065
         MouseIcon       =   "FrmMain.frx":0442
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1275
         Width           =   1170
      End
      Begin VB.CommandButton cmdDecrypt 
         BackColor       =   &H00D5DDDD&
         Caption         =   "Decrypt"
         Height          =   420
         Left            =   5295
         MouseIcon       =   "FrmMain.frx":074C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1275
         Width           =   1170
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00D5DDDD&
         Caption         =   "Close"
         Height          =   420
         Left            =   6525
         MouseIcon       =   "FrmMain.frx":0A56
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1275
         Width           =   1170
      End
      Begin VB.CommandButton cmdFileToEncrypt 
         BackColor       =   &H00168B0A&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7335
         MouseIcon       =   "FrmMain.frx":0D60
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   345
      End
      Begin VB.TextBox txtSource 
         BackColor       =   &H00D5DDDD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   330
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   255
         Width           =   4695
      End
      Begin VB.TextBox txtDestination 
         BackColor       =   &H00D5DDDD&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   330
         Left            =   2625
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   630
         Width           =   4680
      End
      Begin VB.CommandButton cmdFileForEncrypt 
         BackColor       =   &H00168B0A&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7320
         MouseIcon       =   "FrmMain.frx":106A
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select a file to encrypt/decrypt"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   90
         TabIndex        =   4
         Top             =   270
         Width           =   2445
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select file destnation pathh"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   75
         TabIndex        =   3
         Top             =   690
         Width           =   2160
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6000
      Top             =   105
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Encryption/Decryption"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   330
      Left            =   2745
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Const sDefaultValueFirst As String = "ABCDEFGHIJKLMNOPQRSTVUWXYZ_1234567890qwertyuiopasd!@#$%^&*(),. ~`-=\?/'""fghjklzxcvbnm"
Private Const sDefaultValueSecond As String = "IWEHJKTLZVOPFG_1234567890qwerBNMQRYUASDXCfghjklzxc ~`-=\?/'""!@#$%^&*(),.vbnmtyuiopasd"
Dim sPassword As String

Private Sub cmdClose_Click()
    Unload Me
    End
End Sub

Private Sub cmdDecrypt_Click()
Dim sFileName As String
Dim iFileNo As Integer
Dim sData As String
Dim sDecryptData  As String
Dim sDestinationFile As String
Dim iDestinationFileNo As Integer
    
    txtSource.Enabled = False
    txtDestination.Enabled = False
    cmdEncrypt.Enabled = False
    cmdDecrypt.Enabled = False
    cmdClose.Enabled = False
    
    sDestinationFile = txtDestination.Text
    iDestinationFileNo = FreeFile
    
    Open sDestinationFile For Output As #iDestinationFileNo
        sFileName = txtSource.Text
        iFileNo = FreeFile
        Open sFileName For Input As #iFileNo
            Do While Not EOF(iFileNo)
                Line Input #iFileNo, sData
                sDecryptData = DecryptText(sData, sPassword)
                Print #iDestinationFileNo, sDecryptData
                DoEvents
            Loop
        Close #iDestinationFileNo
        MsgBox "Decryption compleated.", vbInformation + vbOKOnly, App.Title
    Close #iFileNo
    
    txtSource.Enabled = True
    txtDestination.Enabled = True
    cmdEncrypt.Enabled = True
    cmdDecrypt.Enabled = True
    txtSource.Text = ""
    txtDestination.Text = ""
    cmdClose.Enabled = True
End Sub

Private Sub cmdEncrypt_Click()
Dim sFileName As String
Dim iFileNo As Integer
Dim sData As String
Dim sEncryptData As String
Dim sDestinationFile As String
Dim iDestinationFileNo As Integer
    
    txtSource.Enabled = False
    txtDestination.Enabled = False
    cmdEncrypt.Enabled = False
    cmdDecrypt.Enabled = False
    cmdClose.Enabled = False
    
    sDestinationFile = txtDestination.Text
    iDestinationFileNo = FreeFile
    
    Open sDestinationFile For Output As #iDestinationFileNo
        sFileName = txtSource.Text
        iFileNo = FreeFile
        Open sFileName For Input As #iFileNo
            Do While Not EOF(iFileNo)
                Line Input #iFileNo, sData
                sEncryptData = EncryptText(sData, sPassword)
                Print #iDestinationFileNo, sEncryptData
                DoEvents
            Loop
        Close #iDestinationFileNo
        MsgBox "Encryption compleated.", vbInformation + vbOKOnly, App.Title
    Close #iFileNo
    
    txtSource.Enabled = True
    txtDestination.Enabled = True
    cmdEncrypt.Enabled = True
    cmdDecrypt.Enabled = True
    txtSource.Text = ""
    txtDestination.Text = ""
    cmdClose.Enabled = True
End Sub

Private Function EncryptText(TextToEncrypt As String, Password As String)
Dim sTextFirst As String
Dim sTextSecond As String
Dim sResults As String
Dim iCount As Integer
Dim sMid As String
Dim iPosition As Integer

    sTextFirst = sDefaultValueFirst
    sTextSecond = sDefaultValueSecond
    
    JoinPassword sTextFirst, sTextSecond, Password
    sResults = ""
    For iCount = 1 To Len(TextToEncrypt)
        sMid = Mid(TextToEncrypt, iCount, 1)
        iPosition = InStr(1, sTextFirst, sMid, vbBinaryCompare)
        If iPosition > 0 Then
            sResults = sResults & Mid(sTextSecond, iPosition, 1)
        Else
            sResults = sResults & sMid
        End If
            sTextFirst = LeftShift(sTextFirst)
            sTextSecond = RightShift(sTextSecond)
        DoEvents
    Next iCount
    EncryptText = sResults
End Function

Private Function DecryptText(TextToDecrypt As String, Password As String)
Dim sTextFirst As String
Dim sTextSecond As String
Dim sResults As String
Dim iCount As Integer
Dim sMid As String
Dim iPosition As Integer

    sTextFirst = sDefaultValueFirst
    sTextSecond = sDefaultValueSecond
    JoinPassword sTextFirst, sTextSecond, Password
    sResults = ""
     For iCount = 1 To Len(TextToDecrypt)
        sMid = Mid(TextToDecrypt, iCount, 1)
        iPosition = InStr(1, sTextSecond, sMid, vbBinaryCompare)
        If iPosition > 0 Then
            sResults = sResults & Mid(sTextFirst, iPosition, 1)
        Else
            sResults = sResults & sMid
        End If
        sTextFirst = LeftShift(sTextFirst)
        sTextSecond = RightShift(sTextSecond)
    Next iCount
    DecryptText = sResults
End Function

Private Sub JoinPassword(ByRef sFirstString As String, ByRef sSecondString As String, sPassword As String)
Dim iCharLen As Long
Dim iCount As Long

For iCharLen = 1 To Len(sPassword)
    For iCount = 1 To Asc(Mid(sPassword, iCharLen, 1)) * iCharLen
        sFirstString = LeftShift(sFirstString)
        sSecondString = RightShift(sSecondString)
    Next iCount
Next iCharLen
End Sub

Function LeftShift(PassString As String) As String
    If Len(PassString) > 0 Then LeftShift = Mid(PassString, 2, Len(PassString) - 1) & Mid(PassString, 1, 1)
End Function

Function RightShift(PassString As String) As String
    If Len(PassString) > 0 Then RightShift = Mid(PassString, Len(PassString), 1) & Mid(PassString, 1, Len(PassString) - 1)
End Function

Private Sub cmdEncryptaFile_Click()
    Frame1.Visible = True
    Frame2.Visible = False
End Sub

Private Sub cmdEncryptSomeText_Click()
    Frame1.Visible = False
    Frame2.Visible = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
    End
End Sub

Private Sub cmdFileForEncrypt_Click()
Dim sDestination As String

    With cdMain
        .DialogTitle = App.ProductName & "-Select a file."
        .Filter = "All files."
        .ShowSave
    End With
    sDestination = Trim(cdMain.FileName)
    txtDestination.Text = sDestination
    
End Sub

Private Sub cmdFileToEncrypt_Click()
Dim sSourcePath As String

    With cdMain
        .DialogTitle = App.ProductName & "- Select a file."
        .Filter = "All files"
        .ShowOpen
    End With
    sSourcePath = Trim(cdMain.FileName)
    txtSource.Text = sSourcePath
End Sub

Private Sub cmdTextDecrypt_Click()
    txtResult.Text = DecryptText(txtSourceText.Text, sPassword)
End Sub

Private Sub cmdTextEncrypt_Click()
    txtResult.Text = EncryptText(txtSourceText.Text, sPassword)
End Sub

Private Sub Form_Load()
sPassword = Winsock1.LocalIP
txtPassword.Text = sPassword
End Sub
