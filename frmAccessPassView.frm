VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAccessPassView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Access PassView"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   ControlBox      =   0   'False
   FontTransparent =   0   'False
   Icon            =   "frmAccessPassView.frx":0000
   LinkTopic       =   "Form"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtGet 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmAccessPassView.frx":030A
      Top             =   120
      Width           =   4095
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3360
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmAccessPassView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str2000 As String
Dim File As String
Option Explicit

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdGet_Click()
cd.Filter = "Microsoft Access Files (*.mdb)|*.mdb|All Files (*.*)|*.*"
cd.DialogTitle = App.FileDescription
cd.ShowOpen
If Not Len(cd.FileName) = 0 Then
File = cd.FileName
GetPassword
txtGet.Text = "Filename: " & cd.FileName & vbCrLf & "The Password is:" & vbCrLf & str2000
End If
End Sub

Private Function GetPassword()

    On Error GoTo ErrHand

    Dim Access2000Decode As Variant
    
    Dim fFile       As Integer
    Dim bCnt        As Integer
    
    Dim retXPwd(17) As Integer
    Dim wkCode      As Integer
    Dim mgCode      As Integer
    
    
    Access2000Decode = Array(&H6ABA, &H37EC, &HD561, &HFA9C, &HCFFA, _
                      &HE628, &H272F, &H608A, &H568, &H367B, _
                      &HE3C9, &HB1DF, &H654B, &H4313, &H3EF3, _
                      &H33B1, &HF008, &H5B79, &H24AE, &H2A7C)

    If Len(File) > 0 Then
    
        fFile = FreeFile
    
        Open File For Binary As #fFile
            Get #fFile, 67, retXPwd
            Get #fFile, 103, mgCode
        Close #fFile
        
        mgCode = mgCode Xor Access2000Decode(18)

        str2000 = vbNullString

        For bCnt = 0 To 17

            wkCode = retXPwd(bCnt) Xor Access2000Decode(bCnt)
            
            If wkCode < 256 Then
                str2000 = str2000 & Chr(wkCode)
            Else
                str2000 = str2000 & Chr(wkCode Xor mgCode)
            End If
            
        Next bCnt
        
    Else
    
       str2000 = "No file Selected"
    
    End If
    
Exit Function
ErrHand:
    MsgBox "Error with opening file", vbCritical, App.Title


End Function

