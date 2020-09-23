VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " SNMP Demo Form"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbStartOid 
      Height          =   315
      Left            =   6975
      TabIndex        =   9
      Top             =   4035
      Width           =   2655
   End
   Begin VB.CommandButton btnSNMPgetNext 
      Caption         =   "Get Next SNMP"
      Height          =   435
      Left            =   1800
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox txtResults 
      Height          =   3135
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   720
      Width           =   9375
   End
   Begin VB.TextBox txtHostIPadr 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Text            =   "132.112.150.11"
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton btnSNMPgetFirst 
      Caption         =   "Get SNMP"
      Height          =   435
      Left            =   240
      TabIndex        =   5
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox txtOID 
      Height          =   285
      Left            =   5160
      TabIndex        =   3
      Text            =   "1.3.6.1.2.1.1.1.0"
      Top             =   360
      Width           =   4455
   End
   Begin VB.TextBox txtSNMPcounityName 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Text            =   "public"
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblDesc 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"frmMain.frx":0000
      Height          =   840
      Left            =   225
      TabIndex        =   12
      Top             =   4455
      Width           =   5805
   End
   Begin VB.Label lblSNMPlevel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " SNMP version"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   6975
      TabIndex        =   11
      Top             =   4500
      Width           =   2640
   End
   Begin VB.Label Label4 
      Caption         =   "Start Walking at MIB Value"
      Height          =   255
      Left            =   4875
      TabIndex        =   10
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "OID"
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Community Name"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Remote IP Address"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'   A simple demo form for experiments.
'


Option Explicit

Public WithEvents oSNMP As clsSNMP
Attribute oSNMP.VB_VarHelpID = -1
'
'
'

Private Sub cbStartOid_Click()
   
   txtOID.Text = Mid$(cbStartOid.Text, InStr(1, cbStartOid.Text, " - ") + 3)
   
End Sub

Private Sub btnSNMPgetFirst_Click()
   
    With oSNMP
        .DestHost = txtHostIPadr.Text
        .ComunityName = txtSNMPcounityName.Text
        .SNMPGet txtOID.Text
    End With

End Sub

Private Sub btnSNMPgetNext_Click()
   
   oSNMP.SNMPGetNext
   
End Sub


Private Sub Form_Load()
    
    With cbStartOid
        .AddItem "MIB2 - 1.3.6.1.2.1."
        .AddItem "Enterprise - 1.3.6.1.4.1"
        .AddItem "SNMP Modules - 1.3.6.1.6.3"
    End With
    Set oSNMP = New clsSNMP
    
End Sub

Private Sub oSNMP_Error(sErrMsg As String)
    
    MsgBox sErrMsg, vbExclamation, " Error:"
    
End Sub

Private Sub oSNMP_Result(sValue As String)
    
    With txtResults
        .Text = .Text + sValue + vbCrLf
        .SelStart = Len(.Text)
    End With
        
    txtOID.Text = oSNMP.Oid
    lblSNMPlevel.Caption = " SNMP version: " & oSNMP.SNMPversion
    DoEvents
    
End Sub


' #*#
