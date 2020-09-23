VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin Project1.Command Command2 
      Height          =   675
      Left            =   360
      TabIndex        =   2
      Top             =   930
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   1191
      Caption         =   "Command1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.Command Command1 
      Height          =   690
      Left            =   945
      TabIndex        =   1
      Top             =   1755
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   1217
      Caption         =   "&Command1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.Command Command3 
      Height          =   465
      Left            =   1620
      TabIndex        =   0
      Top             =   405
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   820
      Enabled         =   0   'False
      Caption         =   "&Testing"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontColour      =   12632256
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdButton1_Click()
    If CmdDisabled.Enabled = True Then
        CmdDisabled.Enabled = False
    Else
        CmdDisabled.Enabled = True
    End If
End Sub

Private Sub CmdDisabled_Click()
    MsgBox "Click"
End Sub

Private Sub Command1_Click()
    If Command3.Enabled = True Then
        Command3.Enabled = False
    Else
        Command3.Enabled = True
    End If
    
    MsgBox "Command1 Click"
End Sub

Private Sub Command3_Click()
    MsgBox "Command3 click"
End Sub

Private Sub Form_Load()
    InitialiseButtons Form1, True, True 'Must Call this to initialise the buttons
End Sub

Public Sub InitialiseButtons(FrmForm As Form, Initialise As Boolean, Optional BytEnabled As Boolean)
On Local Error Resume Next
Dim Control As Object 'Define the variable control as an object
    For Each Control In FrmForm 'Check all controls on a form
        If TypeOf Control Is Command Then
            If Initialise = True Then
                Control.Initialize 'If its a XP command button then initialise it.
            Else
                If BytEnabled = True Then
                    Control.Enabled = True
                Else
                    Control.Enabled = False
                End If
            End If
        End If
    Next
End Sub


