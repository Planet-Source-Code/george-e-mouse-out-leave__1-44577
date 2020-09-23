VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2730
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   2730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
      Caption         =   "Command4"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000000FF&
      Caption         =   "Command3"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Command2"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Command1"
      Height          =   375
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseMoveHook Command1
    Command1.Caption = "mousemove"
    Command1.BackColor = vbWhite
    
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseMoveHook Command2
    Command2.Caption = "mousemove"
    Command2.BackColor = vbWhite
    
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseMoveHook Command3
    Command3.Caption = "mousemove"
    Command3.BackColor = vbWhite
    
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseMoveHook Command4
    Command4.Caption = "mousemove"
    Command4.BackColor = vbWhite
    
End Sub

Private Sub Form_Load()

  Dim Control As Control

    Hook Me
    For Each Control In Me.Controls
        If TypeOf Control Is CommandButton Then
            Hook Control
        End If
    Next Control

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseMoveHook Me
    Me.Caption = "mousemove"
    Me.BackColor = vbWhite
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Dim Control As Control

    UnHook Me
    For Each Control In Me.Controls
        If TypeOf Control Is CommandButton Then
            UnHook Control
        End If
    Next Control

End Sub

