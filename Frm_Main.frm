VERSION 5.00
Begin VB.Form Frm_Main 
   BorderStyle     =   0  'None
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Colors!"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "NO API!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"Frm_Main.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Y:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Current X:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "  Drag Example"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This code allows you to do real time updating and code execution while
' a drag is being done.

' Example by Digital Rampage, aka John Harper
' You may freely use this code wherever you like

Dim UserDown As Boolean
Dim UserStartX, UserStartY


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    ' Keep track of when the user wants to drag
    ' The mandatory code:
    ' UserDown tells us they clicked our label
    UserDown = True
    ' UserX catches where the X position of their mouse is
    UserStartX = X
    ' UserY catches where the Y position of their mouse is
    ' The X and Y values are important if we want to set the form
    ' in the correct spot. Without them the form will zap the top and left
    ' right to the mouse position. Try removing them to see what i mean.
    UserStartY = Y
    ' End mandatory code
    Label1.Caption = "  Relocating Form"
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    ' If the user has the mouse button down and they move the mouse
    ' then they must want to drag. We take the current form position
    ' and set it equal to the current position plus where the user's
    ' mouse is on the form minus where they first clicked on the form
    ' to do the drag.
    If UserDown = True Then
        ' These are the only mandatory lines to make the form drag work:
        Frm_Main.Left = Frm_Main.Left + (X - UserStartX)
        Frm_Main.Top = Frm_Main.Top + (Y - UserStartY)
        ' End mandatory code
        
        
        'Unneeded Code:
        'Updates the value labes
        Label2.Caption = "Current X: " & Me.Left
        Label3.Caption = "Current Y: " & Me.Top
        
        If Check1.Value = vbChecked Then
            If Frm_Main.Left > 1 Then Frm_Main.BackColor = RGB((Frm_Main.Left / 20), (Frm_Main.Top / 20), 100)
        End If

    End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    ' This tells us when the user releases the mouse and wants to
    ' set the form at the current location
    ' we also reset the caption back to normal
    UserDown = False
    Label1.Caption = "  Drag Example"
End Sub
