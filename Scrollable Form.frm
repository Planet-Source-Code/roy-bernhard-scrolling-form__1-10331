VERSION 5.00
Begin VB.Form ScrollForm 
   Caption         =   "ScrollForm"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   LinkTopic       =   "ScrollForm"
   ScaleHeight     =   6435
   ScaleWidth      =   9180
   Begin VB.PictureBox picMain 
      Height          =   6435
      Left            =   0
      ScaleHeight     =   6375
      ScaleWidth      =   9120
      TabIndex        =   0
      Top             =   0
      Width           =   9180
      Begin VB.PictureBox picContainer 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   6255
         Left            =   0
         ScaleHeight     =   6225
         ScaleWidth      =   8970
         TabIndex        =   1
         Top             =   0
         Width           =   8995
      End
   End
   Begin VB.VScrollBar fsbVert 
      Height          =   1335
      Left            =   4500
      TabIndex        =   3
      Top             =   5700
      Width           =   255
   End
   Begin VB.HScrollBar fsbHorz 
      Height          =   255
      Left            =   3660
      TabIndex        =   2
      Top             =   6540
      Width           =   1635
   End
End
Attribute VB_Name = "ScrollForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    picMain.Height = Me.ScaleHeight - fsbHorz.Height
    picMain.Width = Me.ScaleWidth - fsbVert.Width
    
    fsbVert.Left = picMain.Width
    fsbVert.Top = picMain.Top
    
    fsbHorz.Left = picMain.Left
    fsbHorz.Top = picMain.Height
    
    fsbHorz.Width = picMain.Width
    fsbVert.Height = picMain.Height
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    'Setup Scroll Bar locations
    picMain.Height = Me.ScaleHeight - fsbHorz.Height
    picMain.Width = Me.ScaleWidth - fsbVert.Width
    
    fsbVert.Left = picMain.Width
    fsbVert.Top = picMain.Top
    
    fsbHorz.Left = picMain.Left
    fsbHorz.Top = picMain.Height
    
    fsbHorz.Width = picMain.Width
    fsbVert.Height = picMain.Height
    
    If picMain.Width <= picContainer.Width Then
        fsbHorz.Max = picMain.ScaleWidth - picContainer.Width
        fsbHorz.LargeChange = picMain.ScaleWidth
        fsbHorz.SmallChange = picMain.ScaleWidth / 5
        'fsbHorz.Visible = True
    Else
        fsbHorz.Max = 0
        'fsbHorz.Visible = False
    End If
    
    If picMain.Height <= picContainer.Height Then
        fsbVert.Max = picMain.ScaleHeight - picContainer.Height
        fsbVert.LargeChange = picMain.ScaleHeight
        fsbVert.SmallChange = picMain.ScaleHeight / 5
        'fsbVert.Visible = True
    Else
        fsbVert.Max = 0
        'fsbVert.Visible = False
    End If
    
End Sub

Private Sub fsbHorz_Change()
    
    picContainer.Left = fsbHorz.Value

End Sub

Private Sub fsbHorz_Scroll()
    
    picContainer.Left = fsbHorz.Value

End Sub

Private Sub fsbVert_Change()
    
    picContainer.Top = fsbVert.Value
    
End Sub

Private Sub fsbVert_Scroll()
    
    picContainer.Top = fsbVert.Value
    
End Sub
