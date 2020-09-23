VERSION 5.00
Begin VB.Form frmmain 
   AutoRedraw      =   -1  'True
   Caption         =   "Custom Toolbar Example"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   135
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   374
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   4275
      TabIndex        =   5
      Top             =   1470
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   180
      TabIndex        =   4
      Top             =   825
      Width           =   3795
   End
   Begin VB.PictureBox picToolbar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00D3DADD&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   165
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   141
      TabIndex        =   2
      Top             =   105
      Width           =   2115
   End
   Begin VB.PictureBox picSrcBut 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00D3DADD&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   465
      ScaleHeight     =   330
      ScaleWidth      =   345
      TabIndex        =   1
      Top             =   1395
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1020
      Picture         =   "frmmain.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   1680
      TabIndex        =   0
      Top             =   1455
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label lbltooltip 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   255
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Custom toolbar example
' By Ben Jones
'
' Please vote if you like this code
' If you use this code in your own projects let me know as I like to know what people think of it.

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Enum ButtonDirection
    ButtonDown = 0
    ButtonUp = 1
    ButtonFlat = 2
End Enum

Private Const ButtonHeight As Integer = 22
Private Const ButtonWidth As Integer = 23
Private NoOfButtons As Integer
Private ButtonToolTip()
Private ButtonX As Integer, ButtonY As Integer

Private Sub MakeSingleButton(tEffect As ButtonDirection)
    DrawEffect tEffect
    TransparentBlt picSrcBut.hdc, 3, 3, 16, 16, picSrc.hdc, ButtonX * 16, 0, 16, 16, RGB(255, 0, 255)
    picSrcBut.Refresh
    ' now draw it to the toolbar
    BitBlt picToolbar.hdc, 1 + ButtonX * ButtonWidth, 4, ButtonWidth, ButtonHeight, picSrcBut.hdc, 0, 0, vbSrcCopy ' blit the new button image along the form
    picSrcBut.Cls
    picToolbar.Refresh
End Sub

Private Sub MakeToolbar()
Dim I As Integer ' hold the number of buttons we need to draw
    lbltooltip.Visible = False
    NoOfButtons = (picSrc.Width / ButtonWidth) + 1 ' how many button are in the strip
    For I = 0 To NoOfButtons
        picSrcBut.Cls ' clear the contents of the pic source picturebox
        DrawEffect ButtonFlat ' apply the flat effect to the pic source picturebox
        TransparentBlt picSrcBut.hdc, 3, 3, 16, 16, picSrc.hdc, 16 * I, 0, 16, 16, RGB(255, 0, 255)
        ' the line above blits the image for the button from picSrc and uses the back colour for the tansparent colour = 255, 0, 255
        BitBlt picToolbar.hdc, 1 + I * ButtonWidth, 4, ButtonWidth, ButtonHeight, picSrcBut.hdc, 0, 0, vbSrcCopy ' blit the new button image along the form
        picToolbar.Refresh ' refresh the form so buttons do not disapear
    Next
    picSrcBut.Cls
    I = 0
End Sub

Private Sub ToolbarButtonUp(ButtonIndex As Integer)
    ' add your button code here
    Select Case ButtonIndex
        Case 0
            Text1.Text = "You clicked button new"
        Case 1
            Text1.Text = "You clicked button Open"
        Case 2
            Text1.Text = "You clicked button Save"
        Case 3
            Text1.Text = "You clicked button Cut"
        Case 4
            Text1.Text = "You clicked button Copy"
        Case 5
            Text1.Text = "You clicked button Paste"
        Case 6
            Text1.Text = "You clicked button Find"
    End Select
    
End Sub

Private Sub ToolbarButtonClick(ButtonIndex As Integer)
    ' add your button code here
    Select Case ButtonIndex
        Case 0
            
        Case 1
        '
        Case 2
        '
        Case 3
        '
        Case 4
        '
        Case 5
        '
        Case 6
        
    End Select
    
End Sub
Private Sub Draw3DLine(Y1 As Integer, Y2 As Integer, PicBox As PictureBox)
    ' This is used to draw a 3d line along the top of the form
    PicBox.Line (0, Y1)-(PicBox.ScaleWidth, Y1), &H8000000C
    PicBox.Line (0, Y2)-(PicBox.ScaleWidth, Y2), vbWhite
End Sub

Private Sub DrawEffect(Direction As ButtonDirection)
    If Direction = ButtonUp Then
        picSrcBut.Line (picSrcBut.ScaleWidth, 0)-(0, 0), vbWhite
        picSrcBut.Line (0, picSrcBut.ScaleHeight)-(0, -8), vbWhite
        picSrcBut.Line (picSrcBut.ScaleWidth - 8, 0)-(picSrcBut.ScaleWidth - 8, picSrcBut.ScaleHeight), &H8000000C
        picSrcBut.Line (picSrcBut.ScaleWidth, picSrcBut.ScaleHeight - 8)-(-8, picSrcBut.ScaleHeight - 8), &H8000000C
        Exit Sub
    ElseIf Direction = ButtonDown Then
        picSrcBut.Line (picSrcBut.ScaleWidth, 0)-(0, 0), &H8000000C
        picSrcBut.Line (0, picSrcBut.ScaleHeight)-(0, -8), &H8000000C
        picSrcBut.Line (picSrcBut.ScaleWidth - 8, 0)-(picSrcBut.ScaleWidth - 8, picSrcBut.ScaleHeight), vbWhite
        picSrcBut.Line (picSrcBut.ScaleWidth, picSrcBut.ScaleHeight - 8)-(-8, picSrcBut.ScaleHeight - 8), vbWhite
    Else
        picSrcBut.Line (picSrcBut.ScaleWidth, 0)-(0, 0), &HD3DADD
        picSrcBut.Line (0, picSrcBut.ScaleHeight)-(0, -8), &HD3DADD
        picSrcBut.Line (picSrcBut.ScaleWidth - 8, 0)-(picSrcBut.ScaleWidth - 8, picSrcBut.ScaleHeight), &HD3DADD
        picSrcBut.Line (picSrcBut.ScaleWidth, picSrcBut.ScaleHeight - 8)-(-8, picSrcBut.ScaleHeight - 8), &HD3DADD
    End If
End Sub

Private Sub cmdexit_Click()
    End
End Sub

Private Sub Form_Load()
    ' position and resize the pictire box for our toolbar
    picToolbar.Width = frmmain.ScaleWidth
    picToolbar.Height = 29
    picToolbar.Top = 0
    picToolbar.Left = 0
    Draw3DLine 0, 1, picToolbar ' draw the 3dline alog the top of the form
    MakeToolbar
    Draw3DLine 27, 28, picToolbar
    
    ' setup the tooltip info text for each buttton
    ReDim Preserve ButtonToolTip(NoOfButtons)
    ButtonToolTip(0) = " New... "
    ButtonToolTip(1) = " Open New project File.... "
    ButtonToolTip(2) = " Save... "
    ButtonToolTip(3) = " Cut... "
    ButtonToolTip(4) = " Copy... "
    ButtonToolTip(5) = " Paste... "
    ButtonToolTip(6) = " Find... "
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MakeToolbar
End Sub

Private Sub picToolbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If ButtonX <= NoOfButtons Then MakeSingleButton ButtonDown
        ToolbarButtonClick ButtonX
    End If
End Sub

Private Sub picToolbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ButtonX = Int(X / ButtonWidth) ' used to keep trak of the button we are over
    ButtonY = Int(Y / ButtonHeight) ' used to keep trak of the button we are over
    
    lbltooltip.Visible = False ' hide the tooltip label
    MakeToolbar ' draw the toolbar
    
    If ButtonX <= NoOfButtons Then
        'if the ButtonX is lower or equal then we do the code below
        MakeSingleButton ButtonUp ' draw the button upstate
        lbltooltip.Visible = True ' show the tooltip label
        lbltooltip.Left = (ButtonX * ButtonWidth) ' position the tooltip label
        lbltooltip.Top = (ButtonHeight + lbltooltip.Height - 4)
        lbltooltip.Caption = ButtonToolTip(ButtonX)
    End If
    
    
End Sub

Private Sub picToolbar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If ButtonX <= NoOfButtons Then MakeSingleButton ButtonUp
        ToolbarButtonUp ButtonX
    End If
End Sub
