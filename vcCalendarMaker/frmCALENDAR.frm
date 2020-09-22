VERSION 5.00
Begin VB.Form frmCALENDAR 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Calendar Maker"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCALENDAR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSAVE 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1680
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cboMONTH 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblSAVE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Save..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblPRINT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblIMAGE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Image..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblEXIT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
End
Attribute VB_Name = "frmCALENDAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'------------------------------------------------------------
' Stores the last image we picked for the calendar
'------------------------------------------------------------
Public LastImage As String
Private Sub cboMONTH_Change()
    On Error Resume Next
    '------------------------------------------------------------
    ' Redraw the calendar on month change
    '------------------------------------------------------------
    Me.Cls
    DrawGrid Me
End Sub
Private Sub cboMONTH_Click()
    cboMONTH_Change
End Sub
Private Sub Form_Load()
    On Error Resume Next
    '------------------------------------------------------------
    ' Prepare the form making it the 50 percent the
    ' size of a standard sheet of paper [US standard
    ' paper that is]
    '------------------------------------------------------------
    Me.LastImage = ""
    Me.BackColor = vbWhite
    Me.Width = ScaleX(8.5, vbInches, vbTwips) / 2
    Me.Height = ScaleX(11, vbInches, vbTwips) / 2
    '------------------------------------------------------------
    ' Need to set AutoRedraw to true since we are drawing
    ' to these things.
    '------------------------------------------------------------
    Me.AutoRedraw = True
    Me.picSAVE.AutoRedraw = True
    '------------------------------------------------------------
    ' Fill the month dropdown with months
    '------------------------------------------------------------
    FillcboMONTH
End Sub
Private Sub FillcboMONTH()
    On Error Resume Next
    Dim X As Long
    '------------------------------------------------------------
    ' Loop 12 times to fill in months
    '------------------------------------------------------------
    For X = 1 To 12
        Me.cboMONTH.AddItem Format(CDate(DateSerial(Year(Date), X, 1)), "MMMM")
        Me.cboMONTH.ItemData(Me.cboMONTH.NewIndex) = X
    Next X
    '------------------------------------------------------------
    ' Set the current month as the default
    '------------------------------------------------------------
    Me.cboMONTH.Text = Format(Date, "MMMM")
End Sub
Public Sub DrawGrid(DrawTo As Object, Optional MonthSize As Long = 18, Optional DaySize As Long = 10)
    On Error Resume Next
    Dim sd As String
    Dim X As Long, Y As Long
    Dim p As Picture, ix As Long, iy As Long
    Dim xStep As Long, v As String
    Dim yStep As Long
    Screen.MousePointer = vbHourglass
    '------------------------------------------------------------
    ' Get step values for the loops to draw the grid
    '------------------------------------------------------------
    xStep = DrawTo.ScaleWidth / 7
    yStep = DrawTo.ScaleHeight / 2 / 5
    '------------------------------------------------------------
    ' If there was an image already picked, draw it
    ' to the top half of the paper/form.
    '------------------------------------------------------------
    If Me.LastImage <> "" Then
        Set p = LoadPicture(Me.LastImage)
        DrawTo.PaintPicture p, 0, 0, DrawTo.ScaleWidth, DrawTo.ScaleHeight / 2
    End If
    DrawTo.FontBold = True
    DrawTo.FontSize = MonthSize
    DrawTo.ForeColor = vbWhite
    DrawTo.CurrentX = DrawTo.ScaleWidth / 2 - (DrawTo.TextWidth(cboMONTH.Text) / 2) - 15
    DrawTo.CurrentY = DrawTo.ScaleHeight / 2 - (DrawTo.TextHeight(cboMONTH.Text)) - 15
    DrawTo.Print cboMONTH.Text
    DrawTo.ForeColor = vbBlack
    DrawTo.CurrentX = DrawTo.ScaleWidth / 2 - (DrawTo.TextWidth(cboMONTH.Text) / 2)
    DrawTo.CurrentY = DrawTo.ScaleHeight / 2 - (DrawTo.TextHeight(cboMONTH.Text))
    DrawTo.Print cboMONTH.Text
    DrawTo.FontSize = DaySize
    '------------------------------------------------------------
    ' Draw a box around the whole form/paper
    '------------------------------------------------------------
    DrawTo.Line (0, 0)-(DrawTo.ScaleWidth - (ScaleX(1, vbPixels, vbTwips)), DrawTo.ScaleHeight - (ScaleY(1, vbPixels, vbTwips))), vbBlack, B
    '------------------------------------------------------------
    ' Starting at the first of the month picked, loop
    ' backwards till we find Sunday so we have a starting
    ' point.
    '------------------------------------------------------------
    sd = CDate(DateSerial(Year(Date), Me.cboMONTH.ItemData(Me.cboMONTH.ListIndex), 1))
    While DatePart("w", sd) <> 1
        sd = DateAdd("d", -1, sd)
    Wend
    '------------------------------------------------------------
    ' Draw vertical lines
    '------------------------------------------------------------
    For X = xStep To DrawTo.ScaleWidth Step xStep
        DrawTo.Line (X, DrawTo.ScaleHeight / 2)-(X, DrawTo.ScaleHeight), vbBlack
    Next X
    '------------------------------------------------------------
    ' Draw horizontal lines
    '------------------------------------------------------------
    For Y = DrawTo.ScaleHeight / 2 To DrawTo.ScaleHeight Step yStep
        DrawTo.Line (0, Y)-(DrawTo.ScaleWidth, Y), vbBlack
    Next Y
    '------------------------------------------------------------
    ' Put the numbers in.
    '------------------------------------------------------------
    DrawTo.FontBold = False
    iy = DrawTo.ScaleHeight / 2
    For Y = 1 To 5
        ix = 0
        For X = 1 To 7
            DrawTo.CurrentX = ix + 10
            DrawTo.CurrentY = iy + 15
            DrawTo.Print DatePart("d", sd)
            sd = DateAdd("d", 1, sd)
            ix = ix + xStep
        Next X
        iy = iy + yStep
    Next Y
    '------------------------------------------------------------
    ' If we did not have enough squares to finish the
    ' month then
    '------------------------------------------------------------
    If DatePart("m", sd) = cboMONTH.ItemData(cboMONTH.ListIndex) Then
        '------------------------------------------------------------
        ' Lets back up...
        '------------------------------------------------------------
        DrawTo.FillColor = vbWhite
        iy = iy - yStep
        ix = 0
        '------------------------------------------------------------
        ' Loop till we finish the month
        '------------------------------------------------------------
        While DatePart("m", sd) = cboMONTH.ItemData(cboMONTH.ListIndex)
            '------------------------------------------------------------
            ' Get a string of seven days prior and the new
            ' day we are making and put them together with
            ' a slash
            '------------------------------------------------------------
            v = DatePart("d", DateAdd("d", -7, sd)) & "/" & DatePart("d", sd)
            '------------------------------------------------------------
            ' Draw a white solid box over the old number we
            ' had in there.
            '------------------------------------------------------------
            DrawTo.Line (ix + 10, iy + 15)-(ix + 10 + DrawTo.TextWidth(v), iy + 15 + DrawTo.TextHeight(v)), vbWhite, BF
            '------------------------------------------------------------
            ' Write out this new string.
            '------------------------------------------------------------
            DrawTo.CurrentX = ix + 10
            DrawTo.CurrentY = iy + 15
            DrawTo.Print v
            '------------------------------------------------------------
            ' Add a day and move our X position
            '------------------------------------------------------------
            sd = DateAdd("d", 1, sd)
            ix = ix + xStep
        Wend
    End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
End Sub
Private Sub lblEXIT_Click()
    On Error Resume Next
    '------------------------------------------------------------
    ' Unload this form
    '------------------------------------------------------------
    Unload Me
End Sub
Private Sub lblIMAGE_Click()
On Error Resume Next
    Dim fNAME As String, obj As CDLG
    '------------------------------------------------------------
    ' using the CDLG class, show the Common Dialog
    ' box asking for a file to load.  If one is picked,
    ' set LastImage and call the draw function again.
    '------------------------------------------------------------
    Set obj = New CDLG
    obj.VBGetOpenFileName fNAME, , , , , , "Image Files|*.bmp;*.jpg;*.gif|All Files|*.*", , CurDir, "Load Image", , Me.hwnd
    If fNAME <> "" Then
        Me.LastImage = fNAME
        DrawGrid Me
    End If
End Sub
Private Sub lblPRINT_Click()
    On Error Resume Next
    Dim obj As CDLG, p As Object
    '------------------------------------------------------------
    ' Set our hidden picture box to the size of paper
    ' with quarter inch margins, the draw the calendar
    ' to it then paint that image to the printer.
    '------------------------------------------------------------
    Set obj = New CDLG
    Set p = Printer
    obj.VBPrintDlg Me.hdc, , , , , , , True, False, , , , Me.hwnd, p
    If obj.SelectedPrinterName <> "" Then
        For Each p In Printers
            If p.DeviceName = obj.SelectedPrinterName Then
                Set Printer = p
                Exit For
            End If
        Next
    End If
    If obj.SelectedPrinterName <> "" Then
        Screen.MousePointer = vbHourglass
        With Me.picSAVE
            .Cls
            .Width = ScaleX(8, vbInches, vbTwips)
            .Height = ScaleX(10.5, vbInches, vbTwips)
            DrawGrid picSAVE
        End With
        Printer.PaintPicture picSAVE.Image, 0, 0, Printer.ScaleWidth, Printer.ScaleHeight
        Printer.EndDoc
        Screen.MousePointer = vbDefault
    End If
End Sub
Private Sub lblSAVE_Click()
    On Error Resume Next
    Dim fNAME As String, obj As CDLG
    '------------------------------------------------------------
    ' Using the CDLG class, show the Common Dialog
    ' box asking for a place to save to, if a file
    ' name is given then draw our calendar to the hidden
    ' picture box which is the size of a stanard piece
    ' of paper with quarter inch margins.  Then save
    ' that image to the file given.
    '------------------------------------------------------------
    Set obj = New CDLG
    obj.VBGetSaveFileName fNAME, , , "BMP Files|*.bmp", , CurDir, "Save to", "*.bmp", Me.hwnd
    If fNAME <> "" Then
        Screen.MousePointer = vbHourglass
        With Me.picSAVE
            .Cls
            .Width = ScaleX(8, vbInches, vbTwips)
            .Height = ScaleX(10.5, vbInches, vbTwips)
            DrawGrid picSAVE
            SavePicture .Image, fNAME
        End With
        Screen.MousePointer = vbDefault
    End If
End Sub
