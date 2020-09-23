VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00004000&
   Caption         =   "Calculate LED Resistor Value"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   7290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "General Guide"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   3240
      TabIndex        =   24
      Top             =   75
      Width           =   3945
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1530
         Left            =   2235
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   1500
         ScaleWidth      =   1500
         TabIndex        =   28
         Top             =   1725
         Width           =   1530
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Equation for Resistor"
         Height          =   240
         Left            =   255
         TabIndex        =   30
         Top             =   2385
         Width           =   1680
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "R = (Vs - Vl)/ I"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   29
         Top             =   2625
         Width           =   1995
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":7574
         Height          =   840
         Left            =   105
         TabIndex        =   27
         Top             =   840
         Width           =   3720
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "blue and white - 30ma / 4 volts"
         Height          =   255
         Left            =   105
         TabIndex        =   26
         Top             =   435
         Width           =   2610
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "red and green - 20ma / 2 volts"
         Height          =   255
         Left            =   105
         TabIndex        =   25
         Top             =   210
         Width           =   2580
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Led Values"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   90
      TabIndex        =   3
      Top             =   75
      Width           =   3105
      Begin VB.CommandButton cmdCalculate 
         Caption         =   "Calculate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         TabIndex        =   16
         Top             =   1905
         Width           =   2955
      End
      Begin VB.OptionButton optLED 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Series"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2115
         TabIndex        =   15
         Top             =   270
         Width           =   900
      End
      Begin VB.OptionButton optLED 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Parallel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1065
         TabIndex        =   14
         Top             =   255
         Width           =   975
      End
      Begin VB.OptionButton optLED 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Single"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Top             =   255
         Value           =   -1  'True
         Width           =   825
      End
      Begin VB.TextBox txtNoLED 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   11
         Text            =   "1"
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtDFC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   10
         Text            =   "20"
         Top             =   1230
         Width           =   495
      End
      Begin VB.TextBox txtDFV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Text            =   "2"
         Top             =   915
         Width           =   495
      End
      Begin VB.TextBox txtSV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Text            =   "9"
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Supply Current"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1065
         TabIndex        =   23
         Top             =   3075
         Width           =   1455
      End
      Begin VB.Label lbltotCur 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   60
         TabIndex        =   22
         Top             =   3045
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Resistor Wattage"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1050
         TabIndex        =   21
         Top             =   2760
         Width           =   1590
      End
      Begin VB.Label lblWatt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   60
         TabIndex        =   20
         Top             =   2730
         Width           =   975
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Resistor Value (ohms)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1050
         TabIndex        =   18
         Top             =   2460
         Width           =   1965
      End
      Begin VB.Label lblResistorValue 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   60
         TabIndex        =   17
         Top             =   2415
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of LED's"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   570
         TabIndex        =   12
         Top             =   1605
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Diode Forward Current (ma)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   570
         TabIndex        =   7
         Top             =   1275
         Width           =   2385
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Diode Forward Voltage"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   570
         TabIndex        =   6
         Top             =   945
         Width           =   1995
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Source Voltage"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   570
         TabIndex        =   5
         Top             =   630
         Width           =   1275
      End
   End
   Begin VB.PictureBox picVert 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   6585
      Picture         =   "Form1.frx":7642
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   2
      Top             =   540
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox picHor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   6570
      Picture         =   "Form1.frx":9436
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   1
      Top             =   1395
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox picCir 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H000040C0&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   75
      ScaleHeight     =   187
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   474
      TabIndex        =   0
      Top             =   3555
      Width           =   7140
      Begin VB.Line Line4 
         Visible         =   0   'False
         X1              =   270
         X2              =   326
         Y1              =   216
         Y2              =   216
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "O (gnd)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3060
         TabIndex        =   19
         Top             =   1425
         Width           =   645
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   150
         Left            =   2805
         Shape           =   3  'Circle
         Top             =   1470
         Width           =   135
      End
      Begin VB.Line Line3 
         X1              =   200
         X2              =   254
         Y1              =   216
         Y2              =   216
      End
      Begin VB.Line Line2 
         X1              =   133
         X2              =   159
         Y1              =   102
         Y2              =   102
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000040C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   315
         Left            =   1170
         Top             =   1395
         Width           =   825
      End
      Begin VB.Line Line1 
         X1              =   29
         X2              =   79
         Y1              =   103
         Y2              =   103
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   300
         Shape           =   3  'Circle
         Top             =   1485
         Width           =   150
      End
      Begin VB.Label lblVolt 
         BackStyle       =   0  'Transparent
         Caption         =   "+ Vcc"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   1125
         Width           =   525
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Ken Foster
'May 2010
'use any way you want, no copyrights invoked

Option Explicit

Private Declare Function BitBlt Lib "gdi32" ( _
      ByVal hDestDC As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long, _
      ByVal hSrcDC As Long, _
      ByVal xSrc As Long, _
      ByVal ySrc As Long, _
      ByVal dwRop As Long) As Long

Private Sub Form_Load()
   BitBlt picCir.hDC, 140, 61, picHor.ScaleWidth, picHor.ScaleHeight, picVert.hDC, 0, 0, vbSrcCopy
End Sub

Private Sub Form_Resize()
   If Form1.Width < 7410 Then Form1.Width = 7410
   picCir.Width = Form1.Width - 300
End Sub

Private Sub cmdCalculate_Click()
  Dim x As Integer
  Dim totVD As Single
   If txtSV.Text = "" Or txtSV.Text = "0" Then Exit Sub
   If txtDFV.Text = "" Or txtDFV.Text = "0" Then Exit Sub
   If txtDFC.Text = "" Or txtDFC.Text = "0" Then Exit Sub
   If txtNoLED.Text = "" Or txtNoLED.Text = "0" Then Exit Sub

   totVD = Val(txtDFV.Text) * Val(txtNoLED.Text)   'voltage needed as per number of led's
   picCir.Cls

   If optLED(0).Value = True Then                              'single
      If totVD > Val(txtSV.Text) Then                           ' if voltage required by number of led's is greater than supply voltage
         lblResistorValue.Caption = ""
         lblWatt.Caption = ""
         lbltotCur.Caption = ""
         picCir.Cls
         Exit Sub
      End If
      'calculate values
      lblResistorValue.Caption = Format((Val(txtSV.Text) - Val(txtDFV.Text)) / (Val(txtDFC.Text) / 1000), "#,###,###")
      lblWatt.Caption = (Val(txtSV.Text) * ((Val(txtDFC.Text)) / 1000))
      lbltotCur.Caption = Val(txtDFC.Text) / 1000
      'draw circuit
      BitBlt picCir.hDC, 140, 61, picHor.ScaleWidth, picHor.ScaleHeight, picVert.hDC, 0, 0, vbSrcCopy
      Shape3.Left = 187
      Shape3.Top = 98
      Label6.Top = 95
      Label6.Left = 204
   End If

   If optLED(1).Value = True Then                                   'parallel
      'calculate values
      lblResistorValue.Caption = Format(Val(txtSV.Text - Val(txtDFV.Text)) / ((Val(txtDFC.Text) * Val(txtNoLED.Text)) / 1000), "#,###,###")
      lblWatt.Caption = Val(txtSV.Text) * (((Val(txtDFC.Text) * Val(txtNoLED.Text)) / 1000))
      lbltotCur.Caption = ((Val(txtDFC.Text) * Val(txtNoLED.Text)) / 1000)
      'draw circuit
      For x = 1 To Val(txtNoLED.Text)
         BitBlt picCir.hDC, 100 + (50 * (x)), 103, picHor.ScaleWidth, picHor.ScaleHeight, picHor.hDC, 0, 0, vbSrcCopy
         If x = Val(txtNoLED.Text) Then                             'if its the last led then draw connecting wires
            'positive wire
            Line3.X1 = 150
            Line3.Y1 = 102
            Line3.X2 = (69 + (50 * x)) + 40
            Line3.Y2 = 102

            If txtNoLED = 1 Then                                         'ground wire placement
               Line4.X1 = 158
               Line4.Y1 = 153
               Line4.X2 = 189
               Line4.Y2 = 153
               Shape3.Left = 187
               Shape3.Top = 150
               Label6.Top = 160
               Label6.Left = 204
             Else
               Line4.X1 = 158
               Line4.Y1 = 153
               Line4.X2 = (100 + (60 * x))
               Line4.Y2 = 153
               Shape3.Left = (100 + (60 * x))
               Shape3.Top = 150
               Label6.Left = (100 + (60 * x))
               Label6.Top = 160
            End If
         End If
      Next x
   End If

   If optLED(2).Value = True Then                                  'series
      If totVD > Val(txtSV.Text) Then                               ' if voltage required by number of led's is greater than supply voltage
         lblResistorValue.Caption = ""
         lblWatt.Caption = ""
         lbltotCur.Caption = ""
         picCir.Cls
         Exit Sub
      End If
      'calculate values
      lblResistorValue.Caption = Format((Val(txtSV.Text) - totVD) / (Val(txtDFC.Text) / 1000), "#," & "###,###")
      lblWatt.Caption = (Val(txtSV.Text) * ((Val(txtDFC.Text)) / 1000))
      lbltotCur.Caption = Val(txtDFC.Text) / 1000
      'draw circuit
      For x = 1 To Val(txtNoLED.Text)
         BitBlt picCir.hDC, 90 + (50 * x), 61, picHor.ScaleWidth, picHor.ScaleHeight, picVert.hDC, 0, 0, vbSrcCopy
         If x = Val(txtNoLED.Text) Then
            Shape3.Left = (140 + 50 * x)
            Shape3.Top = 98
            Label6.Left = (140 + 50 * x)
            Label6.Top = 95
         End If
      Next x
   End If
End Sub

Private Sub optLED_Click(Index As Integer)
   'clear values  for next selection
   lblResistorValue.Caption = ""
   lblWatt.Caption = ""
   lbltotCur.Caption = ""
   'draw and show starting circuit when option is clicked
   picCir.Cls
   If optLED(1).Value = True Then                              'show or hide line4 (used to draw ground on parallel circuits only)
      Line4.Visible = True
      Line4.X1 = 158
      Line4.Y1 = 400
      Line4.X2 = 168
      Line4.Y2 = 400
    Else
      Line4.Visible = False
   End If
   Select Case Index
    Case 0                                                                   'single
      txtNoLED.Visible = False
      Label4.Visible = False
      BitBlt picCir.hDC, 140, 61, picHor.ScaleWidth, picHor.ScaleHeight, picVert.hDC, 0, 0, vbSrcCopy
      lblVolt.Top = 75
      lblVolt.Left = 8
      Shape1.Top = 99
      Shape1.Left = 20
      Shape2.Top = 93
      Shape2.Left = 78
      Line1.X1 = 29
      Line1.X2 = 79
      Line1.Y1 = 103
      Line1.Y2 = 103
      Line2.X1 = 133
      Line2.X2 = 159
      Line2.Y1 = 102
      Line2.Y2 = 102
      Line3.X1 = 190
      Line3.X2 = 194
      Line3.Y1 = 102
      Line3.Y2 = 102
      Shape3.Left = 187
      Shape3.Top = 98
      Label6.Top = 95
      Label6.Left = 204
    Case 1                                                                   'parallel
      txtNoLED.Visible = True
      Label4.Visible = True
      BitBlt picCir.hDC, 150, 103, picHor.ScaleWidth, picHor.ScaleHeight, picHor.hDC, 0, 0, vbSrcCopy
      lblVolt.Top = 75
      lblVolt.Left = 8
      Shape1.Top = 99
      Shape1.Left = 20
      Shape2.Top = 93
      Shape2.Left = 78
      Line1.X1 = 29
      Line1.X2 = 79
      Line1.Y1 = 103
      Line1.Y2 = 103
      Line2.X1 = 133
      Line2.X2 = 159
      Line2.Y1 = 102
      Line2.Y2 = 102
      Line3.X1 = 158
      Line3.X2 = 194
      Line3.Y1 = 153
      Line3.Y2 = 153
      Shape3.Left = 187
      Shape3.Top = 150
      Label6.Top = 160
      Label6.Left = 204
    Case 2                                                                  'series
      txtNoLED.Visible = True
      Label4.Visible = True
      BitBlt picCir.hDC, 140, 61, picHor.ScaleWidth, picHor.ScaleHeight, picVert.hDC, 0, 0, vbSrcCopy
      lblVolt.Top = 75
      lblVolt.Left = 8
      Shape1.Top = 99
      Shape1.Left = 20
      Shape2.Top = 93
      Shape2.Left = 78
      Line1.X1 = 29
      Line1.X2 = 79
      Line1.Y1 = 103
      Line1.Y2 = 103
      Line2.X1 = 133
      Line2.X2 = 159
      Line2.Y1 = 102
      Line2.Y2 = 102
      Line3.X1 = 190
      Line3.X2 = 194
      Line3.Y1 = 102
      Line3.Y2 = 102
      Shape3.Left = 187
      Shape3.Top = 98
      Label6.Top = 95
      Label6.Left = 204
   End Select
End Sub


