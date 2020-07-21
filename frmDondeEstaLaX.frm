VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmDondeEstaLaX 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Juego Donde está la X entre medio de tantas Y v1.0"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6945
   Icon            =   "frmDondeEstaLaX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   6945
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cd 
      Left            =   3000
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picColor 
      Height          =   255
      Left            =   840
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   6000
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Nueva Partida"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5880
      TabIndex        =   5
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Casi Imposible"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4080
      TabIndex        =   4
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Abanzado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2760
      TabIndex        =   3
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Medio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1440
      TabIndex        =   2
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Facíl"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Columns         =   1000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   1.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5130
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
   Begin VB.Label lblc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cordenadas:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2760
      TabIndex        =   9
      Top             =   6000
      Width           =   1065
   End
   Begin VB.Label lblp 
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos:0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "color:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   6000
      Width           =   555
   End
End
Attribute VB_Name = "frmDondeEstaLaX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'X
'X
'X Juego Donde está la X entre medio de tantas Y v1.0 :)!.
'X Autor: Martin Grasso
'X Dato: 05/11/2017
'X
'X
Dim X, Vida, dificultad, puntos As Integer
Const Juego As String = _
"Juego Donde está la X entre medio de tantas Y v1.0"

Private Sub Command1_Click()
 crearJuego 7
 dificultad = 7
 Dificultad_Activa False
End Sub

Private Sub Command2_Click()
 crearJuego 5
 dificultad = 5
 Dificultad_Activa False
End Sub

Private Sub Command3_Click()
 crearJuego 4
 dificultad = 4
 Dificultad_Activa False
End Sub

Private Sub Command4_Click()
 crearJuego 1
 dificultad = 1
 Dificultad_Activa False
End Sub

Private Sub crearJuego(ByVal tx As Byte)
Dim xS As Integer
With List1
.Clear
.FontSize = CInt(tx)
End With
   For xS = 0 To 10007
       List1.AddItem "Y"
   Next xS                   '
    X = Int(10007 * Rnd) + 1 ' Posissionar X!.
        List1.List(X) = "X"  '
End Sub

Private Sub Command5_Click()
Form_Load
Dificultad_Activa True
End Sub

Private Sub Form_Load()
         '
Vida = 7 'cargar Vidas
         '
visBoton False
picColor.BackColor = List1.ForeColor
         '
Me.Caption = Juego
         '
Dificultad_Activa False
End Sub

Private Sub List1_Click()
Dim X1, Y1, Z1 As Byte
visBoton False
If List1.ListIndex = X And Vida > 0 Then
   MsgBox " Ganaste", vbInformation, Juego
   puntos = puntos + 1 ' le sumo un punto al Puntaje Actual
   crearJuego dificultad
   X1 = Int(255 * Rnd) + 1
   Y1 = Int(255 * Rnd) + 1
   Z1 = Int(255 * Rnd) + 1
   List1.ForeColor = RGB(X1, Y1, Z1)
   lblp.Caption = "Puntos:" & puntos
   ElseIf Vida >= 1 Then
   MsgBox "Prediste te resto una Vida Vidas restantes:" & Vida, vbInformation, Juego
   Vida = Vida - 1 'le resto una Vida
   Else
   MsgBox "Fin del Juego", vbInformation, Juego
   puntos = 0 ' elimino el puntaje
   lblp.Caption = "Puntos:0"
   List1.Clear
   visBoton True
End If
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblc.Caption = "X=" & X & "." & "|" & "Y=" & Y & "."
lblc.Refresh
End Sub

Private Sub picColor_Click()
With cd
     .ShowColor
     If Not .Color = 0 Then
       List1.ForeColor = .Color
       picColor.BackColor = List1.ForeColor
     End If
End With
End Sub

Private Sub visBoton(ByVal b As Boolean)
Command5.Enabled = b
End Sub

Private Sub Dificultad_Activa(ByVal X As Boolean)
 Command1.Enabled = X
 Command2.Enabled = X
 Command3.Enabled = X
 Command4.Enabled = X
 visBoton True
End Sub
