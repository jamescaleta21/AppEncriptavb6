VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encriptar/Desencriptar"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7065
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Resultado"
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   6855
      Begin VB.TextBox txtResultado 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   6615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Desencriptar"
      Height          =   1815
      Left            =   3600
      TabIndex        =   10
      Top             =   960
      Width           =   3375
      Begin VB.CommandButton cmdDesencriptar 
         Caption         =   "Desencriptar"
         Height          =   480
         Left            =   1920
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtCadena2 
         Height          =   375
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadena:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ingresar Semilla"
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6855
      Begin VB.TextBox txtSemilla 
         Height          =   375
         Left            =   1800
         TabIndex        =   0
         Text            =   "jamescaleta"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semilla:"
         Height          =   195
         Left            =   960
         TabIndex        =   9
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Encriptar"
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   3375
      Begin VB.TextBox txtCadena1 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   3135
      End
      Begin VB.CommandButton cmdEncriptar 
         Caption         =   "Encriptar"
         Height          =   480
         Left            =   1920
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cadena:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDesencriptar_Click()
Dim strSemilla As String
strSemilla = Semilla(Me.txtSemilla.Text)

Dim strResultado As String
strResultado = DeCodificar(Me.txtCadena2.Text, strSemilla)

Me.txtResultado.Text = strResultado
End Sub

Private Sub cmdEncriptar_Click()
Dim strSemilla  As String
strSemilla = Semilla(Me.txtSemilla.Text)

Dim strResultado As String

strResultado = Codificar(Me.txtCadena1.Text, strSemilla)

Me.txtCadena2.Text = strResultado

End Sub
