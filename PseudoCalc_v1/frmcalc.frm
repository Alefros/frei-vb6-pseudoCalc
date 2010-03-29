VERSION 5.00
Begin VB.Form frmcalc 
   BackColor       =   &H80000013&
   Caption         =   "pseudo Calculador"
   ClientHeight    =   3450
   ClientLeft      =   2625
   ClientTop       =   2205
   ClientWidth     =   9045
   Icon            =   "frmcalc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   9045
   Begin VB.TextBox txtN1 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3960
      TabIndex        =   0
      Text            =   " "
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton btnLimpar 
      BackColor       =   &H80000014&
      Caption         =   "limpar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5880
      Picture         =   "frmcalc.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Ao prescionar este botão, removerá os itens digitados"
      Top             =   2400
      Width           =   2895
   End
   Begin VB.CommandButton btnDividir 
      BackColor       =   &H80000014&
      Caption         =   "dividir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7320
      Picture         =   "frmcalc.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "ao pressionar este botão ocorrerá a divisão entre esses dois números"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton btnMultiplicar 
      BackColor       =   &H80000014&
      Caption         =   "multiplicar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5880
      Picture         =   "frmcalc.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "ao pressionar este botão, ocorrerá a multiplicação entre esses dois números"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton btnSubtrair 
      BackColor       =   &H80000014&
      Caption         =   "subtrair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7320
      Picture         =   "frmcalc.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "ao pressionar este botão, ocorrerá a subtração entre esses dois números"
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton btnSomar 
      BackColor       =   &H80000014&
      Caption         =   "somar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5880
      Picture         =   "frmcalc.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "ao pressionar este botâo, ocorrerá a soma entre esses dois números"
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox txtN2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3960
      TabIndex        =   1
      Text            =   " "
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtResultado 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3000
      TabIndex        =   2
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "o resultado ficou em"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "digite o o segundo número"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   4
      Top             =   1320
      Width           =   3420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      Caption         =   "digite o primeiro número"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   3090
   End
End
Attribute VB_Name = "frmcalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDividir_Click()
            On Error GoTo A
            txtResultado = txtN1 / txtN2
             Exit Sub
A:
            MsgBox "Atenção! algo de errado; divisão por zero. Por favor, verificar", vbExclamation, "vcalc"
            
End Sub

Private Sub btnLimpar_Click()
            txtN1 = Clear
            txtN2 = Clear
            txtResultado = Clear
            txtN1.SetFocus
            
End Sub

Private Sub btnMultiplicar_Click()
             On Error GoTo B
            txtResultado = txtN1 * txtN2
            Exit Sub
B:
             MsgBox "Por favor, digite os valores", vbExclamation
End Sub

Private Sub btnSomar_Click()
            On Error GoTo C
            txtResultado = CCur(txtN1) + CCur(txtN2)
             Exit Sub
C:
             MsgBox "Por favor, digite os valores", vbExclamation
End Sub

Private Sub btnSubtrair_Click()
            On Error GoTo D
            txtResultado = txtN1 - txtN2
            Exit Sub
D:
            MsgBox "por favor, digite os valores", vbExclamation
End Sub

Private Sub dsf()

End Sub

Private Sub txtN1_KeyPress(KeyAscii As Integer)
            If KeyAscii < 48 Then
                If KeyAscii <> 44 Then
                    If KeyAscii <> 45 Then
                        KeyAscii = 8
                    Else
                        If Len(txtN1) >= 1 Then KeyAscii = 0
                    End If
                End If
            End If
            
End Sub

