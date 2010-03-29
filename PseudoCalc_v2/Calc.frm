VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form frm_calc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pseudo Calculadora"
   ClientHeight    =   2835
   ClientLeft      =   150
   ClientTop       =   525
   ClientWidth     =   4155
   Icon            =   "Calc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   4155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_limpar 
      Caption         =   "Limpar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Picture         =   "Calc.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Limpa todos os campos"
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmd_utilizar 
      Caption         =   "Utilizar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      Picture         =   "Calc.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Utilizar o resultado como primeiro número"
      Top             =   2040
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2640
      OleObjectBlob   =   "Calc.frx":0B8E
      Top             =   720
   End
   Begin VB.Frame Frame4 
      Caption         =   "Operações"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   3975
      Begin VB.CommandButton cmd_divisao 
         Height          =   615
         Left            =   3240
         Picture         =   "Calc.frx":0DC2
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Realiza o calculo de divisão"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmd_multiplicacao 
         Height          =   615
         Left            =   2160
         Picture         =   "Calc.frx":1204
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Realiza o calculo de multiplicação"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmd_subtracao 
         Height          =   615
         Left            =   1200
         Picture         =   "Calc.frx":1646
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Realiza o calculo de subtração"
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmd_adicao 
         Height          =   615
         Left            =   240
         Picture         =   "Calc.frx":1A88
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Realiza o calculo de adição"
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Segundo número"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      TabIndex        =   9
      Top             =   0
      Width           =   1935
      Begin VB.TextBox txt_num2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Digital-7"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Primeiro número"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   1935
      Begin VB.TextBox txt_num1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Digital-7"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   10
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resultado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
      Begin VB.TextBox txt_resultado 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Digital-7"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Menu mnu_modo 
      Caption         =   "&Modo"
      Begin VB.Menu mnu_comum 
         Caption         =   "Comum"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu_monetario 
         Caption         =   "Monetário"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mnu_sobre 
      Caption         =   "&Sobre"
   End
   Begin VB.Menu mnu_ajuda 
      Caption         =   "&Ajuda"
      Begin VB.Menu mnu_to 
         Caption         =   "Tópicos de ajuda"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnu_sair 
      Caption         =   "Sai&r"
   End
End
Attribute VB_Name = "frm_calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim virgula, virgula2 As Integer
Option Explicit
Private Sub cmd_adicao_Click()
            If txt_num1 = Empty Then
                MsgBox "Por favor, digite o primeiro valor!", vbInformation, "Pseudo-calculadora"
                Call cmd_limpar_Click
                Exit Sub
            ElseIf txt_num2 = Empty Then
                MsgBox "Por favor, digite o segundo valor!", vbInformation, "Pseudo-calculadora"
                Call cmd_limpar_Click
                Exit Sub
            End If
            txt_resultado = CCur(txt_num1) + CCur(txt_num2)
            cmd_utilizar.SetFocus
            Exit Sub
            Call cmd_limpar_Click
End Sub

Private Sub cmd_divisao_Click()
            If txt_num1 = Empty Then
                MsgBox "Por favor, digite o primeiro valor!", vbInformation, "Pseudo-calculadora"
                Call cmd_limpar_Click
                Exit Sub
            ElseIf txt_num2 = Empty Then
                MsgBox "Por favor, digite o segundo valor!", vbInformation, "Pseudo-calculadora"
                Call cmd_limpar_Click
                Exit Sub
            End If
            If txt_num2 = "0" Then
                MsgBox "Não se pode dividir por zero", vbInformation, "Pseudo-calculadora"
                txt_num2 = Empty
                txt_num2.SetFocus
                Exit Sub
            End If
            txt_resultado = CCur(txt_num1) / CCur(txt_num2)
            cmd_utilizar.SetFocus
            Exit Sub
            Call cmd_limpar_Click
End Sub

Private Sub cmd_limpar_Click()
            txt_num1.Text = Empty
            txt_num2.Text = Empty
            txt_resultado.Text = Empty
            txt_num1.SetFocus
            virgula = "0"
            virgula2 = "0"
            cmd_utilizar.Enabled = False
End Sub

Private Sub cmd_multiplicacao_Click()
            If txt_num1 = Empty Then
                MsgBox "Por favor, digite o primeiro valor!", vbInformation, "Pseudo-calculadora"
                Call cmd_limpar_Click
                Exit Sub
            ElseIf txt_num2 = Empty Then
                MsgBox "Por favor, digite o segundo valor!", vbInformation, "Pseudo-calculadora"
                Call cmd_limpar_Click
                Exit Sub
            End If
            txt_resultado = CCur(txt_num1) * CCur(txt_num2)
            cmd_utilizar.SetFocus
            Exit Sub
            Call cmd_limpar_Click
End Sub

Private Sub cmd_subtracao_Click()
            If txt_num1 = Empty Then
                MsgBox "Por favor, digite o primeiro valor!", vbInformation, "Pseudo-calculadora"
                Call cmd_limpar_Click
                Exit Sub
            ElseIf txt_num2 = Empty Then
                MsgBox "Por favor, digite o segundo valor!", vbInformation, "Pseudo-calculadora"
                Call cmd_limpar_Click
                Exit Sub
            End If
            txt_resultado = CCur(txt_num1) - CCur(txt_num2)
            cmd_utilizar.SetFocus
            Exit Sub
            Call cmd_limpar_Click
End Sub

Private Sub cmd_utilizar_Click()
            txt_num1.Text = Empty
            txt_num2.Text = Empty
            txt_num1.Text = txt_resultado.Text
            txt_resultado.Text = Empty
            txt_num2.SetFocus
End Sub

Private Sub virgula_Click()
            'txt_num1 = txt_num1 & "," +
            
End Sub

Private Sub Form_Load()
            virgula = "0"
            virgula2 = "0"
            modo = "Comum"
            frm_calc.mnu_comum.Caption = "• Comum"
             frm_calc.mnu_monetario.Caption = "Monetário"
            
            Skin1.LoadSkin App.Path & "\LinuxGnome.skn"
            Skin1.ApplySkin Me.hWnd
End Sub

Private Sub mnu_comum_Click()
            modo = "Comum"
            frm_calc.mnu_comum.Caption = "• Comum"
             frm_calc.mnu_monetario.Caption = "Monetário"
End Sub

Private Sub mnu_monetario_Click()
            modo = "Money"
            frm_calc.mnu_monetario.Caption = "• Monetário"
            frm_calc.mnu_comum.Caption = "Comum"
            'If txt_resultado <> Empty Then
             '   txt_resultado = Format(txt_resultado, "currency")
            
            'End If
                
End Sub
Private Sub mnu_sair_Click()
            End
End Sub
Private Sub mnu_sobre_Click()
            Unload Me
            frm_sobre.Show
End Sub
Private Sub mnu_to_Click()
            frm_ajuda.Show
            Unload Me
End Sub
Private Sub txt_num1_KeyPress(KeyAscii As Integer)
           If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{tab}" 'ENTER virar TAB
           If KeyAscii = vbKeyEscape Then KeyAscii = 0: SendKeys "+{tab}" 'Esc voltar
              
              
            If KeyAscii = 44 Then
                If virgula = "1" Then
                    KeyAscii = 0
                    GoTo a:
                End If
                virgula = "1"
                KeyAscii = 44
a:
            Else
            If KeyAscii = 8 Then
                KeyAscii = 8
            ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
                KeyAscii = 0
'                KeyAscii = 8
            End If
            End If
End Sub

Private Sub txt_num1_LostFocus()
            If modo = "Money" Then
            txt_num1 = Format(txt_num1, "currency")
            End If
End Sub

Private Sub txt_num2_KeyPress(KeyAscii As Integer)
            If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{tab}" 'ENTER virar TAB
            If KeyAscii = vbKeyEscape Then KeyAscii = 0: SendKeys "+{tab}" 'Esc voltar
            
         If KeyAscii = 44 Then
                If virgula2 = "1" Then
                    KeyAscii = 0
                    GoTo a:
                End If
                virgula2 = "1"
                KeyAscii = 44
a:
            Else
            If KeyAscii = 8 Then
                KeyAscii = 8
            ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
                KeyAscii = 0
'                KeyAscii = 8
            End If
            End If
End Sub

Private Sub txt_num2_LostFocus()
            If modo = "Money" Then
            txt_num2 = Format(txt_num2, "currency")
            End If
End Sub

Private Sub txt_resultado_Change()
            If txt_resultado <> Empty Then
                cmd_utilizar.Enabled = True
            End If
            If modo = "Money" Then
                txt_resultado = Format(txt_resultado, "currency")
            End If
            
End Sub

Private Sub txt_resultado_KeyPress(KeyAscii As Integer)
            If KeyAscii = vbKeyReturn Then KeyAscii = 0: SendKeys "{tab}" 'ENTER virar TAB
           If KeyAscii = vbKeyEscape Then KeyAscii = 0: SendKeys "+{tab}" 'Esc voltar
End Sub
