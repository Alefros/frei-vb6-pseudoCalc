VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{3BDC7450-30C1-4C34-8F9F-85075AF785F8}#1.0#0"; "XTab.ocx"
Begin VB.Form frm_ajuda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajuda"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   6000
      TabIndex        =   8
      Top             =   6240
      Visible         =   0   'False
      Width           =   2775
      Begin VB.CommandButton cmd_sair 
         Caption         =   "Sair"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         ToolTipText     =   "Sair dos tópicos de ajuda"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmd_ok 
         Caption         =   "Voltar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Voltar para página anterior"
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   8775
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   6120
         Visible         =   0   'False
         Width           =   2775
         Begin VB.CommandButton cmd_proximo 
            Caption         =   "Próximo >"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmd_anterior 
            Caption         =   "< Anterior"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1215
         End
      End
      Begin prjXTab.XTab XTab1 
         Height          =   6015
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   10610
         TabCaption(0)   =   "Menus"
         TabContCtrlCnt(0)=   1
         Tab(0)ContCtrlCap(1)=   "Picture4"
         TabCaption(1)   =   "Números"
         TabContCtrlCnt(1)=   7
         Tab(1)ContCtrlCap(1)=   "Picture14"
         Tab(1)ContCtrlCap(2)=   "Picture13"
         Tab(1)ContCtrlCap(3)=   "Picture11"
         Tab(1)ContCtrlCap(4)=   "Picture10"
         Tab(1)ContCtrlCap(5)=   "Picture9"
         Tab(1)ContCtrlCap(6)=   "Picture8"
         Tab(1)ContCtrlCap(7)=   "Picture7"
         TabCaption(2)   =   "Botões"
         TabContCtrlCnt(2)=   6
         Tab(2)ContCtrlCap(1)=   "Picture19"
         Tab(2)ContCtrlCap(2)=   "Picture18"
         Tab(2)ContCtrlCap(3)=   "Picture17"
         Tab(2)ContCtrlCap(4)=   "Picture16"
         Tab(2)ContCtrlCap(5)=   "Picture15"
         Tab(2)ContCtrlCap(6)=   "Picture12"
         TabTheme        =   3
         ActiveTabBackStartColor=   16316664
         InActiveTabBackStartColor=   15066597
         InActiveTabBackEndColor=   -2147483626
         ActiveTabForeColor=   16711680
         InActiveTabForeColor=   9474192
         BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterBorderColor=   9474192
         DisabledTabBackColor=   -2147483633
         DisabledTabForeColor=   -2147483627
         Begin VB.PictureBox Picture19 
            Height          =   1935
            Left            =   -74880
            ScaleHeight     =   1875
            ScaleWidth      =   8475
            TabIndex        =   47
            Top             =   3840
            Width           =   8535
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               Caption         =   "Botões de comandos (Limpar e utilizar)"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   49
               Top             =   120
               Width           =   8175
            End
            Begin VB.Label Label17 
               Alignment       =   2  'Center
               Caption         =   $"frm_ajuda.frx":0000
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   120
               TabIndex        =   48
               Top             =   480
               Width           =   8295
            End
         End
         Begin VB.PictureBox Picture18 
            Height          =   1095
            Left            =   -74880
            ScaleHeight     =   1035
            ScaleWidth      =   6315
            TabIndex        =   44
            Top             =   2640
            Width           =   6375
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               Caption         =   "Os botões de operação realizam a operação indicada entre os campos primeiro e segundo número."
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   120
               TabIndex        =   46
               Top             =   360
               Width           =   6135
            End
            Begin VB.Label Label15 
               Alignment       =   2  'Center
               Caption         =   "Botões de operações"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   120
               Width           =   6135
            End
         End
         Begin VB.PictureBox Picture17 
            Height          =   495
            Left            =   -70560
            ScaleHeight     =   435
            ScaleWidth      =   4155
            TabIndex        =   42
            Top             =   1800
            Width           =   4215
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               Caption         =   "Botões de comandos: Limpar e utilizar"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   43
               Top             =   0
               Width           =   3975
            End
         End
         Begin VB.PictureBox Picture16 
            Height          =   1095
            Left            =   -68400
            Picture         =   "frm_ajuda.frx":013E
            ScaleHeight     =   1035
            ScaleWidth      =   1995
            TabIndex        =   40
            Top             =   2640
            Width           =   2055
         End
         Begin VB.PictureBox Picture15 
            Height          =   1215
            Left            =   -74880
            Picture         =   "frm_ajuda.frx":21E7
            ScaleHeight     =   1155
            ScaleWidth      =   4155
            TabIndex        =   39
            Top             =   1320
            Width           =   4215
         End
         Begin VB.PictureBox Picture12 
            Height          =   855
            Left            =   -75000
            ScaleHeight     =   795
            ScaleWidth      =   8715
            TabIndex        =   38
            Top             =   360
            Width           =   8775
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               Caption         =   "A Pseudo calculadora possui dois tipos de botões: os botões de operações e os comandos."
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   0
               TabIndex        =   41
               Top             =   0
               Width           =   8535
            End
         End
         Begin VB.PictureBox Picture14 
            Height          =   1815
            Left            =   -74640
            Picture         =   "frm_ajuda.frx":4DED
            ScaleHeight     =   1755
            ScaleWidth      =   3435
            TabIndex        =   37
            Top             =   4080
            Width           =   3495
         End
         Begin VB.PictureBox Picture13 
            Height          =   1815
            Left            =   -70440
            Picture         =   "frm_ajuda.frx":8997
            ScaleHeight     =   1755
            ScaleWidth      =   3555
            TabIndex        =   36
            Top             =   4080
            Width           =   3615
         End
         Begin VB.PictureBox Picture11 
            Height          =   615
            Left            =   -75000
            ScaleHeight     =   555
            ScaleWidth      =   8715
            TabIndex        =   34
            Top             =   3360
            Width           =   8775
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               Caption         =   $"frm_ajuda.frx":D7A3
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   120
               TabIndex        =   35
               Top             =   0
               Width           =   8415
            End
         End
         Begin VB.PictureBox Picture10 
            Height          =   1095
            Left            =   -71640
            Picture         =   "frm_ajuda.frx":D863
            ScaleHeight     =   1035
            ScaleWidth      =   2115
            TabIndex        =   33
            Top             =   2160
            Width           =   2175
         End
         Begin VB.PictureBox Picture9 
            Height          =   975
            Left            =   -68400
            Picture         =   "frm_ajuda.frx":F32E
            ScaleHeight     =   915
            ScaleWidth      =   2115
            TabIndex        =   32
            Top             =   1680
            Width           =   2175
         End
         Begin VB.PictureBox Picture8 
            Height          =   975
            Left            =   -75000
            Picture         =   "frm_ajuda.frx":1146B
            ScaleHeight     =   915
            ScaleWidth      =   2115
            TabIndex        =   31
            Top             =   1680
            Width           =   2175
         End
         Begin VB.PictureBox Picture7 
            Height          =   1335
            Left            =   -75000
            ScaleHeight     =   1275
            ScaleWidth      =   8715
            TabIndex        =   29
            Top             =   360
            Width           =   8775
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               Caption         =   $"frm_ajuda.frx":13540
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1215
               Left            =   120
               TabIndex        =   30
               Top             =   120
               Width           =   8535
            End
         End
         Begin VB.PictureBox Picture4 
            Height          =   5655
            Left            =   0
            ScaleHeight     =   5595
            ScaleWidth      =   8715
            TabIndex        =   15
            Top             =   360
            Width           =   8775
            Begin VB.PictureBox Picture6 
               Height          =   975
               Left            =   6360
               Picture         =   "frm_ajuda.frx":1365D
               ScaleHeight     =   915
               ScaleWidth      =   2115
               TabIndex        =   21
               Top             =   2400
               Width           =   2175
            End
            Begin VB.PictureBox Picture5 
               Height          =   975
               Left            =   240
               Picture         =   "frm_ajuda.frx":15F5E
               ScaleHeight     =   915
               ScaleWidth      =   2115
               TabIndex        =   20
               Top             =   2400
               Width           =   2175
            End
            Begin VB.PictureBox Picture2 
               Height          =   735
               Left            =   0
               ScaleHeight     =   675
               ScaleWidth      =   8715
               TabIndex        =   17
               Top             =   0
               Width           =   8775
               Begin VB.Label Label4 
                  Alignment       =   2  'Center
                  Caption         =   $"frm_ajuda.frx":186D9
                  BeginProperty Font 
                     Name            =   "Calibri"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   615
                  Left            =   120
                  TabIndex        =   18
                  Top             =   120
                  Width           =   8415
               End
            End
            Begin VB.PictureBox Picture3 
               Height          =   735
               Left            =   2280
               Picture         =   "frm_ajuda.frx":18760
               ScaleHeight     =   675
               ScaleWidth      =   4275
               TabIndex        =   16
               Top             =   840
               Width           =   4335
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               Caption         =   "Sair"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   28
               Top             =   4440
               Width           =   8415
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               Caption         =   "Abre a tela principal dos tópicos de ajuda."
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   27
               Top             =   4080
               Width           =   8295
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               Caption         =   "Possibilita ao usuario fechar a Pseudo calculadora."
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   26
               Top             =   4800
               Width           =   8775
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               Caption         =   "O menu sobre, proporciona ao a possibilidade de acessar informações sobre o sistema."
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   25
               Top             =   3480
               Width           =   8535
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               Caption         =   "Ajuda"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   3840
               Width           =   8535
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Caption         =   "Sobre"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   3240
               Width           =   8535
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               Caption         =   $"frm_ajuda.frx":1AC36
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   855
               Left            =   120
               TabIndex        =   22
               Top             =   1800
               Width           =   8535
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               Caption         =   "Modo"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   19
               Top             =   1560
               Width           =   8775
            End
         End
      End
   End
   Begin VB.CommandButton cmd_ajuda 
      Caption         =   "Menus de ajuda"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   5
      ToolTipText     =   "Abrir menus de ajuda"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Como utilizar a pseudo calculadora"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   4815
      Begin VB.CommandButton cmd_voltar 
         Caption         =   "Voltar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         ToolTipText     =   "Voltar para a Pseudo Calculadora"
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2595
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   $"frm_ajuda.frx":1AD0F
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   4575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "_______________________________________________________"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Visão geral sobre a Pseudo Calculadora"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4575
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3480
      OleObjectBlob   =   "frm_ajuda.frx":1ADAD
      Top             =   240
   End
End
Attribute VB_Name = "frm_ajuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()

End Sub

Private Sub cmd_ajuda_Click()
            frm_ajuda.Height = 7740
            frm_ajuda.ScaleHeight = 7275
            frm_ajuda.ScaleWidth = 9016
            frm_ajuda.Width = 9150
            Frame2.Visible = True
            Frame3.Visible = True
            Frame4.Visible = True
            XTab1.Visible = True
            If XTab1.ActiveTab = 0 Then
                cmd_anterior.Enabled = False
            End If
            
End Sub

Private Sub cmd_anterior_Click()
            
            If XTab1.ActiveTab = 2 Then
                XTab1.ActiveTab = 1
                cmd_proximo.Enabled = True
            ElseIf XTab1.ActiveTab = 1 Then
                    XTab1.ActiveTab = 0
                    cmd_anterior.Enabled = False
                    cmd_proximo.Enabled = True
            End If
                    
            
End Sub

Private Sub cmd_ok_Click()
            Frame2.Visible = False
            Frame3.Visible = False
            XTab1.Visible = False
            frm_ajuda.Height = 4125
            frm_ajuda.ScaleHeight = 3660
            frm_ajuda.ScaleWidth = 5085
            frm_ajuda.Width = 5175
            frm_ajuda.Caption = "Ajuda"
            
End Sub

Private Sub cmd_proximo_Click()
            If XTab1.ActiveTab = 0 Then
                XTab1.ActiveTab = 1
                cmd_anterior.Enabled = True
            ElseIf XTab1.ActiveTab = 1 Then
                    XTab1.ActiveTab = 2
                    cmd_proximo.Enabled = False
                    cmd_anterior.Enabled = True
            End If
End Sub

Private Sub cmd_sair_Click()
            Call cmd_voltar_Click
End Sub

Private Sub cmd_voltar_Click()
            Unload Me
            frm_calc.Show
            frm_ajuda.Caption = "Ajuda"
End Sub

Private Sub Form_Load()
            frm_ajuda.Height = 4125
            frm_ajuda.ScaleHeight = 3660
            frm_ajuda.ScaleWidth = 5085
            frm_ajuda.Width = 5175
            frm_ajuda.Caption = "Menus de Ajuda"
            Skin1.LoadSkin App.Path & "\LinuxGnome.skn"
            Skin1.ApplySkin Me.hWnd
End Sub

Private Sub XTab1_Click()
            If XTab1.ActiveTab = 0 Then
                cmd_anterior.Enabled = False
                cmd_proximo.Enabled = True
            ElseIf XTab1.ActiveTab = 1 Then
                cmd_anterior.Enabled = True
                cmd_proximo.Enabled = True
                ElseIf XTab1.ActiveTab = 2 Then
                    cmd_proximo.Enabled = False
                    cmd_anterior.Enabled = True
            End If
End Sub
