VERSION 5.00
Object = "{A4B55B03-8129-101D-836D-3E0683BCA07A}#1.0#0"; "TEXT50S.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{604A59D5-2409-101D-97D5-C6626B63EF2D}#1.0#0"; "NUM50S.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{FE1D09E3-6FC7-101D-836D-3E0683BCA07A}#1.0#0"; "DATE50S.OCX"
Begin VB.Form ADF010 
   Caption         =   "�L���b�g�n���h�ڋq�Ǘ�"
   ClientHeight    =   13530
   ClientLeft      =   2700
   ClientTop       =   915
   ClientWidth     =   15900
   BeginProperty Font 
      Name            =   "�l�r �o�S�V�b�N"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   13530
   ScaleWidth      =   15900
   Begin FPSpread.vaSpread va�������X�g 
      Height          =   4215
      Left            =   9000
      OleObjectBlob   =   "ADF010.frx":0000
      TabIndex        =   22
      Top             =   1560
      Width           =   6495
   End
   Begin FPSpread.vaSpread va�ڋq���X�g 
      Height          =   4215
      Left            =   240
      OleObjectBlob   =   "ADF010.frx":360A
      TabIndex        =   21
      Top             =   1560
      Width           =   8775
   End
   Begin FPSpread.vaSpread va�����f���� 
      Height          =   1095
      Left            =   9000
      OleObjectBlob   =   "ADF010.frx":4B2E
      TabIndex        =   126
      Top             =   120
      Width           =   6375
   End
   Begin VB.CommandButton cmd�ڍs 
      Caption         =   "�ڍs"
      Height          =   375
      Left            =   5400
      TabIndex        =   129
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmd�⍇�ԍ� 
      Caption         =   "�⍇�ԍ�"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   128
      Top             =   12840
      Width           =   1095
   End
   Begin VB.CommandButton cmd�X�V 
      Caption         =   "�X�V"
      Height          =   375
      Left            =   14400
      TabIndex        =   127
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdTOOL 
      Caption         =   "TOOL"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   25
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmd�ʃ��[�� 
      Caption         =   "�ʃ��[��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11280
      TabIndex        =   116
      Top             =   12840
      Width           =   1095
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   14760
      TabIndex        =   119
      Top             =   12840
      Width           =   1095
   End
   Begin VB.CommandButton cmd�ꊇ���� 
      Caption         =   "�ꊇ����"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   113
      Top             =   12840
      Width           =   1095
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����o��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13680
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   118
      Top             =   12840
      Width           =   1095
   End
   Begin VB.CommandButton cmd�I�[�g�V�b�v 
      Cancel          =   -1  'True
      Caption         =   "�I�[�g�V�b�v"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   112
      Top             =   12840
      Width           =   1335
   End
   Begin VB.CommandButton cmd�R�����C�t 
      Caption         =   "�R�����C�t"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
   Begin VB.Frame frm���[������ 
      Height          =   5535
      Left            =   600
      TabIndex        =   104
      Top             =   7080
      Width           =   14775
      Begin FPSpread.vaSpread va���[������ 
         Height          =   4935
         Left            =   480
         OleObjectBlob   =   "ADF010.frx":519C
         TabIndex        =   105
         Top             =   240
         Width           =   8895
      End
      Begin ImTextCtrl.ImText txt���[���{�� 
         Height          =   4815
         Left            =   9600
         TabIndex        =   106
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   8493
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   4
         InsertMode      =   -1  'True
         MaxLength       =   0
         MultiLine       =   -1  'True
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   2
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   ""
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   0
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5497
         MousePointer    =   0
      End
   End
   Begin VB.CommandButton cmd�A�[�f���w���� 
      Caption         =   "�A�[�f���w����"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   18
      Top             =   5880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd�����}�K���s 
      Caption         =   "�����}�K"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   28
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Frame frm���� 
      Height          =   5535
      Left            =   600
      TabIndex        =   58
      Top             =   7080
      Width           =   14775
      Begin VB.CheckBox chk����1_1 
         Caption         =   "���ʓI���p"
         Height          =   255
         Left            =   7080
         TabIndex        =   125
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CheckBox chk����2_1 
         Caption         =   "�����̐Ϗd"
         Height          =   255
         Left            =   8640
         TabIndex        =   124
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CheckBox chk����3_1 
         Caption         =   "�c�u�c"
         Height          =   255
         Left            =   10320
         TabIndex        =   123
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CheckBox chk����4_1 
         Caption         =   "�^���ƈ��"
         Height          =   255
         Left            =   11400
         TabIndex        =   122
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CheckBox chk����5_1 
         Caption         =   "�閧�̈��"
         Height          =   255
         Left            =   13080
         TabIndex        =   121
         Top             =   5160
         Width           =   1455
      End
      Begin VB.ComboBox cmb���� 
         Height          =   345
         IMEMode         =   4  '�S�p�Ђ炪��
         Left            =   4560
         TabIndex        =   66
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox cmb��s 
         Height          =   345
         IMEMode         =   4  '�S�p�Ђ炪��
         Left            =   4560
         TabIndex        =   70
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmd�v�Z 
         Caption         =   "�R�����C�t�v�Z"
         Height          =   375
         Left            =   6840
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CommandButton cmd�e���v���[�g 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13560
         TabIndex        =   102
         Top             =   2280
         Width           =   735
      End
      Begin VB.ComboBox cmb�e���v���[�g 
         Height          =   345
         Left            =   8880
         TabIndex        =   101
         Top             =   2280
         Width           =   4575
      End
      Begin ImTextCtrl.ImText txt����URL 
         Height          =   375
         Left            =   240
         TabIndex        =   81
         Top             =   4680
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   0   'False
         IMEMode         =   0
         InsertMode      =   -1  'True
         MaxLength       =   255
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "����ID"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   72
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":54B3
         MousePointer    =   0
      End
      Begin VB.CommandButton cmd����5 
         Caption         =   "�艿"
         Height          =   375
         Left            =   12840
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   1320
         Width           =   735
      End
      Begin ImDateCtrl.ImDate txt�o�ח\��� 
         Height          =   375
         Left            =   11760
         TabIndex        =   100
         Top             =   1800
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   65537
         HolidayColor    =   255
         AlignHorizontal =   0
         AlignVertical   =   0
         CursorPosition  =   0
         MaxDate         =   20991231
         MinDate         =   18680908
         Number          =   20100717
         Value           =   40376
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         ClipMode        =   0
         HighlightText   =   0
         DataProperty    =   0
         FirstMonth      =   4
         Format          =   "yyyy/mm/dd"
         DisplayFormat   =   "yyyy/mm/dd"
         NationalHolidays=   "1/1"
         UserHolidays    =   ""
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyNow          =   "{F3}"
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         PromptChar      =   "_"
         Text            =   "2010/07/17"
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�o�ח\���"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   88
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":54CF
         MousePointer    =   0
      End
      Begin VB.CommandButton cmd����4 
         Caption         =   "-20%"
         Height          =   375
         Left            =   11880
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmd����3 
         Caption         =   "-10%"
         Height          =   375
         Left            =   10920
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmd����2 
         Caption         =   "-4600"
         Height          =   375
         Left            =   9840
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txt���v���z 
         Alignment       =   1  '�E����
         Enabled         =   0   'False
         Height          =   375
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   96
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   3720
         Width           =   1935
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "-770"
         Height          =   375
         Left            =   8880
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmd�{��2 
         Caption         =   "�{��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmd�{��1 
         Caption         =   "�{��"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   8.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   2760
         Width           =   495
      End
      Begin ImDateCtrl.ImDate txt������ 
         Height          =   375
         Left            =   5760
         TabIndex        =   82
         Top             =   360
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   65537
         HolidayColor    =   255
         AlignHorizontal =   0
         AlignVertical   =   0
         CursorPosition  =   0
         MaxDate         =   20991231
         MinDate         =   18680908
         Number          =   20081103
         Value           =   39755
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         ClipMode        =   0
         HighlightText   =   0
         DataProperty    =   0
         FirstMonth      =   4
         Format          =   "yyyy/mm/dd"
         DisplayFormat   =   "yyyy/mm/dd"
         NationalHolidays=   "1/1"
         UserHolidays    =   ""
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyNow          =   "{F3}"
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         PromptChar      =   "_"
         Text            =   "2008/11/03"
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "������"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":54EB
         MousePointer    =   0
      End
      Begin VB.ComboBox cmb������ 
         Height          =   345
         IMEMode         =   4  '�S�p�Ђ炪��
         ItemData        =   "ADF010.frx":5507
         Left            =   3240
         List            =   "ADF010.frx":550E
         TabIndex        =   63
         Top             =   840
         Width           =   2535
      End
      Begin ImTextCtrl.ImText txt����ID 
         Height          =   375
         Left            =   480
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   0   'False
         IMEMode         =   0
         InsertMode      =   -1  'True
         MaxLength       =   0
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   0   'False
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "����ID"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5518
         MousePointer    =   0
      End
      Begin VB.ComboBox cmb��z�Ǝ� 
         Height          =   345
         IMEMode         =   4  '�S�p�Ђ炪��
         Left            =   3120
         TabIndex        =   75
         Top             =   2760
         Width           =   1935
      End
      Begin ImTextCtrl.ImText txt���[�����M 
         Height          =   375
         Left            =   8760
         TabIndex        =   97
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   0   'False
         IMEMode         =   4
         InsertMode      =   -1  'True
         MaxLength       =   100
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   0   'False
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "���[�����M"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   80
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5534
         MousePointer    =   0
      End
      Begin ImTextCtrl.ImText txt���l2 
         Height          =   2295
         Left            =   8880
         TabIndex        =   103
         Top             =   2640
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   4048
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   4
         InsertMode      =   -1  'True
         MaxLength       =   0
         MultiLine       =   -1  'True
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   2
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "���l"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   0
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5550
         MousePointer    =   0
      End
      Begin ImTextCtrl.ImText txt�⍇�ԍ� 
         Height          =   375
         Left            =   240
         TabIndex        =   77
         Top             =   3720
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   2
         InsertMode      =   -1  'True
         MaxLength       =   30
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�⍇�ԍ�"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   72
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":556C
         MousePointer    =   0
      End
      Begin ImTextCtrl.ImText txt�x���ԍ� 
         Height          =   375
         Left            =   240
         TabIndex        =   76
         Top             =   3240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   2
         InsertMode      =   -1  'True
         MaxLength       =   30
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�x���ԍ�"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   72
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5588
         MousePointer    =   0
      End
      Begin ImDateCtrl.ImDate txt�o�ד� 
         Height          =   375
         Left            =   480
         TabIndex        =   73
         Top             =   2760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   65537
         HolidayColor    =   255
         AlignHorizontal =   0
         AlignVertical   =   0
         CursorPosition  =   0
         MaxDate         =   20991231
         MinDate         =   18680908
         Number          =   20081026
         Value           =   39747
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         ClipMode        =   0
         HighlightText   =   0
         DataProperty    =   0
         FirstMonth      =   4
         Format          =   "yyyy/mm/dd"
         DisplayFormat   =   "yyyy/mm/dd"
         NationalHolidays=   "1/1"
         UserHolidays    =   ""
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyNow          =   "{F3}"
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         PromptChar      =   "_"
         Text            =   "2008/10/26"
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�o�ד�"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":55A4
         MousePointer    =   0
      End
      Begin ImNumberCtrl.ImNumber txt���̑��萔�� 
         Height          =   375
         Left            =   5280
         TabIndex        =   94
         Top             =   3240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   1
         ClipMode        =   0
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         HighlightText   =   -1  'True
         ZeroAllowed     =   -1  'True
         MinusColor      =   255
         MaxValue        =   99999
         MinValue        =   -99999
         Value           =   0
         SelStart        =   1
         SelLength       =   0
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         KeyThreeZero    =   ""
         SepDecimal      =   "."
         SepThousand     =   ","
         Text            =   "0"
         Format          =   "####0"
         DisplayFormat   =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�|�C���g���p"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   88
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":55C0
         MousePointer    =   0
      End
      Begin ImNumberCtrl.ImNumber txt�ԋ� 
         Height          =   375
         Left            =   6000
         TabIndex        =   93
         Top             =   2760
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   1
         ClipMode        =   0
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         HighlightText   =   -1  'True
         ZeroAllowed     =   -1  'True
         MinusColor      =   255
         MaxValue        =   99999
         MinValue        =   -99999
         Value           =   0
         SelStart        =   1
         SelLength       =   0
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         KeyThreeZero    =   ""
         SepDecimal      =   "."
         SepThousand     =   ","
         Text            =   "0"
         Format          =   "####0"
         DisplayFormat   =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   0   'False
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�ԋ�"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   40
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":55DC
         MousePointer    =   0
      End
      Begin ImNumberCtrl.ImNumber txt���� 
         Height          =   375
         Left            =   6000
         TabIndex        =   92
         Top             =   2280
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   1
         ClipMode        =   0
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         HighlightText   =   -1  'True
         ZeroAllowed     =   -1  'True
         MinusColor      =   255
         MaxValue        =   99999
         MinValue        =   -99999
         Value           =   0
         SelStart        =   1
         SelLength       =   0
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         KeyThreeZero    =   ""
         SepDecimal      =   "."
         SepThousand     =   ","
         Text            =   "0"
         Format          =   "####0"
         DisplayFormat   =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "����"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   40
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":55F8
         MousePointer    =   0
      End
      Begin ImNumberCtrl.ImNumber txt���� 
         Height          =   375
         Left            =   6000
         TabIndex        =   91
         Top             =   1800
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   1
         ClipMode        =   0
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         HighlightText   =   -1  'True
         ZeroAllowed     =   -1  'True
         MinusColor      =   255
         MaxValue        =   300
         MinValue        =   1
         Value           =   1
         SelStart        =   1
         SelLength       =   0
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         KeyThreeZero    =   ""
         SepDecimal      =   "."
         SepThousand     =   ","
         Text            =   "1"
         Format          =   "####0"
         DisplayFormat   =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "����"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   40
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5614
         MousePointer    =   0
      End
      Begin ImNumberCtrl.ImNumber txt���� 
         Height          =   375
         Left            =   6000
         TabIndex        =   85
         Top             =   1320
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   1
         ClipMode        =   0
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         HighlightText   =   -1  'True
         ZeroAllowed     =   -1  'True
         MinusColor      =   255
         MaxValue        =   99999
         MinValue        =   -99999
         Value           =   0
         SelStart        =   1
         SelLength       =   0
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         KeyThreeZero    =   ""
         SepDecimal      =   "."
         SepThousand     =   ","
         Text            =   "0"
         Format          =   "####0"
         DisplayFormat   =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "����"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   40
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5630
         MousePointer    =   0
      End
      Begin ImNumberCtrl.ImNumber txt�P�� 
         Height          =   375
         Left            =   6000
         TabIndex        =   84
         Top             =   840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   1
         ClipMode        =   0
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         HighlightText   =   -1  'True
         ZeroAllowed     =   -1  'True
         MinusColor      =   255
         MaxValue        =   99999
         MinValue        =   -99999
         Value           =   0
         SelStart        =   1
         SelLength       =   0
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         KeyThreeZero    =   ""
         SepDecimal      =   "."
         SepThousand     =   ","
         Text            =   "0"
         Format          =   "####0"
         DisplayFormat   =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�P��"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   40
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":564C
         MousePointer    =   0
      End
      Begin ImTextCtrl.ImText txt�z�B���� 
         Height          =   375
         Left            =   240
         TabIndex        =   71
         Top             =   2280
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   4
         InsertMode      =   -1  'True
         MaxLength       =   100
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�z�B����"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   72
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5668
         MousePointer    =   0
      End
      Begin VB.ComboBox cmb�������@ 
         Height          =   345
         IMEMode         =   4  '�S�p�Ђ炪��
         Left            =   1320
         TabIndex        =   68
         Top             =   1800
         Width           =   2535
      End
      Begin VB.ComboBox cmb���i�� 
         Height          =   345
         IMEMode         =   4  '�S�p�Ђ炪��
         Left            =   1320
         TabIndex        =   65
         Top             =   1320
         Width           =   3255
      End
      Begin VB.ComboBox cmb�X�e�[�^�X 
         Height          =   345
         IMEMode         =   4  '�S�p�Ђ炪��
         ItemData        =   "ADF010.frx":5684
         Left            =   1320
         List            =   "ADF010.frx":568B
         TabIndex        =   62
         Top             =   840
         Width           =   1815
      End
      Begin ImDateCtrl.ImDate txt�󒍓� 
         Height          =   375
         Left            =   2760
         TabIndex        =   60
         Top             =   360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   65537
         HolidayColor    =   255
         AlignHorizontal =   0
         AlignVertical   =   0
         CursorPosition  =   0
         MaxDate         =   20991231
         MinDate         =   18680908
         Number          =   20081026
         Value           =   39747
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         ClipMode        =   0
         HighlightText   =   0
         DataProperty    =   0
         FirstMonth      =   4
         Format          =   "yyyy/mm/dd"
         DisplayFormat   =   "yyyy/mm/dd"
         NationalHolidays=   "1/1"
         UserHolidays    =   ""
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyNow          =   "{F3}"
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         PromptChar      =   "_"
         Text            =   "2008/10/26"
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�󒍓�"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5695
         MousePointer    =   0
      End
      Begin ImTextCtrl.ImText txt�����ԍ� 
         Height          =   375
         Left            =   9000
         TabIndex        =   98
         Top             =   720
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   2
         InsertMode      =   -1  'True
         MaxLength       =   50
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�����ԍ�"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   64
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":56B1
         MousePointer    =   0
      End
      Begin ImTextCtrl.ImText txt�R�����C�t 
         Height          =   375
         Left            =   8880
         TabIndex        =   99
         Top             =   1800
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   2
         InsertMode      =   -1  'True
         MaxLength       =   30
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "��ײ�NO"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   72
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":56CD
         MousePointer    =   0
      End
      Begin ImNumberCtrl.ImNumber txt�ב��^�� 
         Height          =   375
         Left            =   3480
         TabIndex        =   79
         Top             =   4200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   1
         ClipMode        =   0
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         HighlightText   =   -1  'True
         ZeroAllowed     =   -1  'True
         MinusColor      =   255
         MaxValue        =   99999
         MinValue        =   -99999
         Value           =   0
         SelStart        =   1
         SelLength       =   0
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         KeyThreeZero    =   ""
         SepDecimal      =   "."
         SepThousand     =   ","
         Text            =   "0"
         Format          =   "####0"
         DisplayFormat   =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   0   'False
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�ב��^��"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   64
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":56E9
         MousePointer    =   0
      End
      Begin ImNumberCtrl.ImNumber txt�d�����z 
         Height          =   375
         Left            =   240
         TabIndex        =   78
         Top             =   4200
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   1
         ClipMode        =   0
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         HighlightText   =   -1  'True
         ZeroAllowed     =   -1  'True
         MinusColor      =   255
         MaxValue        =   99999
         MinValue        =   -99999
         Value           =   0
         SelStart        =   1
         SelLength       =   0
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         KeyThreeZero    =   ""
         SepDecimal      =   "."
         SepThousand     =   ","
         Text            =   "0"
         Format          =   "####0"
         DisplayFormat   =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   0   'False
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�d�����z"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   72
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5705
         MousePointer    =   0
      End
      Begin ImTextCtrl.ImText txt�z�B����2 
         Height          =   375
         Left            =   3480
         TabIndex        =   72
         Top             =   2280
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   4
         InsertMode      =   -1  'True
         MaxLength       =   100
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   ""
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   0
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5721
         MousePointer    =   0
      End
      Begin VB.Label lbl��s 
         Caption         =   "��s"
         Height          =   375
         Left            =   3960
         TabIndex        =   69
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lb���v���z 
         Caption         =   "���v���z"
         Height          =   375
         Left            =   5550
         TabIndex        =   95
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label lb�������@ 
         Caption         =   "�������@"
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lb���i�� 
         Caption         =   "���i��"
         Height          =   255
         Left            =   480
         TabIndex        =   64
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lb�X�e�[�^�X 
         Caption         =   "�X�e�[�^�X"
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame frm�ڋq 
      Height          =   5415
      Left            =   600
      TabIndex        =   30
      Top             =   7200
      Width           =   14775
      Begin VB.CommandButton cmd�X�֔ԍ� 
         Caption         =   "��"
         Height          =   375
         Left            =   2880
         TabIndex        =   0
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox chk����5 
         Caption         =   "�閧�̈��"
         Height          =   255
         Left            =   13320
         TabIndex        =   55
         Top             =   4800
         Width           =   1455
      End
      Begin VB.CheckBox chk����4 
         Caption         =   "�^���ƈ��"
         Height          =   255
         Left            =   11640
         TabIndex        =   54
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CheckBox chk����3 
         Caption         =   "�c�u�c"
         Height          =   255
         Left            =   10560
         TabIndex        =   53
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CheckBox chk����2 
         Caption         =   "�����̐Ϗd"
         Height          =   255
         Left            =   8880
         TabIndex        =   52
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CheckBox chk����1 
         Caption         =   "���ʓI���p"
         Height          =   255
         Left            =   7320
         TabIndex        =   51
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton cmd�]�� 
         Caption         =   "�]��"
         Height          =   375
         Left            =   5520
         TabIndex        =   46
         Top             =   4800
         Width           =   1095
      End
      Begin ImDateCtrl.ImDate txt�a���� 
         Height          =   375
         Left            =   2160
         TabIndex        =   45
         Top             =   4800
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   65537
         HolidayColor    =   255
         AlignHorizontal =   0
         AlignVertical   =   0
         CursorPosition  =   0
         MaxDate         =   20991231
         MinDate         =   18680908
         Number          =   20100717
         Value           =   40376
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         ClipMode        =   0
         HighlightText   =   0
         DataProperty    =   0
         FirstMonth      =   4
         Format          =   "yyyy/mm/dd"
         DisplayFormat   =   "yyyy/mm/dd"
         NationalHolidays=   "1/1"
         UserHolidays    =   ""
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyNow          =   "{F3}"
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         PromptChar      =   "_"
         Text            =   "2010/07/17"
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�a����"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   64
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":573D
         MousePointer    =   0
      End
      Begin ImDateCtrl.ImDate txt�މ�� 
         Height          =   375
         Left            =   11280
         TabIndex        =   49
         Top             =   840
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   65537
         HolidayColor    =   255
         AlignHorizontal =   0
         AlignVertical   =   0
         CursorPosition  =   0
         MaxDate         =   20991231
         MinDate         =   18680908
         Number          =   20100717
         Value           =   40376
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         ClipMode        =   0
         HighlightText   =   0
         DataProperty    =   0
         FirstMonth      =   4
         Format          =   "yyyy/mm/dd"
         DisplayFormat   =   "yyyy/mm/dd"
         NationalHolidays=   "1/1"
         UserHolidays    =   ""
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyNow          =   "{F3}"
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         PromptChar      =   "_"
         Text            =   "2010/07/17"
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�މ��"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5759
         MousePointer    =   0
      End
      Begin ImTextCtrl.ImText txt�y�V���[�� 
         Height          =   375
         Left            =   240
         TabIndex        =   43
         Top             =   4200
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   0   'False
         IMEMode         =   2
         InsertMode      =   -1  'True
         MaxLength       =   250
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�y�V���[��"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   88
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5775
         MousePointer    =   0
      End
      Begin VB.CheckBox chk���[�����M 
         Height          =   255
         Left            =   1680
         TabIndex        =   44
         Top             =   4920
         Width           =   495
      End
      Begin VB.CommandButton cmd�]�L 
         Caption         =   "�ڋq���]�L"
         Height          =   495
         Left            =   2880
         TabIndex        =   32
         Top             =   240
         Width           =   1815
      End
      Begin VB.Frame frm�j�� 
         BorderStyle     =   0  '�Ȃ�
         Height          =   495
         Left            =   4080
         TabIndex        =   120
         Top             =   1200
         Width           =   3375
         Begin VB.OptionButton opt���� 
            Caption         =   "����"
            Height          =   345
            Left            =   1560
            TabIndex        =   37
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton opt�j�� 
            Caption         =   "�j��"
            Height          =   375
            Left            =   480
            TabIndex        =   36
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin ImTextCtrl.ImText txt���l 
         Height          =   2895
         Left            =   8160
         TabIndex        =   50
         Top             =   1560
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5106
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   4
         InsertMode      =   -1  'True
         MaxLength       =   0
         MultiLine       =   -1  'True
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   2
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "&Caption"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   0
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5791
         MousePointer    =   0
      End
      Begin ImDateCtrl.ImDate txt����� 
         Height          =   375
         Left            =   8160
         TabIndex        =   48
         Top             =   840
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         _Version        =   65537
         HolidayColor    =   255
         AlignHorizontal =   0
         AlignVertical   =   0
         CursorPosition  =   0
         MaxDate         =   20991231
         MinDate         =   18680908
         Number          =   20081026
         Value           =   39747
         ErrorBeep       =   0   'False
         ReadOnly        =   0   'False
         ClipMode        =   0
         HighlightText   =   0
         DataProperty    =   0
         FirstMonth      =   4
         Format          =   "yyyy/mm/dd"
         DisplayFormat   =   "yyyy/mm/dd"
         NationalHolidays=   "1/1"
         UserHolidays    =   ""
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyNow          =   "{F3}"
         KeyPopup        =   "{SPACE}"
         KeyPrevious     =   ""
         PromptChar      =   "_"
         Text            =   "2008/10/26"
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�����"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   64
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":57AD
         MousePointer    =   0
      End
      Begin VB.ComboBox cmb�A�[�f���N���u 
         Height          =   345
         IMEMode         =   4  '�S�p�Ђ炪��
         ItemData        =   "ADF010.frx":57C9
         Left            =   9840
         List            =   "ADF010.frx":57CB
         TabIndex        =   47
         Top             =   360
         Width           =   2175
      End
      Begin ImTextCtrl.ImText txt���[�� 
         Height          =   375
         Left            =   720
         TabIndex        =   42
         Top             =   3720
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   2
         InsertMode      =   -1  'True
         MaxLength       =   50
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "���[��"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":57CD
         MousePointer    =   0
      End
      Begin ImTextCtrl.ImText txt�d�b�ԍ� 
         Height          =   375
         Left            =   480
         TabIndex        =   41
         Top             =   3240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   2
         InsertMode      =   -1  'True
         MaxLength       =   15
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�d�b�ԍ�"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   72
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":57E9
         MousePointer    =   0
      End
      Begin ImTextCtrl.ImText txt�Z��_���i 
         Height          =   375
         Left            =   360
         TabIndex        =   40
         Top             =   2760
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   4
         InsertMode      =   -1  'True
         MaxLength       =   100
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�Z��_���i"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   80
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5805
         MousePointer    =   0
      End
      Begin ImTextCtrl.ImText txt�Z��_��i 
         Height          =   375
         Left            =   360
         TabIndex        =   38
         Top             =   1800
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   4
         InsertMode      =   -1  'True
         MaxLength       =   100
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�Z��_��i"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   80
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5821
         MousePointer    =   0
      End
      Begin ImTextCtrl.ImText txt�X�֔ԍ� 
         Height          =   375
         Left            =   1200
         TabIndex        =   35
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   2
         InsertMode      =   -1  'True
         MaxLength       =   8
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "��"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   24
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":583D
         MousePointer    =   0
      End
      Begin ImTextCtrl.ImText txt�t���K�i 
         Height          =   375
         Left            =   4320
         TabIndex        =   34
         Top             =   840
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   5
         InsertMode      =   -1  'True
         MaxLength       =   40
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�t���K�i"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5859
         MousePointer    =   0
      End
      Begin ImTextCtrl.ImText txt�ڋq�� 
         Height          =   375
         Left            =   720
         TabIndex        =   33
         Top             =   840
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   4
         InsertMode      =   -1  'True
         MaxLength       =   40
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�ڋq��"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5875
         MousePointer    =   0
      End
      Begin ImTextCtrl.ImText txt�ڋqID 
         Height          =   375
         Left            =   720
         TabIndex        =   31
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   0
         InsertMode      =   -1  'True
         MaxLength       =   5
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   0   'False
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�ڋqID"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":5891
         MousePointer    =   0
      End
      Begin ImTextCtrl.ImText txt�Z��_���i 
         Height          =   375
         Left            =   360
         TabIndex        =   39
         Top             =   2280
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   661
         _Version        =   65537
         AlignHorizontal =   0
         AlignVertical   =   0
         AllowAll        =   3
         AllowHiragana   =   0
         AllowKatakana   =   0
         AllowLower      =   0
         AllowNumber     =   0
         AllowSpace      =   -1  'True
         AllowSymbol     =   0
         AllowUpper      =   0
         CursorPosition  =   -1
         ErrorBeep       =   0   'False
         FuriganaOn      =   0   'False
         HighlightText   =   -1  'True
         IMEMode         =   4
         InsertMode      =   -1  'True
         MaxLength       =   100
         MultiLine       =   0   'False
         ReadOnly        =   0   'False
         SelStart        =   0
         SelLength       =   0
         ScrollBars      =   0
         UsePopup        =   0   'False
         WindowHeight    =   24
         WindowWidth     =   50
         KeyClear        =   "{F2}"
         KeyNext         =   ""
         KeyPrevious     =   ""
         PasswordChar    =   ""
         Text            =   ""
         Appearance      =   1
         BackColor       =   -2147483643
         Enabled         =   -1  'True
         ForeColor       =   -2147483640
         BorderStyle     =   1
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         DropdownButton  =   0   'False
         SpinButton      =   0   'False
         Caption         =   "�Z��_���i"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   80
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SpinAutowrap    =   0   'False
         _StockProps     =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "ADF010.frx":58AD
         MousePointer    =   0
      End
      Begin VB.Label lb���[�����M 
         Caption         =   "���[�����M"
         Height          =   375
         Left            =   480
         TabIndex        =   56
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label lb�A�[�f���N���u 
         Caption         =   "�A�[�f���N���u"
         Height          =   255
         Left            =   8160
         TabIndex        =   57
         Top             =   405
         Width           =   1455
      End
   End
   Begin MSComctlLib.TabStrip tab��� 
      Height          =   6015
      Left            =   360
      TabIndex        =   29
      Top             =   6720
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   10610
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�ڋq"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�z����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���[������"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmd�L�����Z������ 
      Caption         =   "��ݾٌ���"
      Height          =   375
      Left            =   10920
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd�ۗ������� 
      Caption         =   "�ۗ�������"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmd�o�׍ς� 
      Caption         =   "�o�׍ς�"
      Height          =   375
      Left            =   9120
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd�o�׏����� 
      Caption         =   "�o�׏�����"
      Height          =   375
      Left            =   11640
      TabIndex        =   15
      Top             =   5760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd�����҂� 
      Caption         =   "�����҂�"
      Height          =   375
      Left            =   7920
      TabIndex        =   13
      Top             =   6360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd�V�K���� 
      Caption         =   "�V�K����"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdCSV�o�� 
      Caption         =   "e��`�o��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12480
      TabIndex        =   117
      Top             =   12840
      Width           =   1095
   End
   Begin VB.CommandButton cmd�A�[�f���N���u 
      Caption         =   "�N���u���[��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   26
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmd�o�ח\��ꗗ 
      Caption         =   "�o�ח\��ꗗ"
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd������ 
      Caption         =   "����������"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox cmb�������� 
      Height          =   345
      Left            =   240
      Style           =   2  '��ۯ���޳� ؽ�
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmd�S���� 
      Caption         =   "�S����"
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmd�S�I�� 
      Caption         =   "�S�I��"
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmd�N���u������ 
      Caption         =   "�N���u������"
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmd�N���u���� 
      Caption         =   "�N���u����"
      Height          =   375
      Left            =   9840
      TabIndex        =   19
      Top             =   6000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd�������� 
      Caption         =   "��������"
      Height          =   375
      Left            =   11640
      TabIndex        =   20
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd���o�׈ꗗ 
      Caption         =   "���o�׌���"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmd���[�� 
      Caption         =   "���[��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   111
      Top             =   12840
      Width           =   1095
   End
   Begin ImNumberCtrl.ImNumber txt�ݐϐ� 
      Height          =   375
      Left            =   9960
      TabIndex        =   115
      Top             =   13080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   65537
      AlignHorizontal =   1
      ClipMode        =   0
      ErrorBeep       =   0   'False
      ReadOnly        =   0   'False
      HighlightText   =   0   'False
      ZeroAllowed     =   -1  'True
      MinusColor      =   255
      MaxValue        =   99999
      MinValue        =   -99999
      Value           =   0
      SelStart        =   1
      SelLength       =   0
      KeyClear        =   "{F2}"
      KeyNext         =   ""
      KeyPopup        =   "{SPACE}"
      KeyPrevious     =   ""
      KeyThreeZero    =   ""
      SepDecimal      =   "."
      SepThousand     =   ","
      Text            =   "0"
      Format          =   "####0"
      DisplayFormat   =   ""
      Appearance      =   1
      BackColor       =   -2147483643
      Enabled         =   0   'False
      ForeColor       =   -2147483640
      BorderStyle     =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      DropdownButton  =   0   'False
      SpinButton      =   0   'False
      Caption         =   "&Caption"
      CaptionAlignment=   3
      CaptionColor    =   0
      CaptionWidth    =   0
      CaptionPosition =   0
      CaptionSpacing  =   3
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpinAutowrap    =   0   'False
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "ADF010.frx":58C9
      MousePointer    =   0
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "����"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmd��� 
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   110
      Top             =   12840
      Width           =   1095
   End
   Begin VB.CommandButton cmd�[�i�� 
      Caption         =   "�[�i��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   109
      Top             =   12840
      Width           =   1095
   End
   Begin VB.CommandButton cmd�폜2 
      Caption         =   "�����폜"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   108
      Top             =   12840
      Width           =   1095
   End
   Begin VB.CommandButton cmd�ǉ�2 
      Caption         =   "�V�K����"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   107
      Top             =   12840
      Width           =   1095
   End
   Begin VB.CommandButton cmd�폜1 
      Caption         =   "�ڋq�폜"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   24
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmd�ǉ�1 
      Caption         =   "�V�K�ڋq"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   23
      Top             =   6000
      Width           =   1095
   End
   Begin ImTextCtrl.ImText txt�������� 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   65537
      AlignHorizontal =   0
      AlignVertical   =   0
      AllowAll        =   3
      AllowHiragana   =   0
      AllowKatakana   =   0
      AllowLower      =   0
      AllowNumber     =   0
      AllowSpace      =   -1  'True
      AllowSymbol     =   0
      AllowUpper      =   0
      CursorPosition  =   -1
      ErrorBeep       =   0   'False
      FuriganaOn      =   0   'False
      HighlightText   =   0   'False
      IMEMode         =   4
      InsertMode      =   -1  'True
      MaxLength       =   255
      MultiLine       =   0   'False
      ReadOnly        =   0   'False
      SelStart        =   0
      SelLength       =   0
      ScrollBars      =   0
      UsePopup        =   0   'False
      WindowHeight    =   24
      WindowWidth     =   50
      KeyClear        =   "{F2}"
      KeyNext         =   ""
      KeyPrevious     =   ""
      PasswordChar    =   ""
      Text            =   ""
      Appearance      =   1
      BackColor       =   -2147483643
      Enabled         =   -1  'True
      ForeColor       =   -2147483640
      BorderStyle     =   1
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      DropdownButton  =   0   'False
      SpinButton      =   0   'False
      Caption         =   "&Caption"
      CaptionAlignment=   3
      CaptionColor    =   0
      CaptionWidth    =   0
      CaptionPosition =   0
      CaptionSpacing  =   3
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpinAutowrap    =   0   'False
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "ADF010.frx":58E5
      MousePointer    =   0
   End
   Begin VB.Label lbl���ӊ��N 
      Alignment       =   2  '��������
      BackColor       =   &H000000FF&
      Caption         =   "���ӊ��N"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   130
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lbl���2 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   14
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label txt���ӊ��N 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   6120
      TabIndex        =   27
      Top             =   6000
      Width           =   9375
   End
   Begin VB.Label lb�ݐϖ{�� 
      Caption         =   "�ݐϖ{��"
      Height          =   375
      Left            =   10200
      TabIndex        =   114
      Top             =   12840
      Width           =   975
   End
End
Attribute VB_Name = "ADF010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private G_�s�ԍ�        As Integer
Private G_�t���O        As Boolean
Private G_ROW           As Long
Private G_�^�uNO        As Integer
Public G_�ڋq���X�g_ROW As Long
Public G_�������X�g_ROW As Long
Public G_������         As String
Public G_���i��         As String

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SendMail Lib "bsmtp" _
      (szServer As String, szTo As String, szFrom As String, _
      szSubject As String, szBody As String, szFile As String) As String
'************************
'�I���W�i�����̓��[�h�萔
'************************
'�S�p�Ђ炪�ȓ���
Private Const MY_IME_CHMODE_ZEN_HIRA = IME_CMODE_ROMAN Or IME_CMODE_JAPANESE Or IME_CMODE_FULLSHAPE
'�S�p�J�^�J�i����
Private Const MY_IME_CHMODE_ZEN_KATA = IME_CMODE_ROMAN Or IME_CMODE_JAPANESE Or IME_CMODE_KATAKANA Or IME_CMODE_FULLSHAPE
'�S�p�p������
Private Const MY_IME_CHMODE_ZEN_EISU = IME_CMODE_ROMAN Or IME_CMODE_FULLSHAPE
'���p�J�^�J�i����
Private Const MY_IME_CHMODE_HAN_KATA = IME_CMODE_ROMAN Or IME_CMODE_JAPANESE Or IME_CMODE_KATAKANA Or IME_CMODE_LANGUAGE
'���p�p������
Private Const MY_IME_CHMODE_HAN_EISU = IME_CMODE_ROMAN

'************************************************************************
'�@  �\ :�t�H�[�����[�h
'************************************************************************
Private Sub Form_Load()
    
    Dim i As Integer
    Dim �X�܃}�X�^RS As New ADODB.Recordset
    
    Call �R�l�N�V����

    If va�ڋq���X�g.MaxRows >= 1 Then
        Call �����\��(1)
    End If
        
    ' �X�܃}�X�^�����[�h����
    G_����� = 0.08
    G_�d�� = "�d��8%"
    G_���� = "����8%"
    
    Call �X�܃}�X�^�擾(�X�܃}�X�^RS)
    If Not �X�܃}�X�^RS.EOF Then
        G_�X�ܖ� = �X�܃}�X�^RS!�X�ܖ�
        G_�X�ܗ��� = �X�܃}�X�^RS!�X�ܗ���
        G_�X�ܐF = �X�܃}�X�^RS!�X�ܐF
        G_���[�� = �X�܃}�X�^RS!���[��
        G_�T�[�o = �X�܃}�X�^RS!�T�[�o                                          ' mail.cathand.jp:587
        G_����� = CDbl(�X�܃}�X�^RS!�����)                                   ' = 1.05
        G_�d�� = �X�܃}�X�^RS!�d��                                              ' = "�d��5%"
        G_���� = �X�܃}�X�^RS!����                                              ' = "����5%"

        
        If �X�܃}�X�^RS!���M��2 <> "" Then
            G_���M�� = �X�܃}�X�^RS!���M��1 & vbTab & �X�܃}�X�^RS!���M��2          ' info@cathand.jp & vbTab & info@cathand.jp:info
        Else
            G_���M�� = �X�܃}�X�^RS!���M��1
        End If
        
        If G_�X�ܖ� = "�g���j�e�B�[�y�V�s��X" Then
            G_���M�� = G_���M�� & vbTab & "CRAM-MD5"
        End If
        G_���[�U = �X�܃}�X�^RS!���[�U                                          ' order2@cathand.jp
        G_�p�X���[�h = �X�܃}�X�^RS!�p�X���[�h                                  ' order2@cathand.jp
    End If
    
    �X�܃}�X�^RS.Close

    Call �y�V_�X�܃}�X�^�擾(�X�܃}�X�^RS)
    If Not �X�܃}�X�^RS.EOF Then
        G_�T�[�o2 = �X�܃}�X�^RS!�T�[�o                                         ' sub.fw.rakuten.ne.jp:587
        
        If �X�܃}�X�^RS!���M��2 <> "" Then
            G_���M��2 = �X�܃}�X�^RS!���M��1 & vbTab & �X�܃}�X�^RS!���M��2     ' 251377:IwK93MZNj0
        Else
            G_���M��2 = �X�܃}�X�^RS!���M��1
        End If
        
        G_���[��2 = �X�܃}�X�^RS!���[��
        G_���M��2 = G_���M��2 & vbTab & "CRAM-MD5"
    End If
    
    �X�܃}�X�^RS.Close


    Call cmb��������.Clear
    Call cmb��������.AddItem("�ڋq��")
    Call cmb��������.AddItem("���͂��於")
    Call cmb��������.AddItem("�t���K�i")
    Call cmb��������.AddItem("�d�b�ԍ�")
    Call cmb��������.AddItem("���[��")
    Call cmb��������.AddItem("�y�V���[��")
    Call cmb��������.AddItem("��")
    Call cmb��������.AddItem("�Z��1")
    Call cmb��������.AddItem("�Z��2")
    Call cmb��������.AddItem("�Z��3")
    Call cmb��������.AddItem("�����ԍ�")
    Call cmb��������.AddItem("�⍇�ԍ�")
    Call cmb��������.AddItem("�R�����C�tNO")
    Call cmb��������.AddItem("����ID")
    Call cmb��������.AddItem("�o�ד�")
    cmb��������.ListIndex = 0
    
    '�Z��2�̗���\���ɂ���
    va�ڋq���X�g.Col = COL_�Z��2
    va�ڋq���X�g.ColHidden = True
    
    '�Z��3�̗���\���ɂ���
    va�ڋq���X�g.Col = COL_�Z��3
    va�ڋq���X�g.ColHidden = True
    
    '�`�F�b�N�{�b�N�X�̗���\���ɂ���
    'va�ڋq���X�g.col = COL_�`�F�b�N
    'va�ڋq���X�g.ColHidden = True
    
    '����ł̗���\���ɂ���
    va�������X�g.Col = COL_�����
    va�������X�g.ColHidden = True
    
    '����ID�̗���\���ɂ���
    va�������X�g.Col = COL_����ID
    va�������X�g.ColHidden = True
    
    '�ڋqID�̗���\���ɂ���
    va�������X�g.Col = COL_�ڋqID2
    va�������X�g.ColHidden = True
    
    '�ڋq���̗���\���ɂ���
    va�������X�g.Col = COL_�ڋq��2
    va�������X�g.ColHidden = True
    
    ' �Q�ƌ����\���ɂ���
    va�������X�g.Col = COL_�Q�ƌ�
    va�������X�g.ColHidden = True
    
    ' �L�[���[�h���\���ɂ���
    va�������X�g.Col = COL_�L�[���[�h
    va�������X�g.ColHidden = True
    
    ' ���̓|�C���g���\���ɂ���
    va�������X�g.Col = COL_���̓|�C���g
    va�������X�g.ColHidden = True
    
    ' ���t�������\���ɂ���
    va�������X�g.Col = COL_���t����
    va�������X�g.ColHidden = True
    
    ' �ԕi�Ώۂ��\���ɂ���
    va�������X�g.Col = COL_�ԕi�Ώ�
    va�������X�g.ColHidden = True
    
    ' ���C�����e�B�[���\���ɂ���
    va�������X�g.Col = COL_���C�����e�B�[
    va�������X�g.ColHidden = True
    
    For i = 2 To va�ڋq���X�g.MaxCols
        va�ڋq���X�g.Col = i
        va�ڋq���X�g.row = -1
        va�ڋq���X�g.Protect = True
        va�ڋq���X�g.Lock = True
    Next i
    
    For i = 2 To va�������X�g.MaxCols
        va�������X�g.Col = i
        va�������X�g.row = -1
        va�������X�g.Protect = True
        va�������X�g.Lock = True
    Next i
    
    For i = 1 To va���[������.MaxCols
        va���[������.Col = i
        va���[������.row = -1
        va���[������.Protect = True
        va���[������.Lock = True
    Next i
           
    ' �A�[�f���N���u
    Call cmb�A�[�f���N���u.Clear
    Call cmb�A�[�f���N���u.AddItem("")
    Call cmb�A�[�f���N���u.AddItem("�A�[�f���N���u")
'    Call cmb�A�[�f���N���u.AddItem("�A�[�f���R����")
'    Call cmb�A�[�f���N���u.AddItem("�A�[�f���U����")
'    Call cmb�A�[�f���N���u.AddItem("�V�u�X�^�R����")
'    Call cmb�A�[�f���N���u.AddItem("�V�u�X�^�U����")
'    Call cmb�A�[�f���N���u.AddItem("�V�n�C�u���b�^�[�R����")
'    Call cmb�A�[�f���N���u.AddItem("�V�n�C�u���b�^�[�U����")
'    Call cmb�A�[�f���N���u.AddItem("------------------------------")
'    Call cmb�A�[�f���N���u.AddItem("�u�[�X�^�[�R����")
'    Call cmb�A�[�f���N���u.AddItem("�u�[�X�^�[�U����")
'    Call cmb�A�[�f���N���u.AddItem("�n�C�u���b�h�R����")
'    Call cmb�A�[�f���N���u.AddItem("�n�C�u���b�h�U����")
    Call cmb�A�[�f���N���u.AddItem("�Ȃ�")
    
    ' �X�e�[�^�X
    Call cmb�X�e�[�^�X.Clear
    Call cmb�X�e�[�^�X.AddItem("�V�K����")
    Call cmb�X�e�[�^�X.AddItem("������")
    Call cmb�X�e�[�^�X.AddItem("��������")
    Call cmb�X�e�[�^�X.AddItem("�N���W�b�g����")
    Call cmb�X�e�[�^�X.AddItem("�o�׏���")
    Call cmb�X�e�[�^�X.AddItem("�o�׊���")
    Call cmb�X�e�[�^�X.AddItem("�L�����Z��")
    Call cmb�X�e�[�^�X.AddItem("�R�����C�t")
    Call cmb�X�e�[�^�X.AddItem("�ۗ�")
    
    ' ���i��
    Call cmb���i��.Clear
    Call cmb���i��.AddItem("")
    Call cmb���i��.AddItem("�A�[�f��")
    Call cmb���i��.AddItem("�A�[�f��2�{�Z�b�g")
    Call cmb���i��.AddItem("�A�[�f���{�V�����v�[")
    
    Call cmb���i��.AddItem("------------------------------")
    Call cmb���i��.AddItem("�V�u�X�^")
    Call cmb���i��.AddItem("�V�n�C�u���b�^�[")
    Call cmb���i��.AddItem("�V�u�X�^�{�V�����v�[")
    Call cmb���i��.AddItem("�V�n�C�u���b�^�[�{�V�����v�[")
    Call cmb���i��.AddItem("------------------------------")
    Call cmb���i��.AddItem("�i�C�X���f�B�[")
    Call cmb���i��.AddItem("�i�C�X���f�B�[�{�V�����v�[")
    Call cmb���i��.AddItem("------------------------------")
    Call cmb���i��.AddItem("�u�[�X�^�[")
    Call cmb���i��.AddItem("�u�[�X�^�[�i�v���ь��ԁj")
    Call cmb���i��.AddItem("�n�C�u���b�h")
    Call cmb���i��.AddItem("�u�[�X�^�[�{�V�����v�[")
    Call cmb���i��.AddItem("�n�C�u���b�h�{�V�����v�[")
    
    Call cmb���i��.AddItem("------------------------------")
    Call cmb���i��.AddItem("�V�����v�[")
    Call cmb���i��.AddItem("�V�����v�[2�{�Z�b�g")
    Call cmb���i��.AddItem("�V�����v�[�i�v���[���g�j")
    Call cmb���i��.AddItem("�V�����v�[�{�g���[�g�����g")
    Call cmb���i��.AddItem("------------------------------")
    Call cmb���i��.AddItem("�g���[�g�����g")
    Call cmb���i��.AddItem("�g���[�g�����g�i�v���[���g�j")
    Call cmb���i��.AddItem("------------------------------")
    Call cmb���i��.AddItem("�u�X�^�T�O��OFF��")
    Call cmb���i��.AddItem("�n�C�u���b�^�[�T�O��OFF��")
    
    Call cmb���i��.AddItem("------------------------------")
'    Call cmb���i��.AddItem("�A�[�f�����p�E�}�j���A���i�v���[���g�j")
'    Call cmb���i��.AddItem("�����̐ςݏd�˂���؂ł��E�}�j���A���i�v���[���g�j")
'    Call cmb���i��.AddItem("�h�N�^�[�A�[�f���E��тc�u�c�i�v���[���g�j")
'    Call cmb���i��.AddItem("��тƉ^���E�}�j���A���i�v���[���g�j")
'    Call cmb���i��.AddItem("��сE���у}�j���A���i�v���[���g�j")
'    Call cmb���i��.AddItem("------------------------------")
    Call cmb���i��.AddItem("�A�[�f�����V�����v�[�����i")
    Call cmb���i��.AddItem("�A�[�f�������i")
    Call cmb���i��.AddItem("�V�����v�[�����i")
    
'    Call cmb���i��.AddItem("------------------------------")
'    Call cmb���i��.AddItem("���C�X�g���b�` �N�����W���O")
'    Call cmb���i��.AddItem("���C�X�g���b�` �E�H�b�V���O")
'    Call cmb���i��.AddItem("���C�X�g���b�` ���[�V����")
'    Call cmb���i��.AddItem("���C�X�g���b�` �W�F��")
'    Call cmb���i��.AddItem("���C�X�g���b�` ���C�����G�b�Z���X")
'    Call cmb���i��.AddItem("���C�X�g���b�` ��b���ϕi�Z�b�g")
    
    Call cmb���i��.AddItem("------------------------------")
    Call cmb���i��.AddItem("�A�[�f������")
    Call cmb���i��.AddItem("�~�j�܂�")
    
    ' ����
    Call cmb����.Clear
    Call cmb����.AddItem("�����")
    Call cmb����.AddItem("��ײ�")
    Call cmb����.AddItem("���̑�")
    
    ' �������@
    Call cmb�������@.Clear
    Call cmb�������@.AddItem("")
    Call cmb�������@.AddItem("�N���W�b�g")
    Call cmb�������@.AddItem("�����N���W�b�g")
    Call cmb�������@.AddItem("���i���")
    Call cmb�������@.AddItem("�R���r�j")
    Call cmb�������@.AddItem("��s�U��")
    Call cmb�������@.AddItem("�y�V�o���N����")
    Call cmb�������@.AddItem("�y�C�W�[")
    Call cmb�������@.AddItem("�㕥��")
    Call cmb�������@.AddItem("�|�C���g")
    Call cmb�������@.AddItem("�g�ь���")
    Call cmb�������@.AddItem("�d�q�}�l�[")
    Call cmb�������@.AddItem("���t�I�N")
    Call cmb�������@.AddItem("�|")
    
    ' ��s
    Call cmb��s.Clear
    Call cmb��s.AddItem("")
    Call cmb��s.AddItem("�݂���")
    Call cmb��s.AddItem("�y�V��s")
    Call cmb��s.AddItem("�X�֐U�֌���")
    
    ' ��z�Ǝ�
    Call cmb��z�Ǝ�.Clear
    Call cmb��z�Ǝ�.AddItem("����}��")
    Call cmb��z�Ǝ�.AddItem("�N���l�R���}�g")
    Call cmb��z�Ǝ�.AddItem("�䂤�p�b�N")
    Call cmb��z�Ǝ�.AddItem("���^�[�p�b�N")
    Call cmb��z�Ǝ�.AddItem("�y���J��")
    
    ' ������
    Call cmb������.Clear
    Call cmb������.AddItem("")
    Call cmb������.AddItem(G_�X�ܗ���)
    Call cmb������.AddItem("���ЃT�C�g")
'    Call cmb������.AddItem("�����g���b�N�X")
'    Call cmb������.AddItem("������̂��l�b�g")
    Call cmb������.AddItem("�A�}�]��")
'    Call cmb������.AddItem("�R�}�`")
    Call cmb������.AddItem("���t�I�N")
'    Call cmb������.AddItem("�C���t�H�g�b�v")
'    Call cmb������.AddItem("���[��")
'    Call cmb������.AddItem("FAX")
'    Call cmb������.AddItem("�d�b")
'    Call cmb������.AddItem("�������")
    Call cmb������.AddItem("���̑�")
    If G_�X�ܖ� <> "�g���j�e�B�[�y�V�s��X" Then
        Call cmb������.AddItem("�y�V")
    End If
    
    ' �e���v���[�g
    Call cmb�e���v���[�g.Clear
    Call cmb�e���v���[�g.AddItem("")
    Call cmb�e���v���[�g.AddItem("�A�[�f���V�K")
    Call cmb�e���v���[�g.AddItem("-------------------------------")
    Call cmb�e���v���[�g.AddItem("�V�u�X�^�V�K")
    Call cmb�e���v���[�g.AddItem("�V�n�C�u���b�^�[�V�K")
    Call cmb�e���v���[�g.AddItem("-------------------------------")
    Call cmb�e���v���[�g.AddItem("�u�[�X�^�[�V�K")
    Call cmb�e���v���[�g.AddItem("�n�C�u���b�h�V�K")
    Call cmb�e���v���[�g.AddItem("�i�C�X���f�B�[�V�K")
    Call cmb�e���v���[�g.AddItem("-------------------------------")
    Call cmb�e���v���[�g.AddItem("�V�����v�[�Q�{�Z�b�g�V�K")
    Call cmb�e���v���[�g.AddItem("�V�����v�[�V�K")
    Call cmb�e���v���[�g.AddItem("-------------------------------")
    Call cmb�e���v���[�g.AddItem("�V�����v�[���g���[�g�����g�V�K")
    Call cmb�e���v���[�g.AddItem("�g���[�g�����g�V�K")
    Call cmb�e���v���[�g.AddItem("-------------------------------")
    Call cmb�e���v���[�g.AddItem("�A�[�f�����V�����v�[�Z�b�g�V�K")
    Call cmb�e���v���[�g.AddItem("�V�u�X�^���V�����v�[�Z�b�g�V�K")
    Call cmb�e���v���[�g.AddItem("�V�n�C�u���b�^�[���V�����v�[�Z�b�g�V�K")
    Call cmb�e���v���[�g.AddItem("�u�[�X�^�[���V�����v�[�Z�b�g�V�K")
    Call cmb�e���v���[�g.AddItem("�n�C�u���b�g���V�����v�[�Z�b�g�V�K")
    Call cmb�e���v���[�g.AddItem("�i�C�X���f�B�[���V�����v�[�Z�b�g�V�K")
    Call cmb�e���v���[�g.AddItem("�����i�V�K")
    Call cmb�e���v���[�g.AddItem("-------------------------------")
    Call cmb�e���v���[�g.AddItem("�A�[�f�����p")
    Call cmb�e���v���[�g.AddItem("�����̐ςݏd�˂���؂ł�")
    Call cmb�e���v���[�g.AddItem("��тc�u�c")
    Call cmb�e���v���[�g.AddItem("��тƉ^��")
    Call cmb�e���v���[�g.AddItem("��сE����")

    G_�s�ԍ� = 0
    G_�^�uNO = 1
    G_������ = ""
    G_���i�� = ""
    G_�ڋq�}�X�^_�r���t���O = False
    
    cmb������.BackColor = vbRed

#If 0 Then

    If G_�X�ܖ� = "�g���j�e�B�[�y�V�s��X" Then
        txt�y�V���[��.Enabled = True
    Else
        txt�y�V���[��.Enabled = False
    End If
    
#End If
    
    If G_�X�ܖ� = "�g���j�e�B�[�y�V�s��X" Then
        cmd�ڍs.Visible = False
    End If
    
End Sub

'************************************************************************
'�@  �\ :�ڋq�}�X�^��\������
'************************************************************************
Private Sub Form_Activate()

    If G_�t���O = False Then
        
        Dim ADF016      As New ADF016
    
        Dim �ڋq�}�X�^RS As New ADODB.Recordset
        
        ADF010.Caption = G_�X�ܖ�
        ADF010.BackColor = Val(G_�X�ܐF)
        
        Call MsgBox(G_�X�ܖ� & "�p�̌ڋq�Ǘ��ł��B�ԈႦ�Ȃ��悤�ɒ��ӂ��ĉ������I", vbOKOnly, "�ڋq�Ǘ�")
          
        Call cmd���o�׈ꗗ_Click
        
        If ADF016.�ꊇ�N�[_�����m�F() > 0 Then
            If MsgBox("�I�[�g�V�b�v�Ŋm��f�[�^������܂�" & vbCr & vbLf & "�m�肵�܂����H", vbYesNo, "�ڋq�Ǘ�") = vbYes Then
                If MsgBox("�v�����^�̏�����OK�ł����H", vbYesNo, "�ڋq�Ǘ�") = vbYes Then
                    If ADF016.�ꊇ�N�[() > 0 Then
                        Call MsgBox("�A�[�f���N���u�̊m����s���܂����I", vbOKOnly, "�ڋq�Ǘ�")
                    End If
                End If
            End If
        End If
        
#If 0 Then
        ' �ڋq�}�X�^��S�����[�h����
        Call �ڋq�}�X�^�Ǎ�(�ڋq�}�X�^RS)
    
        ' �ڋq���X�g��\������
        Call �ڋq���X�g�\��(�ڋq�}�X�^RS)
        
        If va�������X�g.MaxRows > 0 Then
            
            ' �ŏI�s�̔w�i�F��ύX����
            Call va�������X�g_Click(1, va�������X�g.MaxRows)
            
            ' �Z���̃t�H�[�J�X���ŏI�s�ɐݒ肷��
            Call SpreadSetFocus(va�������X�g, va�������X�g.MaxRows, COL_�X�e�[�^�X)
            
        End If
#End If
        G_�t���O = True
        
    End If
    
    On Error Resume Next
    
    DoEvents
    
    Select Case G_�^�uNO
        Case 1
                txt�ڋq��.SetFocus
        Case 2
                txt�ڋq��.SetFocus
        Case 3
                txt�󒍓�.SetFocus
    End Select

    Call SpreadSetVal(va�����f����, 1, 1, G_�X�ܖ�)

End Sub

'************************************************************************
'�@  �\ :UNLOAD
'************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    
    ' �m�F���b�Z�[�W��\������
    If MsgBox("�I�����Ă�낵���ł����H", vbYesNo, "�ڋq�Ǘ�") <> vbYes Then
        Cancel = 1
        Exit Sub
    End If
    
    Cancel = 0
    End
    
End Sub

'************************************************************************
'�@  �\ :����{�^��
'************************************************************************
Private Sub cmd����_Click()

    ' �m�F���b�Z�[�W��\������
    If MsgBox("�I�����Ă�낵���ł����H", vbYesNo, "�ڋq�Ǘ�") <> vbYes Then Exit Sub

    End
    
End Sub

'************************************************************************
'�@  �\ :�ڋq���X�g��\������
'************************************************************************
Private Sub �ڋq���X�g�\��(ByRef �ڋq�}�X�^RS As ADODB.Recordset)

    Dim row As Integer
    Dim �Z�� As String
    
    row = 1
    G_�s�ԍ� = 0
    va�ڋq���X�g.ReDraw = False
    va�ڋq���X�g.MaxRows = 0
    
    ' ���������ڋq���X�g��\������
    With �ڋq�}�X�^RS
        Do Until .EOF
            va�ڋq���X�g.MaxRows = row
            
            Call SpreadSetVal(va�ڋq���X�g, row, COL_�`�F�b�N, 0)
          
            If Not IsNull(!�ڋqID) Then
                Call SpreadSetVal(va�ڋq���X�g, row, COL_�ڋqID, !�ڋqID)
            End If
            
            If Not IsNull(!�ڋq��) Then
                Call SpreadSetVal(va�ڋq���X�g, row, COL_�ڋq��, !�ڋq��)
            End If
            
            If Not IsNull(!�t���K�i) Then
                Call SpreadSetVal(va�ڋq���X�g, row, COL_�t���K�i, !�t���K�i)
            End If
            
            If Not IsNull(![��]) Then
                Call SpreadSetVal(va�ڋq���X�g, row, COL_��, ![��])
            End If
            
            �Z�� = ""
            If Not IsNull(!�Z��1) Then
'               Call SpreadSetVal(va�ڋq���X�g, row, COL_�Z��1, !�Z��1)
                �Z�� = �Z�� + !�Z��1
            End If
            
            If Not IsNull(!�Z��2) Then
'               Call SpreadSetVal(va�ڋq���X�g, row, COL_�Z��2, !�Z��2)
                �Z�� = �Z�� + !�Z��2
            End If
            
            If Not IsNull(!�Z��3) Then
'               Call SpreadSetVal(va�ڋq���X�g, row, COL_�Z��3, !�Z��3)
                �Z�� = �Z�� + !�Z��3
            End If
            
            Call SpreadSetVal(va�ڋq���X�g, row, COL_�Z��1, �Z��)
            
            If Not IsNull(!�d�b�ԍ�) Then
                Call SpreadSetVal(va�ڋq���X�g, row, COL_�d�b�ԍ�, !�d�b�ԍ�)
            End If
            
            If Not IsNull(!���[��) Then
                Call SpreadSetVal(va�ڋq���X�g, row, COL_���[��, !���[��)
            End If
            
            If Not IsNull(!�A�[�f���N���u) Then
                Call SpreadSetVal(va�ڋq���X�g, row, COL_�A�[�f���N���u, !�A�[�f���N���u)
            End If
            
            If Not IsNull(!�����) Then
                Call SpreadSetVal(va�ڋq���X�g, row, COL_�����, !�����)
            End If
            
            If Not IsNull(!���l) Then
                Call SpreadSetVal(va�ڋq���X�g, row, COL_���l, !���l)
                txt���ӊ��N.Caption = !���l
            End If
            
            If Not IsNull(!���͂��於) Then
                Call SpreadSetVal(va�ڋq���X�g, row, COL_���͂��於, !���͂��於)
            End If
            
            If Not IsNull(!���͂��惁�[��) Then
                Call SpreadSetVal(va�ڋq���X�g, row, COL_���͂��惁�[��, !���͂��惁�[��)
            End If
            
            If Not IsNull(!�y�V���[��) Then
                Call SpreadSetVal(va�ڋq���X�g, row, COL_�y�V���[��, !�y�V���[��)
            End If

            Call .MoveNext
            row = row + 1
        Loop
    End With
    
    �ڋq�}�X�^RS.Close
    
    va�ڋq���X�g.ReDraw = True
    
    If va�ڋq���X�g.MaxRows > 0 Then
        
        ' �Z���̃t�H�[�J�X���ŏI�s�ɐݒ肷��
        Call SpreadSetFocus(va�ڋq���X�g, va�ڋq���X�g.MaxRows, COL_�ڋq��)

        ' �ŏI�s�̔w�i�F��ύX����
        Call va�ڋq���X�g_Click(1, va�ڋq���X�g.MaxRows)
                
        ' �擪�s�̒����f�[�^��\������
        If va�ڋq���X�g.MaxRows >= 1 Then
            Call �����\��(va�ڋq���X�g.MaxRows)
        End If
    Else
        va�������X�g.MaxRows = 0
    End If
End Sub

'************************************************************************
'�@  �\ :�ڋq���X�g�ɂP�s�ǉ�����B
'************************************************************************
Private Sub cmd�ǉ�1_Click()
    
    Call �g�����U�N�V�����f�[�^�̍X�V
    
    G_�^�uNO = 1
    tab���.Tabs(G_�^�uNO).Selected = True
    
    ' �ڋq���X�g�ɂP�s�ǉ�����
    va�ڋq���X�g.MaxRows = va�ڋq���X�g.MaxRows + 1
    
    ' �Z���̃t�H�[�J�X��ǉ������s�ɐݒ肷��
    Call SpreadSetFocus(va�ڋq���X�g, va�ڋq���X�g.MaxRows, COL_�ڋq��)
        
    ' �ǉ������s�̔w�i�F��ύX����
    Call va�ڋq���X�g_Click(1, va�ڋq���X�g.MaxRows)
    
    ' �������X�g����������
    va�������X�g.MaxRows = 0
    
    Call �ڋq���N���A
    Call �������N���A
    
End Sub

'************************************************************************
'�@  �\ :�ڋq���X�g�őI������Ă���s���폜����B
'************************************************************************
Private Sub cmd�폜1_Click()
    
    Dim i As Integer
    Dim row As Integer
    Dim ���� As Integer
    Dim �ڋqID As String
    
    Call �g�����U�N�V�����f�[�^�̍X�V

    ' �m�F���b�Z�[�W��\������
    If MsgBox("�폜���Ă�낵���ł����H", vbYesNo, "�ڋq�Ǘ�") <> vbYes Then Exit Sub
    
    ���� = 0
    For i = 1 To va�ڋq���X�g.MaxRows
        If SpreadGetVal(va�ڋq���X�g, i, COL_�`�F�b�N) = "1" Then
            �ڋqID = SpreadGetVal(va�ڋq���X�g, i, COL_�ڋqID)
    
            Call �ڋq�}�X�^�폜(�ڋqID)
        
        End If
    Next i
    
    ' �ڋq���X�g��\��������
    If txt�ڋq��.Text <> "" Then
        Call cmd����_Click
    Else
        G_�t���O = False
        Call Form_Activate
    End If
    
    ' �擪�s�̒����f�[�^��\������
    If va�ڋq���X�g.MaxRows >= 1 Then
        Call �����\��(va�ڋq���X�g.MaxRows)
    End If
    
    Call MsgBox("�ڋq�f�[�^���폜���܂���", vbOKOnly, "�ڋq�Ǘ�")
    
End Sub

'************************************************************************
'�@  �\ :�`�[����
'************************************************************************
Private Sub cmdTOOL_Click()

    Dim ADF022      As New ADF022
            
    Call ADF022.Show(1)
    
End Sub

'************************************************************************
'�@  �\ :�ڋq���X�g�̍s���ς�����璍����\��������
'************************************************************************
Private Sub va�ڋq���X�g_Click(ByVal Col As Long, ByVal row As Long)
    
    Dim �d�b�ԍ�        As String
    Dim �ڋq�}�X�^RS    As New ADODB.Recordset
    
    If row < 1 Then Exit Sub
    
    va�ڋq���X�g.ReDraw = False
    
    With va�ڋq���X�g
        .ReDraw = False
        .Col = -1
        .row = -1
        .BackColorStyle = 1
        .BackColor = vbWhite
        
        .row = row
        .Col = -1
        .BackColorStyle = 1
        .BackColor = vbCyan
        .ReDraw = True
    End With

    va�ڋq���X�g.ReDraw = True
    
    G_�ڋq���X�g_ROW = row
    
    Call Tab���_Click
    
    txt���ӊ��N.Caption = SpreadGetVal(va�ڋq���X�g, row, COL_���l)
    
    �d�b�ԍ� = SpreadGetVal(va�ڋq���X�g, row, COL_�d�b�ԍ�)
    Call �y�V_�d�b�ԍ�����(�ڋq�}�X�^RS, �d�b�ԍ�)
    
    lbl���ӊ��N.Visible = False
    
    If G_�X�ܖ� = "�g���j�e�B�[�y�V�s��X" Then
        If �ڋq�}�X�^RS!���� > 0 Then
            lbl���ӊ��N.Visible = True
            lbl���ӊ��N.Caption = "Yahoo�ڋq"
        Else
            lbl���ӊ��N.Visible = False
        End If
    Else
        If �ڋq�}�X�^RS!���� > 0 Then
            lbl���ӊ��N.Visible = True
            lbl���ӊ��N.Caption = "�y�V�ڋq"
        Else
            lbl���ӊ��N.Visible = False
        End If
    End If
    
    �ڋq�}�X�^RS.Close
    
    Call �����\��(row)
    
End Sub

'************************************************************************
'�@  �\ :������\������
'************************************************************************
Private Sub �����\��(ByVal �s As Integer)

    Dim row             As Integer
    Dim �ݐϖ{��        As Integer
    Dim ���㖾��RS      As New ADODB.Recordset
    Dim �ڋqID          As String
    Dim �z�B��]����    As String
    
    �ڋqID = SpreadGetVal(va�ڋq���X�g, �s, COL_�ڋqID)
    
    ' �I�����ꂽ�ڋq�̒����f�[�^���擾����
    Call ��������(�ڋqID, ���㖾��RS)

    �ݐϖ{�� = 0
    
    row = 1
    
    va�������X�g.ReDraw = False
    va�������X�g.MaxRows = 0
    
    ' �I�����ꂽ�ڋq�̒����f�[�^��\������
    With ���㖾��RS
        Do Until .EOF
            va�������X�g.MaxRows = row
            
            Call SpreadSetVal(va�������X�g, row, 1, 0)
          
            If Not IsNull(!�󒍓�) Then
                Call SpreadSetVal(va�������X�g, row, COL_�󒍓�, !�󒍓�)
            End If
          
            If Not IsNull(!�X�e�[�^�X) Then
                Call SpreadSetVal(va�������X�g, row, COL_�X�e�[�^�X, !�X�e�[�^�X)
            End If
          
            If Not IsNull(!���i��) Then
                Call SpreadSetVal(va�������X�g, row, COL_���i��, !���i��)
            End If
          
            If Not IsNull(!�������@) Then
                Call SpreadSetVal(va�������X�g, row, COL_�������@, !�������@)
            End If
            
            �z�B��]���� = ""
            
            If Not IsNull(!�z�B��]����) Then
                �z�B��]���� = !�z�B��]����
            End If
            
            If Not IsNull(!�z�B��]����2) Then
                �z�B��]���� = �z�B��]���� + " " + !�z�B��]����2
            End If
          
            Call SpreadSetVal(va�������X�g, row, COL_�z�B��]����, �z�B��]����)

            If Not IsNull(!�P��) Then
                Call SpreadSetVal(va�������X�g, row, COL_�P��, !�P��)
            End If
          
            If Not IsNull(!����) Then
                Call SpreadSetVal(va�������X�g, row, COL_����, !����)
            End If
          
            If Not IsNull(!����) Then
                Call SpreadSetVal(va�������X�g, row, COL_����, !����)
            End If
          
            If Not IsNull(!���z) Then
                Call SpreadSetVal(va�������X�g, row, COL_���z, !���z)
            End If
          
            If Not IsNull(!�����) Then
                Call SpreadSetVal(va�������X�g, row, COL_�����, !�����)
            End If
          
            If Not IsNull(!����) Then
                Call SpreadSetVal(va�������X�g, row, COL_����, !����)
            End If
          
            If Not IsNull(!�ԋ�) Then
                Call SpreadSetVal(va�������X�g, row, COL_�ԋ�, !�ԋ�)
            End If
          
            If Not IsNull(!���̑��萔��) Then
                Call SpreadSetVal(va�������X�g, row, COL_���̑��萔��, !���̑��萔��)
            End If
          
            If Not IsNull(!���v���z) Then
                Call SpreadSetVal(va�������X�g, row, COL_���v���z, !���v���z)
            End If
          
            If Not IsNull(!������) Then
                Call SpreadSetVal(va�������X�g, row, COL_������, !������)
            End If
          
            If Not IsNull(!�o�ד�) Then
                Call SpreadSetVal(va�������X�g, row, COL_�o�ד�, !�o�ד�)
            End If
          
            If Not IsNull(!���ד�) Then
                Call SpreadSetVal(va�������X�g, row, COL_���ד�, !���ד�)
            End If
          
            If Not IsNull(!��z�Ǝ�) Then
                Call SpreadSetVal(va�������X�g, row, COL_��z�Ǝ�, !��z�Ǝ�)
            End If
            
            If Not IsNull(!������) Then
                Call SpreadSetVal(va�������X�g, row, COL_������, !������)
            End If
            
            If Not IsNull(!Yahoo�����ԍ�) Then
                Call SpreadSetVal(va�������X�g, row, COL_Yahoo�����ԍ�, !Yahoo�����ԍ�)
            End If
            
            If Not IsNull(!�Q�ƌ�) Then
                Call SpreadSetVal(va�������X�g, row, COL_�Q�ƌ�, !�Q�ƌ�)
            End If
            
            If Not IsNull(!�L�[���[�h) Then
                Call SpreadSetVal(va�������X�g, row, COL_�L�[���[�h, !�L�[���[�h)
            End If
            
            If Not IsNull(!���̓|�C���g) Then
                Call SpreadSetVal(va�������X�g, row, COL_���̓|�C���g, !���̓|�C���g)
            End If
            
            If Not IsNull(!���i�R�[�h) Then
                Call SpreadSetVal(va�������X�g, row, COL_���i�R�[�h, !���i�R�[�h)
            End If
            
            If Not IsNull(!���C�����e�B�[) Then
                Call SpreadSetVal(va�������X�g, row, COL_���C�����e�B�[, !���C�����e�B�[)
            End If
            
            If Not IsNull(!���t����) Then
                Call SpreadSetVal(va�������X�g, row, COL_���t����, !���t����)
            End If
            
            If Not IsNull(!�ԕi�Ώ�) Then
                Call SpreadSetVal(va�������X�g, row, COL_�ԕi�Ώ�, !�ԕi�Ώ�)
            End If
            
            If Not IsNull(!�x���ԍ�) Then
                Call SpreadSetVal(va�������X�g, row, COL_�x���ԍ�, !�x���ԍ�)
            End If
            
            If Not IsNull(!�⍇�ԍ�) Then
                Call SpreadSetVal(va�������X�g, row, COL_�⍇�ԍ�, !�⍇�ԍ�)
            End If
            
            If Not IsNull(!���l1) Then
                Call SpreadSetVal(va�������X�g, row, COL_���l1, !���l1)
            End If
            
            If Not IsNull(!���l2) Then
                Call SpreadSetVal(va�������X�g, row, COL_���l2, !���l2)
            End If
            
            If Not IsNull(!���l3) Then
                Call SpreadSetVal(va�������X�g, row, COL_���l3, !���l3)
            End If
            
            If Not IsNull(!����ID) Then
                Call SpreadSetVal(va�������X�g, row, COL_����ID, !����ID)
            End If
            
            If Not IsNull(!�ڋqID) Then
                Call SpreadSetVal(va�������X�g, row, COL_�ڋqID2, !�ڋqID)
            End If
            
            If Not IsNull(!�ڋq��) Then
                Call SpreadSetVal(va�������X�g, row, COL_�ڋq��2, !�ڋq��)
            End If
            
            If Not IsNull(!���[�����M) Then
                Call SpreadSetVal(va�������X�g, row, COL_���[�����M, !���[�����M)
            End If

            If !�����敪 = "%" Then
                Call SpreadSetVal(va�������X�g, row, COL_�����敪, "��")
            Else
                Call SpreadSetVal(va�������X�g, row, COL_�����敪, "�~")
            End If
            
            If Not IsNull(!�o�ח\���) Then
                Call SpreadSetVal(va�������X�g, row, COL_�o�ח\���, !�o�ח\���)
            End If
            
            If Not IsNull(!����URL) Then
                Call SpreadSetVal(va�������X�g, row, COL_����URL, !����URL)
            End If
            
            G_����_�X�V���� = IIf(IsNull(!�X�V����), Now, !�X�V����)
            
            Call .MoveNext
            row = row + 1
        Loop
    End With
    
    ���㖾��RS.Close
    
    If va�������X�g.MaxRows > 0 Then
        
        ' �ŏI�s�̔w�i�F��ύX����
        Call va�������X�g_Click(1, va�������X�g.MaxRows)
        
        ' �Z���̃t�H�[�J�X���ŏI�s�ɐݒ肷��
        'Call SpreadSetFocus(va�������X�g, va�������X�g.MaxRows, COL_�X�e�[�^�X)
    Else
        'Call �������N���A
    End If
    
    va�������X�g.ReDraw = True
    
    ' �ݐϖ{����\������
    txt�ݐϐ�.Text = �ݐϐ��v�Z()

End Sub

'************************************************************************
'�@  �\ :�������X�g�ɂP�s�ǉ�����
'************************************************************************
Private Sub cmd�ǉ�2_Click()
    
    Call �g�����U�N�V�����f�[�^�̍X�V

    G_�^�uNO = 3
    tab���.Tabs(G_�^�uNO).Selected = True
    
    If txt�ڋqID.Text = "" Then
        Call MsgBox("�悸�ڋq�f�[�^��o�^���ĉ�����", vbOKOnly, "�ڋq�Ǘ�")
        G_�^�uNO = 1
        tab���.Tabs(G_�^�uNO).Selected = True
        Exit Sub
    End If
    
    va�������X�g.MaxRows = va�������X�g.MaxRows + 1
    
    txt�󒍓�.Text = Format(Now, "YYYY/MM/DD")
    cmb�X�e�[�^�X.Text = "�V�K����"
    cmb���i��.Text = ""
    cmb����.Text = "�����"
    cmb�������@.Text = "�N���W�b�g"
    cmb��s.Text = ""
    txt�P��.Value = 0
    txt����.Value = 0
    txt����.Value = 1
    txt����.Value = 0
    txt�ԋ�.Value = 0
    txt���̑��萔��.Value = 0
    txt���v���z.Text = 0
    txt����ID.Text = "-1"
    txt�R�����C�t.Text = ""
            
    txt�z�B����.Text = ""
    txt�o�ד�.Text = "____/__/__"
    txt�x���ԍ�.Text = ""
    txt�⍇�ԍ�.Text = ""
    txt������.Text = "____/__/__"
    txt���[�����M.Text = ""
    cmb������.Text = ""
    txt�����ԍ�.Text = ""
    txt���l2.Text = ""
    txt�o�ח\���.Text = "____/__/__"
    txt����URL.Text = ""
    
    G_����_�X�V���� = Now
    G_������ = ""
    
    cmb��z�Ǝ�.Text = "����}��"
    G_���i�� = ""


    ' �Z���̃t�H�[�J�X��ǉ������s�ɐݒ肷��
    Call SpreadSetFocus(va�������X�g, va�������X�g.MaxRows, COL_�X�e�[�^�X)

    ' �w�i�F��ύX����
    Call va�������X�g_Click(1, va�������X�g.MaxRows)
        
    ' ���㖾�ׂ��o�͂���
    Call ����_�X�V
    
End Sub

'************************************************************************
'�@  �\ :�������X�g�őI������Ă���s���폜����B
'************************************************************************
Private Sub cmd�폜2_Click()
    
    Dim i As Integer
    Dim ���� As Integer
    Dim row As Integer
    Dim ����ID As String
    
    Call �g�����U�N�V�����f�[�^�̍X�V

    ' �m�F���b�Z�[�W��\������
    If MsgBox("�폜���Ă�낵���ł����H", vbYesNo, "�ڋq�Ǘ�") <> vbYes Then Exit Sub
    
    ���� = 0
    For i = 1 To va�������X�g.MaxRows
        If SpreadGetVal(va�������X�g, i, COL_�`�F�b�N) = "1" Then
            ����ID = SpreadGetVal(va�������X�g, i, COL_����ID)
    
            If ����ID <> "" Then
                Call ���㖾�׍폜(����ID)
                ���� = ���� + 1
            End If
        End If
    Next i
    
    ' �����f�[�^��\��������
    Call �����\��(G_�ڋq���X�g_ROW)
    
    If ���� > 0 Then
        Call MsgBox("�����f�[�^���폜���܂���", vbOKOnly, "�ڋq�Ǘ�")
    End If
    
End Sub

'************************************************************************
'�@  �\ :�������X�g�̍s���ς�����璍����\��������
'************************************************************************
Private Sub va�������X�g_Click(ByVal Col As Long, ByVal row As Long)
    
    If row < 1 Then Exit Sub
    
    va�������X�g.ReDraw = False
    
    With va�������X�g
        .ReDraw = False
        .Col = -1
        .row = -1
        .BackColorStyle = 1
        .BackColor = vbWhite
        
        .row = row
        .Col = -1
        .BackColorStyle = 1
        .BackColor = vbCyan
        .ReDraw = True
    End With
    
    G_�������X�g_ROW = row
        
    va�������X�g.ReDraw = True
    
    If SpreadGetVal(va�������X�g, G_�������X�g_ROW, COL_�������@) = "�R���r�j" Then
        txt����URL.Caption = "����URL"
    Else
        txt����URL.Caption = "����ID"
    End If
    
    Call Tab���_Click
    
End Sub

'************************************************************************
'�@  �\ :���[�������̖{����\������B
'************************************************************************
Private Sub va���[������_Click(ByVal Col As Long, ByVal row As Long)
    
    Dim ���[������RS As New ADODB.Recordset
    Dim ����ID      As String
    Dim ���M����    As String
    
    If row < 1 Then Exit Sub
    
    va���[������.ReDraw = False
    
    With va���[������
        .ReDraw = False
        .Col = -1
        .row = -1
        .BackColorStyle = 1
        .BackColor = vbWhite
        
        .row = row
        .Col = -1
        .BackColorStyle = 1
        .BackColor = vbCyan
        .ReDraw = True
    End With
    
    va���[������.ReDraw = True
    
    ���M���� = SpreadGetVal(va���[������, row, 1)
    
    ����ID = SpreadGetVal(va�������X�g, G_�������X�g_ROW, COL_����ID)
    
    If ����ID = "" Or ����ID = "����ID" Then Exit Sub
    
    ' �I�����ꂽ���[���������擾����
    Call ���[����������2(����ID, ���M����, ���[������RS)
    
    If Not ���[������RS.EOF Then
        txt���[���{��.Text = ���[������RS!���[���{��
    End If
    
    ���[������RS.Close
    
End Sub

'************************************************************************
'�@  �\�@�����ڍׂ�\������B
'************************************************************************
Private Sub Tab���_Click()
    
    G_�^�uNO = Me.tab���.SelectedItem.Index
    
   
    Select Case G_�^�uNO
        ' �ڋq���^�u
        Case 1
            frm�ڋq.Visible = True
            frm����.Visible = False
            frm���[������.Visible = False
            
            txt�ڋqID.Visible = True
            txt�ڋqID.Enabled = False
            txt�y�V���[��.Visible = True
            lb�A�[�f���N���u.Visible = True
            cmb�A�[�f���N���u.Visible = True
            txt�����.Visible = True
            txt�މ��.Visible = True
            txt�a����.Visible = True
            cmd�]�L.Visible = False
            chk���[�����M.Visible = True
            lb���[�����M.Visible = True
            cmd�]��.Visible = True
            chk����1.Visible = True
            chk����2.Visible = True
            chk����3.Visible = True
            chk����4.Visible = True
            chk����5.Visible = True
            Call �ڋq�^�u_�\��
        ' �z�B��^�u
        Case 2
            If txt�ڋqID.Text = "" Then
                Call MsgBox("�ڋq��񂪖����͂ł�", vbOKOnly, "�ڋq�Ǘ�")
                G_�^�uNO = 1
                tab���.Tabs(G_�^�uNO).Selected = True
                Exit Sub
            End If
            
            frm�ڋq.Visible = True
            frm����.Visible = False
            frm���[������.Visible = False
            
            txt�ڋqID.Visible = False
            txt�y�V���[��.Visible = False
            lb�A�[�f���N���u.Visible = False
            cmb�A�[�f���N���u.Visible = False
            txt�����.Visible = False
            txt�މ��.Visible = False
            txt�a����.Visible = False
            cmd�]�L.Visible = True
            chk���[�����M.Visible = False
            lb���[�����M.Visible = False
            cmd�]��.Visible = False
            chk����1.Visible = False
            chk����2.Visible = False
            chk����3.Visible = False
            chk����4.Visible = False
            chk����5.Visible = False
            Call �ڋq�^�u_�\��
            
        ' �����^�u
        Case 3
            If txt�ڋqID.Text = "" Then
                Call MsgBox("�ڋq��񂪖����͂ł�", vbOKOnly, "�ڋq�Ǘ�")
                G_�^�uNO = 1
                tab���.Tabs(G_�^�uNO).Selected = True
                Exit Sub
            End If
                    
            frm�ڋq.Visible = False
            frm����.Visible = True
            frm���[������.Visible = False
            
            Call �����^�u_�\��
            
        ' ���[�������^�u
        Case 4
            If txt�ڋqID.Text = "" Then
                Call MsgBox("�ڋq��񂪖����͂ł�", vbOKOnly, "�ڋq�Ǘ�")
                G_�^�uNO = 1
                tab���.Tabs(G_�^�uNO).Selected = True
                Exit Sub
            End If
            
            frm�ڋq.Visible = False
            frm����.Visible = False
            frm���[������.Visible = True
            
            Call ���[������_�\��
            
    End Select
    
    On Error Resume Next
    
    'DoEvents
    
    Select Case G_�^�uNO
        Case 1
                txt�ڋq��.SetFocus
        Case 2
                txt�ڋq��.SetFocus
        Case 3
                txt�󒍓�.SetFocus
        Case 4
                va���[������.SetFocus
    End Select
    
End Sub

'************************************************************************
'�@  �\ :�ڋq�^�u����\������
'************************************************************************
Private Sub �ڋq�^�u_�\��()

    Dim �ڋqID As String
    Dim �ڋq�}�X�^RS As New ADODB.Recordset
  
    Dim ���㖾��RS As New ADODB.Recordset
    Dim ����ID As String
  
    �ڋqID = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�ڋqID)
    
    If �ڋqID = "ID" Then Exit Sub
    
    ' �I�����ꂽ�ڋq�̒����f�[�^���擾����
    Select Case G_�^�uNO
        Case 1
            Call �ڋq�}�X�^1���Ǎ�(�ڋq�}�X�^RS, �ڋqID)
        Case 2
            Call �z����1���Ǎ�(�ڋq�}�X�^RS, �ڋqID)
        Case 3
            Exit Sub
    End Select
    
    With �ڋq�}�X�^RS
        If �ڋq�}�X�^RS.EOF Then
            txt�ڋqID.Text = �ڋqID
            txt�ڋq��.Text = ""
            txt�t���K�i.Text = ""
            txt�X�֔ԍ�.Text = ""
            opt�j��.Value = True
            opt�j��.Value = True
            txt�Z��_��i.Text = ""
            txt�Z��_���i.Text = ""
            txt�Z��_���i.Text = ""
            txt�d�b�ԍ�.Text = ""
            txt���[��.Text = ""
            txt�y�V���[��.Text = ""
            cmb�A�[�f���N���u.Text = ""
            txt�����.Text = "____/__/__"
            txt�މ��.Text = "____/__/__"
            txt���l.Text = ""
            chk���[�����M.Value = 1
            txt�a����.Text = "____/__/__"
            
            G_�ڋq�}�X�^_�r���t���O = True
            chk����1.Value = 0
            chk����2.Value = 0
            chk����3.Value = 0
            chk����4.Value = 0
            chk����5.Value = 0
            G_�ڋq�}�X�^_�r���t���O = False
            
            If G_�^�uNO = 1 Then
                G_�ڋq_�X�V���� = Now
            Else
                G_�z��_�X�V���� = Now
            End If
        Else
            txt�ڋqID.Text = !�ڋqID
            txt�ڋq��.Text = !�ڋq��
            txt�t���K�i.Text = !�t���K�i
            txt�X�֔ԍ�.Text = ![��]
            If !���� = "1" Then opt�j��.Value = True Else opt�j��.Value = False
            If !���� = "2" Then opt����.Value = True Else opt����.Value = False
            txt�Z��_��i.Text = !�Z��1
            txt�Z��_���i.Text = !�Z��2
            txt�Z��_���i.Text = IIf(IsNull(!�Z��3), "", !�Z��3)
            txt�d�b�ԍ�.Text = !�d�b�ԍ�
            txt���[��.Text = !���[��
            
            If G_�^�uNO = 1 Then
                txt�y�V���[��.Text = IIf(IsNull(!�y�V���[��), "", !�y�V���[��)
                chk���[�����M.Value = !���[�����M
                
                If IsNull(!�a����) Or !�a���� = "" Then
                    txt�a����.Text = "____/__/__"
                Else
                    txt�a����.Text = !�a����
                End If
            End If
            
            If G_�^�uNO = 1 Then
                cmb�A�[�f���N���u.Text = !�A�[�f���N���u
                If IsNull(!�����) Or !����� = "" Then
                    txt�����.Text = "____/__/__"
                Else
                    txt�����.Text = !�����
                End If
                
                If IsNull(!�މ��) Or !�މ�� = "" Then
                    txt�މ��.Text = "____/__/__"
                Else
                    txt�މ��.Text = !�މ��
                End If
            End If
            
            txt���l.Text = !���l
            
            If G_�^�uNO = 1 Then
                G_�ڋq�}�X�^_�r���t���O = True
                chk����1.Value = IIf(IsNull(!����1), 0, !����1)
                chk����2.Value = IIf(IsNull(!����2), 0, !����2)
                chk����3.Value = IIf(IsNull(!����3), 0, !����3)
                chk����4.Value = IIf(IsNull(!����4), 0, !����4)
                chk����5.Value = IIf(IsNull(!����5), 0, !����5)
                G_�ڋq�}�X�^_�r���t���O = False
            End If
            
            If G_�^�uNO = 1 Then
                G_�ڋq_�X�V���� = IIf(IsNull(!�X�V����), Now, !�X�V����)
            Else
                G_�z��_�X�V���� = IIf(IsNull(!�X�V����), Now, !�X�V����)
            End If
                        
        End If
    End With
    
    �ڋq�}�X�^RS.Close
    
#If 0 Then
    ����ID = SpreadGetVal(va�������X�g, G_�������X�g_ROW, COL_����ID)
    
    ' �I�����ꂽ�ڋq�̒����f�[�^���擾����
    Call ��������2(����ID, ���㖾��RS)
    G_����_�X�V���� = IIf(IsNull(���㖾��RS!�X�V����), Now, ���㖾��RS!�X�V����)
    ���㖾��RS.Close
#End If

End Sub

'************************************************************************
'�@  �\�@�����ڍׂ�\������B
'************************************************************************
Private Sub �����^�u_�\��()
    
    Dim ���㖾��RS As New ADODB.Recordset
    Dim ����ID As String
    
    Dim �ڋqID As String
    Dim �ڋq�}�X�^RS As New ADODB.Recordset
      
    ����ID = SpreadGetVal(va�������X�g, G_�������X�g_ROW, COL_����ID)
    
    ' �I�����ꂽ�ڋq�̒����f�[�^���擾����
    Call ��������2(����ID, ���㖾��RS)
        
    With ���㖾��RS
        If ���㖾��RS.EOF Then
            txt����ID.Text = ""
            txt�󒍓�.Text = Format(Now, "YYYY/MM/DD")
            cmb�X�e�[�^�X.Text = "�V�K����"
            cmb���i��.Text = ""
            cmb����.Text = "�����"
            cmb�������@.Text = "�N���W�b�g"
            cmb��s.Text = ""
            txt�z�B����.Text = ""
            txt�z�B����2.Text = ""
            txt�o�ד�.Text = "____/__/__"
            cmb��z�Ǝ�.Text = "����}��"
            txt�x���ԍ�.Text = ""
            txt�⍇�ԍ�.Text = ""
            txt�d�����z.Value = 0
            txt������.Text = "____/__/__"
            txt�P��.Value = 0
            txt����.Value = 0
            'cmd����.Caption = "%"
            txt����.Value = 1
            txt����.Value = 0
            txt�ב��^��.Value = 0
            txt�ԋ�.Value = 0
            txt���̑��萔��.Value = 0
            txt���v���z.Text = 0
            txt���[�����M.Text = ""
            cmb������.Text = ""
            txt�����ԍ�.Text = ""
            txt���l2.Text = ""
            txt�R�����C�t = ""
            txt�o�ח\���.Text = "____/__/__"
            txt����URL.Text = ""

            G_����_�X�V���� = Now
            G_���i�� = ""
            G_������ = ""

        Else
            txt����ID.Text = IIf(Not IsNull(!����ID), !����ID, "")
            txt�󒍓�.Text = IIf(Not IsNull(!�󒍓�), IIf(!�󒍓� <> "", !�󒍓�, "____/__/__"), "____/__/__")
            cmb�X�e�[�^�X.Text = IIf(Not IsNull(!�X�e�[�^�X), !�X�e�[�^�X, "")
            cmb���i��.Text = IIf(Not IsNull(!���i��), !���i��, "")
            
            If IsNull(!����) = True Then
                If �A�[�f������(cmb���i��.Text) = 1 Or �A�[�f������(cmb���i��.Text) = 9 Then
                    cmb����.Text = "�����"
                Else
                    cmb����.Text = "��ײ�"
                End If
            Else
                cmb����.Text = !����
            End If
            cmb�������@.Text = IIf(Not IsNull(!�������@), !�������@, "")
            cmb��s.Text = IIf(Not IsNull(!��s), !��s, "")
            txt�z�B����.Text = IIf(Not IsNull(!�z�B��]����), !�z�B��]����, "")
            txt�z�B����2.Text = IIf(Not IsNull(!�z�B��]����2), !�z�B��]����2, "")
            txt�o�ד�.Text = IIf(Not IsNull(!�o�ד�), IIf(!�o�ד� <> "", !�o�ד�, "____/__/__"), "____/__/__")
            cmb��z�Ǝ� = IIf(Not IsNull(!��z�Ǝ�), !��z�Ǝ�, "")
            txt�x���ԍ�.Text = IIf(Not IsNull(!�x���ԍ�), !�x���ԍ�, "")
            txt�⍇�ԍ�.Text = IIf(Not IsNull(!�⍇�ԍ�), !�⍇�ԍ�, "")
            txt�d�����z.Value = IIf(Not IsNull(!�d�����z), !�d�����z, 0)
            txt������.Text = IIf(Not IsNull(!������), IIf(!������ <> "", !������, "____/__/__"), "____/__/__")
            txt�P��.Value = IIf(Not IsNull(!�P��), !�P��, 0)
            txt����.Value = IIf(Not IsNull(!����), !����, 0)
            'cmd����.Caption = IIf(Not IsNull(!�����敪), !�����敪, "%")
            txt����.Value = IIf(Not IsNull(!����), !����, 0)
            txt����.Value = IIf(Not IsNull(!����), !����, 0)
            txt�ב��^��.Value = IIf(Not IsNull(!�ב��^��), !�ב��^��, 0)
            txt�ԋ�.Value = IIf(Not IsNull(!�ԋ�), !�ԋ�, 0)
            txt���̑��萔��.Value = IIf(Not IsNull(!���̑��萔��), !���̑��萔��, 0)
            txt���v���z.Text = IIf(Not IsNull(!���v���z), !���v���z, "")
            txt���[�����M.Text = IIf(Not IsNull(!���[�����M), !���[�����M, "")
            cmb������.Text = IIf(Not IsNull(!������), !������, "")
            txt�����ԍ�.Text = IIf(Not IsNull(!Yahoo�����ԍ�), !Yahoo�����ԍ�, "")
            txt���l2.Text = IIf(Not IsNull(!���l1), !���l1, "")
            txt�R�����C�t = IIf(Not IsNull(!�R�����C�tNO), !�R�����C�tNO, "")
            txt�o�ח\���.Text = IIf(Not IsNull(!�o�ח\���), IIf(!�o�ח\��� <> "", !�o�ח\���, "____/__/__"), "____/__/__")
            txt����URL.Text = IIf(Not IsNull(!����URL), !����URL, "")


            G_����_�X�V���� = IIf(IsNull(!�X�V����), Now, !�X�V����)
        
            G_���i�� = cmb���i��.Text
            G_������ = cmb������.Text
        
        End If
                
    End With
    
    With �ڋq�}�X�^RS
        �ڋqID = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�ڋqID)
        
        If �ڋqID = "ID" Then Exit Sub
        If �ڋqID = "" Then Exit Sub
        
        Call �ڋq�}�X�^1���Ǎ�(�ڋq�}�X�^RS, �ڋqID)
        
        If Not �ڋq�}�X�^RS.EOF Then
            G_�ڋq�}�X�^_�r���t���O = True
            chk����1_1.Value = IIf(IsNull(!����1), 0, !����1)
            chk����2_1.Value = IIf(IsNull(!����2), 0, !����2)
            chk����3_1.Value = IIf(IsNull(!����3), 0, !����3)
            chk����4_1.Value = IIf(IsNull(!����4), 0, !����4)
            chk����5_1.Value = IIf(IsNull(!����5), 0, !����5)
            G_�ڋq�}�X�^_�r���t���O = False
        End If
        �ڋq�}�X�^RS.Close
    End With
    
#If 0 Then

    �ڋqID = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�ڋqID)
    
    If �ڋqID = "ID" Then Exit Sub
    
    ' �I�����ꂽ�ڋq�̒����f�[�^���擾����
    Call �ڋq�}�X�^1���Ǎ�(�ڋq�}�X�^RS, �ڋqID)
    G_�ڋq_�X�V���� = IIf(IsNull(�ڋq�}�X�^RS!�X�V����), Now, �ڋq�}�X�^RS!�X�V����)
    �ڋq�}�X�^RS.Close
    
    Call �z����1���Ǎ�(�ڋq�}�X�^RS, �ڋqID)
    G_�z��_�X�V���� = IIf(IsNull(�ڋq�}�X�^RS!�X�V����), Now, �ڋq�}�X�^RS!�X�V����)
    �ڋq�}�X�^RS.Close
    
#End If

End Sub

'************************************************************************
'�@  �\�@���[��������\������B
'************************************************************************
Private Sub ���[������_�\��()
    
    Dim ���[������RS As New ADODB.Recordset
    Dim ����ID As String
    
    ����ID = SpreadGetVal(va�������X�g, G_�������X�g_ROW, COL_����ID)
    
    If ����ID = "" Or ����ID = "����ID" Then
        Exit Sub
    End If
    
    ' �I�����ꂽ�����̃��[���������擾����
    Call ���[����������(����ID, ���[������RS)
    
    With ���[������RS
        If ���[������RS.EOF Then
            va���[������.MaxRows = 0
            txt���[���{�� = ""
        Else
            va���[������.MaxRows = 0
            Do Until .EOF
                va���[������.MaxRows = va���[������.MaxRows + 1
                Call SpreadSetVal(va���[������, va���[������.MaxRows, 1, Format(!���M����, "yyyy/mm/dd hh:mm:ss"))
                Call SpreadSetVal(va���[������, va���[������.MaxRows, 2, !����)
                .MoveNext
            Loop
            .Close
            If va���[������.MaxRows > 0 Then
                Call va���[������_Click(1, 1)
            End If
        End If
    End With
    
End Sub

'************************************************************************
'�@  �\ :�o�b�N�O���E���h�Ōڋq�^�u�ɏ���ݒ肵�����B
'************************************************************************
Private Sub �ڋq�^�u_�\��2()

    Dim �ڋqID As String
    Dim �ڋq�}�X�^RS As New ADODB.Recordset
  
    �ڋqID = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�ڋqID)
    
    ' �ŏ��ɔz�����ǂݍ���
    Call �z����1���Ǎ�(�ڋq�}�X�^RS, �ڋqID)
    
    ' �z���悪�o�^����Ă��Ȃ���΁A�ڋq����ǂݍ���
    If �ڋq�}�X�^RS.EOF Then
        Call �ڋq�}�X�^1���Ǎ�(�ڋq�}�X�^RS, �ڋqID)
    End If
    
    With �ڋq�}�X�^RS
        If �ڋq�}�X�^RS.EOF Then
            txt�ڋqID.Text = �ڋqID
            txt�ڋq��.Text = ""
            txt�t���K�i.Text = ""
            txt�X�֔ԍ�.Text = ""
            opt�j��.Value = True
            opt�j��.Value = True
            txt�Z��_��i.Text = ""
            txt�Z��_���i.Text = ""
            txt�Z��_���i.Text = ""
            txt�d�b�ԍ�.Text = ""
            txt���[��.Text = ""
            txt�y�V���[��.Text = ""
            cmb�A�[�f���N���u.Text = ""
            txt�����.Text = "____/__/__"
            txt�މ��.Text = "____/__/__"
            txt���l.Text = ""
            chk���[�����M.Value = 1
            txt�a����.Text = "____/__/__"
            chk����1.Value = 0
            chk����2.Value = 0
            chk����3.Value = 0
            chk����4.Value = 0
            chk����5.Value = 0
        Else
            txt�ڋqID.Text = !�ڋqID
            txt�ڋq��.Text = !�ڋq��
            txt�t���K�i.Text = !�t���K�i
            txt�X�֔ԍ�.Text = ![��]
            If !���� = "1" Then opt�j��.Value = True Else opt�j��.Value = False
            If !���� = "2" Then opt����.Value = True Else opt����.Value = False
            txt�Z��_��i.Text = !�Z��1
            txt�Z��_���i.Text = !�Z��2
            txt�Z��_���i.Text = IIf(IsNull(!�Z��3), "", !�Z��3)
            txt�d�b�ԍ�.Text = !�d�b�ԍ�
            txt���[��.Text = !���[��
            txt�y�V���[��.Text = !�y�V���[��
            
            If G_�^�uNO = 1 Then
                chk���[�����M.Value = !���[�����M
                If IsNull(!�a����) Or !�a���� = "" Then
                    txt�a����.Text = "____/__/__"
                Else
                    txt�a����.Text = !�a����
                End If
            End If
            
            If G_�^�uNO = 1 Then
                cmb�A�[�f���N���u.Text = !�A�[�f���N���u
                If IsNull(!�����) Or !����� = "" Then
                    txt�����.Text = "____/__/__"
                Else
                    txt�����.Text = !�����
                End If
                
                If IsNull(!�މ��) Or !�މ�� = "" Then
                    txt�މ��.Text = "____/__/__"
                Else
                    txt�މ��.Text = !�މ��
                End If
            End If
            
            txt���l.Text = !���l
            
            If G_�^�uNO = 1 Then
                chk����1.Value = IIf(IsNull(!����1), 0, !����1)
                chk����2.Value = IIf(IsNull(!����2), 0, !����2)
                chk����3.Value = IIf(IsNull(!����3), 0, !����3)
                chk����4.Value = IIf(IsNull(!����4), 0, !����4)
                chk����5.Value = IIf(IsNull(!����5), 0, !����5)
            End If
        End If
    End With
    
    �ڋq�}�X�^RS.Close
    
End Sub

'************************************************************************
'�@  �\�@�ڋq�����͐���
'************************************************************************
Private Sub txt�ڋq��_KeyDown(KeyCode As Integer, Shift As Integer)
    
'    Dim �c�� As String
'    Dim ���O As String
    
'    Dim �J�n�ʒu As Integer
    
'    �J�n�ʒu = InStr(txt�ڋq��.Text, " ")
    
'    If �J�n�ʒu > 0 Then
'        �c�� = Left(txt�ڋq��.Text, �J�n�ʒu - 1)
'        ���O = Mid(txt�ڋq��.Text, �J�n�ʒu + 1)
'        txt�ڋq��.Text = �c�� & "�@" & ���O
'    End If
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�ڋq��_Validate(Cancel As Boolean)
    
    Dim �c�� As String
    Dim ���O As String
    
    Dim �J�n�ʒu As Integer
    
    �J�n�ʒu = InStr(txt�ڋq��.Text, " ")
    
    If �J�n�ʒu > 0 Then
        �c�� = Left(txt�ڋq��.Text, �J�n�ʒu - 1)
        ���O = Mid(txt�ڋq��.Text, �J�n�ʒu + 1)
        txt�ڋq��.Text = �c�� & "�@" & ���O
    End If
    
    Cancel = �ڋq���_�o�^
    
End Sub

Private Sub txt�ڋq��_GotFocus()

    txt�ڋq��.BackColor = vbYellow
    
End Sub

Private Sub txt�ڋq��_LostFocus()
    
    Call �ڋq���_�o�^
    
    txt�ڋq��.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�t���K�i���͐���
'************************************************************************
Private Sub txt�t���K�i_KeyDown(KeyCode As Integer, Shift As Integer)
    
'    Dim �c�� As String
'    Dim ���O As String
    
'    Dim �J�n�ʒu As Integer
    
'    �J�n�ʒu = InStr(txt�t���K�i.Text, " ")
    
'    If �J�n�ʒu > 0 Then
'        �c�� = Left(txt�t���K�i.Text, �J�n�ʒu - 1)
'        ���O = Mid(txt�t���K�i.Text, �J�n�ʒu + 1)
'        txt�t���K�i.Text = �c�� & "�@" & ���O
'    End If
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�t���K�i_Validate(Cancel As Boolean)
    
    Dim �c�� As String
    Dim ���O As String
    
    Dim �J�n�ʒu As Integer
    
    �J�n�ʒu = InStr(txt�t���K�i.Text, " ")
    
    If �J�n�ʒu > 0 Then
        �c�� = Left(txt�t���K�i.Text, �J�n�ʒu - 1)
        ���O = Mid(txt�t���K�i.Text, �J�n�ʒu + 1)
        txt�t���K�i.Text = �c�� & "�@" & ���O
    End If

    Cancel = �ڋq���_�o�^()
    
End Sub

Private Sub txt�t���K�i_GotFocus()

    txt�t���K�i.BackColor = vbYellow
    
End Sub

Private Sub txt�t���K�i_LostFocus()

    txt�t���K�i.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�t���K�i���͐���
'************************************************************************
Private Sub txt�X�֔ԍ�_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�X�֔ԍ�_Validate(Cancel As Boolean)

    Cancel = �ڋq���_�o�^()
    
    Call �X�֔ԍ�����Z����ϊ�����
    
End Sub

Private Sub txt�X�֔ԍ�_GotFocus()

    txt�X�֔ԍ�.BackColor = vbYellow
    
End Sub

Private Sub txt�X�֔ԍ�_LostFocus()

    txt�X�֔ԍ�.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�j�����͐���
'************************************************************************
Private Sub opt�j��_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub opt�j��_Validate(Cancel As Boolean)

    Cancel = �ڋq���_�o�^()

End Sub

Private Sub opt�j��_GotFocus()

    opt�j��.BackColor = vbYellow
    
End Sub

Private Sub opt�j��_LostFocus()

    opt�j��.BackColor = &H8000000F

End Sub

'************************************************************************
'�@  �\�@�������͐���
'************************************************************************
Private Sub opt����_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub opt����_Validate(Cancel As Boolean)

    Cancel = �ڋq���_�o�^()

End Sub

Private Sub opt����_GotFocus()

    opt����.BackColor = vbYellow
    
End Sub

Private Sub opt����_LostFocus()

    opt����.BackColor = &H8000000F

End Sub

'************************************************************************
'�@  �\�@�Z��_��i���͐���
'************************************************************************
Private Sub txt�Z��_��i_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�Z��_��i_Validate(Cancel As Boolean)

    Cancel = �ڋq���_�o�^()
    
End Sub

Private Sub txt�Z��_��i_GotFocus()

    txt�Z��_��i.BackColor = vbYellow
    
End Sub

Private Sub txt�Z��_��i_LostFocus()

    txt�Z��_��i.BackColor = vbWhite

End Sub


'************************************************************************
'�@  �\�@�Z��_���i���͐���
'************************************************************************
Private Sub txt�Z��_���i_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�Z��_���i_Validate(Cancel As Boolean)

    Cancel = �ڋq���_�o�^()
    
End Sub

Private Sub txt�Z��_���i_GotFocus()

    txt�Z��_���i.BackColor = vbYellow
    
End Sub

Private Sub txt�Z��_���i_LostFocus()

    txt�Z��_���i.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�Z��_���i���͐���
'************************************************************************
Private Sub txt�Z��_���i_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�Z��_���i_Validate(Cancel As Boolean)

    Cancel = �ڋq���_�o�^()
    
End Sub

Private Sub txt�Z��_���i_GotFocus()

    txt�Z��_���i.BackColor = vbYellow
    
End Sub

Private Sub txt�Z��_���i_LostFocus()

    txt�Z��_���i.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�d�b�ԍ����͐���
'************************************************************************
Private Sub txt�d�b�ԍ�_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�d�b�ԍ�_Validate(Cancel As Boolean)

    Cancel = �ڋq���_�o�^()
    
End Sub

Private Sub txt�d�b�ԍ�_GotFocus()

    txt�d�b�ԍ�.BackColor = vbYellow
    
End Sub

Private Sub txt�d�b�ԍ�_LostFocus()

    txt�d�b�ԍ�.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@���[�����͐���
'************************************************************************
Private Sub txt���[��_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt���[��_Validate(Cancel As Boolean)
    
    txt���[��.Text = Trim(txt���[��.Text)
    
    Cancel = �ڋq���_�o�^()
    
End Sub

Private Sub txt���[��_GotFocus()

    txt���[��.BackColor = vbYellow
    
End Sub

Private Sub txt���[��_LostFocus()

    txt���[��.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�y�V���[�����͐���
'************************************************************************
Private Sub txt�y�V���[��_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�y�V���[��_Validate(Cancel As Boolean)
    
    txt�y�V���[��.Text = Trim(txt�y�V���[��.Text)
    
    Cancel = �ڋq���_�o�^()
    
End Sub

Private Sub txt�y�V���[��_GotFocus()

    txt�y�V���[��.BackColor = vbYellow
    
End Sub

Private Sub txt�y�V���[��_LostFocus()

    txt�y�V���[��.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@���[�����M���͐���
'************************************************************************
Private Sub chk���[�����M_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub chk���[�����M_Validate(Cancel As Boolean)

    Cancel = �ڋq���_�o�^()
    
End Sub

Private Sub chk���[�����M_GotFocus()

    chk���[�����M.BackColor = vbYellow
    
End Sub

Private Sub chk���[�����M_LostFocus()

    chk���[�����M.BackColor = &H8000000F

End Sub

'************************************************************************
'�@  �\�@�A�[�f���N���u���͐���
'************************************************************************
Private Sub cmb�A�[�f���N���u_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub cmb�A�[�f���N���u_Validate(Cancel As Boolean)

    If Len(cmb�A�[�f���N���u.Text) > 10 Then
        Call MsgBox("�A�[�f���N���u���������܂��B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If

    Cancel = �ڋq���_�o�^()
    
End Sub

Private Sub cmb�A�[�f���N���u_GotFocus()

    cmb�A�[�f���N���u.BackColor = vbYellow
    
End Sub

Private Sub cmb�A�[�f���N���u_LostFocus()

    cmb�A�[�f���N���u.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@��������͐���
'************************************************************************
Private Sub txt�����_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�����_Validate(Cancel As Boolean)
        
    If txt�����.Text <> "____/__/__" Then
        If IsDate(txt�����.Text) = False Then
            Call MsgBox("���������������͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
            Cancel = True
            Exit Sub
        End If
    End If
    
    Cancel = �ڋq���_�o�^()

End Sub

Private Sub txt�����_GotFocus()

    txt�����.BackColor = vbYellow
    
End Sub

Private Sub txt�����_LostFocus()

    txt�����.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�މ�����͐���
'************************************************************************
Private Sub txt�މ��_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�މ��_Validate(Cancel As Boolean)

    If txt�މ��.Text <> "____/__/__" Then
        If IsDate(txt�މ��.Text) = False Then
            Call MsgBox("�������މ������͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
            Cancel = True
            Exit Sub
        End If
    End If
    
    Cancel = �ڋq���_�o�^()

End Sub

Private Sub txt�މ��_GotFocus()

    txt�މ��.BackColor = vbYellow
    
End Sub

Private Sub txt�މ��_LostFocus()

    txt�މ��.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@���l���͐���
'************************************************************************
Private Sub txt���l_KeyDown(KeyCode As Integer, Shift As Integer)
    
    ' ���l�̓}���`���C�����͂Ȃ̂ŁA���s�L�[����������Ă��t�B�[���h���ړ����Ȃ��B
    'Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt���l_Validate(Cancel As Boolean)

    If Len(txt���l) >= 4096 Then
        Call MsgBox("���l�̓��͌������傫���ł��B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If

    Cancel = �ڋq���_�o�^()
    
End Sub

Private Sub txt���l_GotFocus()

    txt���l.BackColor = vbYellow
    
End Sub

Private Sub txt���l_LostFocus()

    txt���l.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�a�������͐���
'************************************************************************
Private Sub txt�a����_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�a����_Validate(Cancel As Boolean)

    If txt�a����.Text <> "____/__/__" Then
        If IsDate(txt�a����.Text) = False Then
            Call MsgBox("�������a��������͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
            Cancel = True
            Exit Sub
        End If
    End If
    
    Cancel = �ڋq���_�o�^()

End Sub

Private Sub txt�a����_GotFocus()

    txt�a����.BackColor = vbYellow
    
End Sub

Private Sub txt�a����_LostFocus()

    txt�a����.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�����P�{�^���N���b�N
'************************************************************************
Private Sub chk����1_Click()

    If G_�ڋq�}�X�^_�r���t���O = False Then
        Call �ڋq���_�o�^
    End If
    
End Sub

'************************************************************************
'�@  �\�@�����Q�{�^���N���b�N
'************************************************************************
Private Sub chk����2_Click()

    If G_�ڋq�}�X�^_�r���t���O = False Then
        Call �ڋq���_�o�^
    End If

End Sub

'************************************************************************
'�@  �\�@�����R�{�^���N���b�N
'************************************************************************
Private Sub chk����3_Click()

    If G_�ڋq�}�X�^_�r���t���O = False Then
        Call �ڋq���_�o�^
    End If

End Sub

'************************************************************************
'�@  �\�@�����S�{�^���N���b�N
'************************************************************************
Private Sub chk����4_Click()

    If G_�ڋq�}�X�^_�r���t���O = False Then
        Call �ڋq���_�o�^
    End If

End Sub

'************************************************************************
'�@  �\�@�����T�{�^���N���b�N
'************************************************************************
Private Sub chk����5_Click()

    If G_�ڋq�}�X�^_�r���t���O = False Then
        Call �ڋq���_�o�^
    End If

End Sub

'************************************************************************
'�@  �\�@�����P�{�^���N���b�N
'************************************************************************
Private Sub chk����1_1_Click()
    
    Dim �ڋqID As String

    If G_�ڋq�}�X�^_�r���t���O = False Then
        �ڋqID = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�ڋqID)
        Call �����`�F�b�N�X�V1(�ڋqID, chk����1_1.Value)
    End If
    
End Sub

'************************************************************************
'�@  �\�@�����Q�{�^���N���b�N
'************************************************************************
Private Sub chk����2_1_Click()

    Dim �ڋqID As String

    If G_�ڋq�}�X�^_�r���t���O = False Then
        �ڋqID = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�ڋqID)
        Call �����`�F�b�N�X�V2(�ڋqID, chk����2_1.Value)
    End If

End Sub

'************************************************************************
'�@  �\�@�����R�{�^���N���b�N
'************************************************************************
Private Sub chk����3_1_Click()

    Dim �ڋqID As String

    If G_�ڋq�}�X�^_�r���t���O = False Then
        �ڋqID = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�ڋqID)
        Call �����`�F�b�N�X�V3(�ڋqID, chk����3_1.Value)
    End If

End Sub

'************************************************************************
'�@  �\�@�����S�{�^���N���b�N
'************************************************************************
Private Sub chk����4_1_Click()

    Dim �ڋqID As String

    If G_�ڋq�}�X�^_�r���t���O = False Then
        �ڋqID = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�ڋqID)
        Call �����`�F�b�N�X�V4(�ڋqID, chk����4_1.Value)
    End If

End Sub

'************************************************************************
'�@  �\�@�����T�{�^���N���b�N
'************************************************************************
Private Sub chk����5_1_Click()
    
    Dim �ڋqID As String

    If G_�ڋq�}�X�^_�r���t���O = False Then
        �ڋqID = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�ڋqID)
        Call �����`�F�b�N�X�V5(�ڋqID, chk����5_1.Value)
    End If

End Sub

'************************************************************************
'�@  �\�@�]�L�{�^��
'************************************************************************
Private Sub cmd�]�L_Click()

    Dim �ڋqID As String
    Dim �ڋq�}�X�^RS As New ADODB.Recordset
  
    �ڋqID = txt�ڋqID.Text
    
    Call �ڋq�}�X�^1���Ǎ�(�ڋq�}�X�^RS, �ڋqID)
    
    With �ڋq�}�X�^RS
        If �ڋq�}�X�^RS.EOF Then
            txt�ڋq��.Text = ""
            txt�t���K�i.Text = ""
            txt�X�֔ԍ�.Text = ""
            opt�j��.Value = True
            opt�j��.Value = True
            txt�Z��_��i.Text = ""
            txt�Z��_���i.Text = ""
            txt�Z��_���i.Text = ""
            txt�d�b�ԍ�.Text = ""
            txt���[��.Text = ""
        Else
            txt�ڋq��.Text = !�ڋq��
            txt�t���K�i.Text = !�t���K�i
            txt�X�֔ԍ�.Text = ![��]
            If !���� = "1" Then opt�j��.Value = True Else opt�j��.Value = False
            If !���� = "2" Then opt����.Value = True Else opt����.Value = False
            txt�Z��_��i.Text = !�Z��1
            txt�Z��_���i.Text = !�Z��2
            txt�Z��_���i.Text = IIf(IsNull(!�Z��3), "", !�Z��3)
            txt�d�b�ԍ�.Text = !�d�b�ԍ�
            txt���[��.Text = !���[��
            txt�ڋq��.SetFocus
        End If
    End With
    
    �ڋq�}�X�^RS.Close
    
End Sub

'************************************************************************
'�@  �\�@�󒍓����͐���
'************************************************************************
Private Sub txt�󒍓�_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�󒍓�_Validate(Cancel As Boolean)
    
    If txt�󒍓�.Text = "____/__/__" Then
        Call MsgBox("�󒍓�����͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If
    
    If IsDate(txt�󒍓�.Text) = False Then
        Call MsgBox("�������󒍓�����͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If
    
    Cancel = ����_�X�V()

End Sub

Private Sub txt�󒍓�_GotFocus()

    txt�󒍓�.BackColor = vbYellow
    
End Sub

Private Sub txt�󒍓�_LostFocus()

    txt�󒍓�.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�X�e�[�^�X���͐���
'************************************************************************
Private Sub cmb�X�e�[�^�X_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub cmb�X�e�[�^�X_Validate(Cancel As Boolean)
    
    If Len(cmb�X�e�[�^�X.Text) > 20 Then
        Call MsgBox("�X�e�[�^�X�̕����񂪒������܂��B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If
    
    If cmb�X�e�[�^�X.Text = "�o�׊���" Then
        If txt�o�ד�.Text = "____/__/__" Then
            txt�o�ד�.Text = Format(Date, "yyyy/mm/dd")
        End If
    End If
    
    cmb����.Text = "���̑�"
    
    If cmb�X�e�[�^�X.Text = "�R�����C�t" Then
        cmb����.Text = "��ײ�"
    End If
    
    If �A�[�f������(cmb���i��.Text) = 1 Or �A�[�f������(cmb���i��.Text) = 9 Then
        cmb����.Text = "�����"
    End If
    
    Cancel = ����_�X�V()
    
End Sub

Private Sub cmb�X�e�[�^�X_GotFocus()

    cmb�X�e�[�^�X.BackColor = vbYellow
    
End Sub

Private Sub cmb�X�e�[�^�X_LostFocus()

    cmb�X�e�[�^�X.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@���i�����͐���
'************************************************************************
Private Sub cmb���i��_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub cmb���i��_Validate(Cancel As Boolean)
    
    Dim �ݐϖ{�� As Long
    Dim ���i�}�X�^RS As New ADODB.Recordset
    
    If Len(cmb���i��.Text) >= 100 Then
        Call MsgBox("���i�����������܂��B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If
    
    If cmb������.Text = "" Then
        Call MsgBox("��������I�����ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
        Exit Sub
    End If
    
    If cmb���i��.Text = G_���i�� Then
        Exit Sub
    End If
    
    G_���i�� = cmb���i��.Text
    
    Call ���i�}�X�^�擾(cmb���i��.Text, ���i�}�X�^RS)
    
    With ���i�}�X�^RS
        If Not ���i�}�X�^RS.EOF Then
            'If cmb������.Text = "Yahoo" Or cmb������.Text = "�y�V" Or cmb������.Text = "������̂��l�b�g" Or cmb������.Text = "�A�}�]��" Or cmb������.Text = "����ۂۃ��[��" Then
                
                �ݐϖ{�� = �ݐϐ��v�Z2(txt����ID.Text)
                
#If 0 Then
                If cmb���i��.Text = "�A�[�f��" Then
                    
                    If �ݐϖ{�� < 1 Then
                        txt�P��.Value = !�P��
                        cmd����.Caption = "\"
                        txt����.Value = !�������z
                        txt����.Value = !����
                        txt����.Value = 1
                        txt���̑��萔��.Value = 0
                        txt�ԋ�.Value = 0
                        
                    ElseIf �ݐϖ{�� >= 1 And �ݐϖ{�� <= 5 Then
                        txt�P��.Value = !�P��
                        cmd����.Caption = "%"
                        txt����.Value = 10
                        txt����.Value = 0
                        txt����.Value = 1
                        txt���̑��萔��.Value = 0
                        txt�ԋ�.Value = 0
                    
                    ElseIf �ݐϖ{�� >= 6 Then
                         txt�P��.Value = !�P��
                        cmd����.Caption = "%"
                        txt����.Value = 20
                        txt����.Value = 0
                        txt����.Value = 1
                        txt���̑��萔��.Value = 0
                        txt�ԋ�.Value = 0
                   End If
                Else
#End If
                    txt�P��.Value = !�P��
                    'cmd����.Caption = "\"
                    txt����.Value = !�������z
                    txt����.Value = !����
                    txt����.Value = 1
                    txt���̑��萔��.Value = 0
                    txt�ԋ�.Value = 0
                'End If
            'Else
            '    txt�P�� = !�P��
            '    'cmd����.Caption = "%"
            '    txt����.Value = 0
            '    txt����.Text = !����
            '    txt����.Value = 1
            '    txt���̑��萔��.Value = 0
            '    txt�ԋ�.Value = 0
            'End If
        End If
        .Close
    End With
    
    Call �Čv�Z
    
    If cmb���i��.Text = "�A�[�f��" _
        Or cmb���i��.Text = "�A�[�f��2�{�Z�b�g" _
        Or cmb���i��.Text = "�A�[�f���{�V�����v�[" _
        Or cmb���i��.Text = "�V�u�X�^" _
        Or cmb���i��.Text = "�V�u�X�^�{�V�����v�[" _
        Or cmb���i��.Text = "�u�[�X�^�[" _
        Or cmb���i��.Text = "�u�[�X�^�[�i�v���ь��ԁj" _
        Or cmb���i��.Text = "�u�[�X�^�[�{�V�����v�[" _
        Or cmb���i��.Text = "�V�n�C�u���b�^�[" _
        Or cmb���i��.Text = "�V�n�C�u���b�^�[�{�V�����v�[" _
        Or cmb���i��.Text = "�n�C�u���b�h" _
        Or cmb���i��.Text = "�n�C�u���b�h�{�V�����v�[" _
        Or cmb���i��.Text = "�i�C�X���f�B�[" _
        Or cmb���i��.Text = "�i�C�X���f�B�[�{�V�����v�[" _
        Or cmb���i��.Text = "�V�����v�[" _
        Or cmb���i��.Text = "�V�����v�[�i�v���[���g�j" _
        Or cmb���i��.Text = "�V�����v�[2�{�Z�b�g" _
        Or cmb���i��.Text = "�V�����v�[�{�g���[�g�����g" _
        Or cmb���i��.Text = "�g���[�g�����g" _
        Or cmb���i��.Text = "�g���[�g�����g�i�v���[���g�j" _
        Or cmb���i��.Text = "�A�[�f�����V�����v�[�����i" _
        Or cmb���i��.Text = "�A�[�f�������i" _
        Or cmb���i��.Text = "�V�����v�[�����i" Then
        cmb��z�Ǝ�.Text = "����}��"
    ElseIf cmb���i��.Text = "�A�[�f������" _
        Or cmb���i��.Text = "�~�j�܂�" Then
        cmb��z�Ǝ�.Text = "EXPRESS"
    Else
        cmb��z�Ǝ�.Text = "�N���l�R���}�g"
    End If

    Cancel = ����_�X�V()
    
End Sub

Private Sub cmb���i��_GotFocus()

    cmb���i��.BackColor = vbYellow
    
End Sub

Private Sub cmb���i��_LostFocus()

    cmb���i��.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@���吧��
'************************************************************************
Private Sub cmb����_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub cmb����_Validate(Cancel As Boolean)

    If Len(cmb����.Text) > 10 Then
        Call MsgBox("���傪�������܂��B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If
    
    Cancel = ����_�X�V()
    
End Sub

Private Sub cmb����_GotFocus()

    cmb����.BackColor = vbYellow
    
End Sub

Private Sub cmb����_LostFocus()

    cmb����.BackColor = vbWhite

End Sub


'************************************************************************
'�@  �\�@�������@���͐���
'************************************************************************
Private Sub cmb�������@_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub cmb�������@_Validate(Cancel As Boolean)

    If Len(cmb�������@.Text) > 20 Then
        Call MsgBox("�������@���������܂��B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If
    
    If cmb�������@.Text = "�R���r�j" Then
        txt����URL.Caption = "����URL"
    Else
        txt����URL.Caption = "����ID"
    End If
    
    Cancel = ����_�X�V()
    
End Sub

Private Sub cmb�������@_GotFocus()

    cmb�������@.BackColor = vbYellow
    
End Sub

Private Sub cmb�������@_LostFocus()

    cmb�������@.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@��s���͐���
'************************************************************************
Private Sub cmb��s_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub cmb��s_Validate(Cancel As Boolean)

    If Len(cmb��s.Text) > 10 Then
        Call MsgBox("��s�����������܂��B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If
    
    Cancel = ����_�X�V()
    
End Sub

Private Sub cmb��s_GotFocus()

    cmb��s.BackColor = vbYellow
    
End Sub

Private Sub cmb��s_LostFocus()

    cmb��s.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�z�B�������͐���
'************************************************************************
Private Sub txt�z�B����_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�z�B����_Validate(Cancel As Boolean)
    
    Cancel = ����_�X�V()
    
End Sub

Private Sub txt�z�B����_GotFocus()

    txt�z�B����.BackColor = vbYellow
    
End Sub

Private Sub txt�z�B����_LostFocus()

    txt�z�B����.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�z�B����2���͐���
'************************************************************************
Private Sub txt�z�B����2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�z�B����2_Validate(Cancel As Boolean)
    
    Cancel = ����_�X�V()
    
End Sub

Private Sub txt�z�B����2_GotFocus()

    txt�z�B����2.BackColor = vbYellow
    
End Sub

Private Sub txt�z�B����2_LostFocus()

    txt�z�B����2.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�o�ד����͐���
'************************************************************************
Private Sub txt�o�ד�_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�o�ד�_Validate(Cancel As Boolean)

    If txt�o�ד�.Text <> "____/__/__" Then
        If IsDate(txt�o�ד�.Text) = False Then
            Call MsgBox("�������o�ד�����͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
            Cancel = True
            Exit Sub
        End If
    End If
    
    Cancel = ����_�X�V()

End Sub

Private Sub txt�o�ד�_GotFocus()

    txt�o�ד�.BackColor = vbYellow
    
End Sub

Private Sub txt�o�ד�_LostFocus()

    txt�o�ד�.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�o�ד��ݒ�
'************************************************************************
Private Sub cmd�{��1_Click()
    
    txt�o�ד�.Text = Format(Date, "yyyy/mm/dd")
    
    cmb�X�e�[�^�X.Text = "�o�׊���"

    Call ����_�X�V
    
End Sub

'************************************************************************
'�@  �\�@��z�Ǝғ��͐���
'************************************************************************
Private Sub cmb��z�Ǝ�_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub cmb��z�Ǝ�_Validate(Cancel As Boolean)

    If Len(cmb��z�Ǝ�.Text) > 10 Then
        Call MsgBox("��z�Ǝ҂��������܂��B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If

    Cancel = ����_�X�V()
    
End Sub

Private Sub cmb��z�Ǝ�_GotFocus()

    cmb��z�Ǝ�.BackColor = vbYellow
    
End Sub

Private Sub cmb��z�Ǝ�_LostFocus()

    cmb��z�Ǝ�.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�x���ԍ����͐���
'************************************************************************
Private Sub txt�x���ԍ�_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�x���ԍ�_Validate(Cancel As Boolean)

    Cancel = ����_�X�V()
    
End Sub

Private Sub txt�x���ԍ�_GotFocus()

    txt�x���ԍ�.BackColor = vbYellow
    
End Sub

Private Sub txt�x���ԍ�_LostFocus()

    txt�x���ԍ�.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�⍇�ԍ����͐���
'************************************************************************
Private Sub txt�⍇�ԍ�_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�⍇�ԍ�_Validate(Cancel As Boolean)

    If txt�⍇�ԍ�.Text <> "" Then
        
        ' cmb�X�e�[�^�X.Text = "�o�׊���"
        
        If txt�o�ד�.Text = "____/__/__" Then
            txt�o�ד�.Text = Format(Date, "yyyy/mm/dd")
        End If
    End If

    Cancel = ����_�X�V()
    
End Sub

Private Sub txt�⍇�ԍ�_GotFocus()

    txt�⍇�ԍ�.BackColor = vbYellow
    
End Sub

Private Sub txt�⍇�ԍ�_LostFocus()

    txt�⍇�ԍ�.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�d�����z���͐���
'************************************************************************
Private Sub txt�d�����z_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�d�����z_Validate(Cancel As Boolean)
    
    Call �Čv�Z
    
    Cancel = ����_�X�V()

End Sub

Private Sub txt�d�����z_GotFocus()

    txt�d�����z.BackColor = vbYellow
    
    Call psubIMEOnOff(Me.hwnd, False)
    
End Sub

Private Sub txt�d�����z_LostFocus()
    
    txt�d�����z.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@����URL���͐���
'************************************************************************
Private Sub txt����URL_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt����URL_Validate(Cancel As Boolean)

    Cancel = ����_�X�V()
    
End Sub

Private Sub txt����URL_GotFocus()

    txt����URL.BackColor = vbYellow
    
End Sub

Private Sub txt����URL_LostFocus()

    txt����URL.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@���������͐���
'************************************************************************
Private Sub txt������_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt������_Validate(Cancel As Boolean)
   
    Call �Čv�Z
   
    If txt������.Text <> "____/__/__" Then
        If IsDate(txt������.Text) = False Then
            Call MsgBox("����������������͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
            Cancel = True
            Exit Sub
        End If
    End If
    
    Cancel = ����_�X�V()

End Sub

Private Sub txt������_GotFocus()

    txt������.BackColor = vbYellow
    
End Sub

Private Sub txt������_LostFocus()
    
    txt������.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�������ݒ�
'************************************************************************
Private Sub cmd�{��2_Click()
    
    txt������.Text = Format(Date, "yyyy/mm/dd")
    
    Call ����_�X�V
    
End Sub

'************************************************************************
'�@  �\�@�P�����͐���
'************************************************************************
Private Sub txt�P��_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�P��_Validate(Cancel As Boolean)

    If txt�P��.Value < 0 Then
        Call MsgBox("�v���X�l����͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If
    
    Call �Čv�Z
    
    Cancel = ����_�X�V()

End Sub

Private Sub txt�P��_GotFocus()

    txt�P��.BackColor = vbYellow
    
    Call psubIMEOnOff(Me.hwnd, False)
    
End Sub

Private Sub txt�P��_LostFocus()
    
    txt�P��.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�������͐���
'************************************************************************
Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt����_Validate(Cancel As Boolean)

    If txt����.Value > 0 Then
        Call MsgBox("�}�C�i�X�l����͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If

    Call �Čv�Z

    Cancel = ����_�X�V()

End Sub

Private Sub txt����_GotFocus()

    txt����.BackColor = vbYellow
    
    Call psubIMEOnOff(Me.hwnd, False)
    
End Sub

Private Sub txt����_LostFocus()
    
    txt����.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�������͐���
'************************************************************************
Private Sub cmd����_Click()
    
    txt����.Text = -770
    
    Call �Čv�Z
    
End Sub

'************************************************************************
'�@  �\�@�������͐���
'************************************************************************
Private Sub cmd����2_Click()
    
    txt����.Text = -4600
    
    Call �Čv�Z
    
End Sub

'************************************************************************
'�@  �\�@�������͐���
'************************************************************************
Private Sub cmd����3_Click()
    
    Dim �P��
    
    If IsNumeric(txt�P��.Text) Then
        �P�� = CLng(txt�P��.Text)
    Else
        �P�� = 0
    End If
    
    txt����.Text = CLng(Format(((�P�� * 10) / 100), "0")) * -1
    
    Call �Čv�Z
    
End Sub

'************************************************************************
'�@  �\�@�������͐���
'************************************************************************
Private Sub cmd����4_Click()
    
    Dim �P��
    
    If IsNumeric(txt�P��.Text) Then
        �P�� = CLng(txt�P��.Text)
    Else
        �P�� = 0
    End If
    
    txt����.Text = CLng(Format(((�P�� * 20) / 100), "0")) * -1
    
    Call �Čv�Z
    
End Sub

'************************************************************************
'�@  �\�@�������͐���
'************************************************************************
Private Sub cmd����5_Click()
    
    txt����.Text = 0
    
    Call �Čv�Z
    
End Sub

'************************************************************************
'�@  �\�@���ʓ��͐���
'************************************************************************
Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt����_Validate(Cancel As Boolean)

    If txt����.Value < 0 Then
        Call MsgBox("�v���X�l����͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If

    Call �Čv�Z
    
    Cancel = ����_�X�V()

End Sub

Private Sub txt����_GotFocus()

    txt����.BackColor = vbYellow
    
    Call psubIMEOnOff(Me.hwnd, False)
    
End Sub

Private Sub txt����_LostFocus()
    
    txt����.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�������͐���
'************************************************************************
Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt����_Validate(Cancel As Boolean)
    
    If txt����.Value < 0 Then
        Call MsgBox("�v���X�l����͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If
    
    Call �Čv�Z
    
    Cancel = ����_�X�V()

End Sub

Private Sub txt����_GotFocus()

    txt����.BackColor = vbYellow
    
    Call psubIMEOnOff(Me.hwnd, False)
    
End Sub

Private Sub txt����_LostFocus()
    
    txt����.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�ב��^�����͐���
'************************************************************************
Private Sub txt�ב��^��_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�ב��^��_Validate(Cancel As Boolean)
    
    Call �Čv�Z
    
    Cancel = ����_�X�V()

End Sub

Private Sub txt�ב��^��_GotFocus()

    txt�ב��^��.BackColor = vbYellow
    
    Call psubIMEOnOff(Me.hwnd, False)
    
End Sub

Private Sub txt�ב��^��_LostFocus()
    
    txt�ב��^��.BackColor = vbWhite

End Sub
'************************************************************************
'�@  �\�@�ԋ����͐���
'************************************************************************
Private Sub txt�ԋ�_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�ԋ�_Validate(Cancel As Boolean)

    If txt�ԋ�.Value > 0 Then
        Call MsgBox("�}�C�i�X�l����͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If

    Call �Čv�Z
    
    Cancel = ����_�X�V()

End Sub

Private Sub txt�ԋ�_GotFocus()

    txt�ԋ�.BackColor = vbYellow
        
    Call psubIMEOnOff(Me.hwnd, False)

End Sub

Private Sub txt�ԋ�_LostFocus()
    
    txt�ԋ�.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@���̑��萔�����͐���
'************************************************************************
Private Sub txt���̑��萔��_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt���̑��萔��_Validate(Cancel As Boolean)

    If txt���̑��萔��.Value > 0 Then
        Call MsgBox("�}�C�i�X�l����͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If
    
    Call �Čv�Z
    
    Cancel = ����_�X�V()

End Sub

Private Sub txt���̑��萔��_GotFocus()

    txt���̑��萔��.BackColor = vbYellow
    
    Call psubIMEOnOff(Me.hwnd, False)
    
End Sub

Private Sub txt���̑��萔��_LostFocus()

    Call �Čv�Z
    
    Call ����_�X�V
    
    txt���̑��萔��.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@���[�����M���͐���
'************************************************************************
Private Sub txt���[�����M_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt���[�����M_Validate(Cancel As Boolean)

    Cancel = ����_�X�V()
    
End Sub

Private Sub txt���[�����M_GotFocus()

    txt���[�����M.BackColor = vbYellow
    
End Sub

Private Sub txt���[�����M_LostFocus()

    txt���[�����M.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@���������͐���
'************************************************************************
Private Sub cmb������_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub cmb������_Validate(Cancel As Boolean)

    Dim Cancel2 As Boolean
    
    If Len(cmb������.Text) > 10 Then
        Call MsgBox("���������������܂��B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If
    
    If G_������ <> cmb������.Text Then
        G_������ = cmb������.Text
        G_���i�� = ""
    End If
    
    If cmb������.Text = "Yahoo" Or cmb������.Text = "�y�V" Then
    Else
        If txt�����ԍ�.Text = "" Then
            If cmb������.Text = "�����g���b�N�X" Then
'            If cmb������.Text = "������̂��l�b�g" Then
                txt�����ԍ�.Text = ""
            ElseIf cmb������.Text = "�A�}�]��" Then
                txt�����ԍ�.Text = ""
            ElseIf cmb������.Text = "�R�}�`" Then
                txt�����ԍ�.Text = "KOMACHI-" & Format(Now, "yyyymmddhhmmss")
            Else
                txt�����ԍ�.Text = "ETC-" & Format(Now, "yyyymmddhhmmss")
            End If
        End If
        
    End If
    
    Call cmb���i��_Validate(Cancel2)
    
    Cancel = ����_�X�V()
    
End Sub

Private Sub cmb������_GotFocus()

    cmb������.BackColor = vbYellow
    
End Sub

Private Sub cmb������_LostFocus()

    'cmb������.BackColor = vbWhite
    cmb������.BackColor = vbRed

End Sub

'************************************************************************
'�@  �\�@�����ԍ����͐���
'************************************************************************
Private Sub txt�����ԍ�_KeyDown(KeyCode As Integer, Shift As Integer)
        
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�����ԍ�_Validate(Cancel As Boolean)

    Cancel = ����_�X�V()
    
End Sub

Private Sub txt�����ԍ�_GotFocus()

    txt�����ԍ�.BackColor = vbYellow
    
End Sub

Private Sub txt�����ԍ�_LostFocus()

    txt�����ԍ�.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�R�����C�tNO���͐���
'************************************************************************
Private Sub txt�R�����C�t_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�R�����C�t_Validate(Cancel As Boolean)

    Cancel = ����_�X�V()
    
End Sub

Private Sub txt�R�����C�t_GotFocus()

    txt�R�����C�t.BackColor = vbYellow
    
End Sub

Private Sub txt�R�����C�t_LostFocus()

    txt�R�����C�t.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@�o�ח\������͐���
'************************************************************************
Private Sub txt�o�ח\���_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt�o�ח\���_Validate(Cancel As Boolean)

    If txt�o�ח\���.Text <> "____/__/__" Then
        If IsDate(txt�o�ח\���.Text) = False Then
            Call MsgBox("�������o�ד��\�������͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
            Cancel = True
            Exit Sub
        End If
    End If
    
    Cancel = ����_�X�V()
    
End Sub

Private Sub txt�o�ח\���_GotFocus()

    txt�o�ח\���.BackColor = vbYellow
    
End Sub

Private Sub txt�o�ח\���_LostFocus()

    txt�o�ח\���.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\�@���l2���͐���
'************************************************************************
Private Sub txt���l2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    ' ���l�̓}���`���C�����͂Ȃ̂ŁA���s�L�[����������Ă��t�B�[���h���ړ����Ȃ��B
    'Call �^�u�L�[���M(KeyCode)

End Sub

Private Sub txt���l2_Validate(Cancel As Boolean)
    
    If Len(txt���l2) >= 4096 Then
        Call MsgBox("���l�̓��͌������傫���ł��B", vbOKOnly, "�ڋq�Ǘ�")
        Cancel = True
        Exit Sub
    End If

    Cancel = ����_�X�V()
    
End Sub

Private Sub txt���l2_GotFocus()

    txt���l2.BackColor = vbYellow
    
End Sub

Private Sub txt���l2_LostFocus()

    txt���l2.BackColor = vbWhite

End Sub

'************************************************************************
'�@  �\ :�^�u�L�[�𑗐M����
'************************************************************************
Public Sub �^�u�L�[���M(KeyCode As Integer)
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyDown) Then
        Me.Tag = "Through"
        SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyUp Then
        Me.Tag = "Through"
        SendKeys "+{TAB}"
    End If
End Sub

'************************************************************************
'�@  �\�@�����ڍׂ��ύX���p��␳����B
'************************************************************************
Private Sub �Čv�Z()

    Dim ���i�� As String
    Dim �A�[�f���N���u As String
    Dim �P�� As Long
    Dim ���� As String
    Dim ���� As Long
    Dim ���z As Long
    Dim ���� As Long
    Dim �ԋ� As Long
    Dim ���̑��萔�� As Long
    Dim ���v���z As Long

#If 0 Then
    ���i�� = cmb���i��.Text
    �A�[�f���N���u = cmb�A�[�f���N���u.Text
    
    If ���i�� = "�A�[�f��" Then txt�P��.Value = 15750
    
    If ���i�� = "�X�[�p�[�A�[�f��" Then txt�P��.Value = 15750
    
    If ���i�� = "�A�[�f��2�{�Z�b�g" Then txt�P��.Value = 31500
    
    If ���i�� = "�A�[�f���{�V�����v�[" Then txt�P��.Value = 17755
    
    If ���i�� = "�A�[�f���{�V�����v�[�����i" Then txt�P��.Value = 16000
    
    If ���i�� = "�A�[�f��(�Z�[��)" Then txt�P��.Value = 9400
    
    If ���i�� = "�V�����v�[" Then txt�P��.Value = 2940
    
    If ���i�� = "�V�����v�[2�{�Z�b�g" Then txt�P��.Value = 5880
    
    If ���i�� = "�A�[�f�������i" Then txt�P��.Value = 1000
    
    If ���i�� = "�V�����v�[�����i" Then txt�P��.Value = 525
    
    If ���i�� = "�A�[�f�������i�{�V�����v�[�����i" Then txt�P��.Value = 1500
    
    If ���i�� = "�A�[�f���T�v��" Then txt�P��.Value = 9800
    
    If ���i�� = "���Q�C���Q��" Then txt�P��.Value = 3500
    
    If ���i�� = "���Q�C���T��" Then txt�P��.Value = 3500
    
    If ���i�� = "�`�[�Y�X�C�[�g�z�[���^����" Then txt�P��.Value = 3980
    
    If ���i�� = "�u�[�X�^�[" Then txt�P��.Value = 10000
    
    If ���i�� = "�u�[�X�^�[�i�v���ь��ԁj" Then txt�P��.Value = 0
    
    If ���i�� = "�n�C�u�b�h" Then txt�P��.Value = 12600
    
    If ���i�� = "���킠�퐅�f��" Then txt�P��.Value = 2980
    
    If ���i�� = "���킠�퐅�f��2�{�Z�b�g" Then txt�P��.Value = 5960

#End If

    If ���i�� = "�A�[�f������" Then
        txt�P��.Value = 0
        cmb�X�e�[�^�X.Text = "��������"
        cmb�������@.Text = "��������"
    End If
    
    If ���i�� = "�~�j�܂�" Then
        txt�P��.Value = 0
        cmb�X�e�[�^�X.Text = "��������"
        cmb�������@.Text = "��������"
    End If

    �P�� = txt�P��.Value
    ���� = txt����.Value
    ���� = txt����.Value
    ���� = txt����.Value
    �ԋ� = txt�ԋ�.Value
    ���z = CLng(Format(((�P�� + ����) * ����), "0"))
    
#If 0 Then
    If cmd����.Caption = "%" Then
        If ���� > 0 And ���� < 100 Then
            ���z = CLng(Format(((�P�� * (100 - ����)) / 100 * ����), "0"))
        Else
            ���z = �P�� * ����
        End If
    Else
        If ���� <> 0 Then
            ���z = CLng(Format(((�P�� + ����) * ����), "0"))
        Else
            ���z = �P�� * ����
        End If
    End If
#End If

    ���̑��萔�� = txt���̑��萔��.Value
    ���v���z = ���z + ���� + �ԋ� + ���̑��萔��

    txt���v���z.Text = ���v���z
    
    Call ����_�X�V
    
End Sub

'************************************************************************
'�@  �\ :�ڋq���X�g�őI������Ă���s���X�V����B
'************************************************************************
Private Function �ڋq���_�o�^() As Boolean
    
    �ڋq���_�o�^ = True
    
    Select Case G_�^�uNO
        Case 1
            �ڋq���_�o�^ = �ڋq���_�o�^_sub()
        Case 2
            �ڋq���_�o�^ = �z����_�o�^_sub()
    End Select
    
End Function

'************************************************************************
'�@  �\ :�ڋq�}�X�^��o�^����B
'************************************************************************
Private Function �ڋq���_�o�^_sub() As Boolean
    
    Dim �ڋqID As String
    Dim �ڋq�}�X�^ As type�ڋq�}�X�^
    Dim row As Integer
    Dim �Z�� As String
        
    �ڋq���_�o�^_sub = False
    
    On Error GoTo err
    
    If va�ڋq���X�g.MaxRows < 1 Then
        Exit Function
    End If
    
   'If txt�ڋq��.Text = "" Then
   '    Call MsgBox("�ڋq���������͂ł�", vbOKOnly, "�ڋq�Ǘ�")
   '    �ڋq���_�o�^_sub = True
   '    Exit Function
   'End If
    
    
    'MsgBox txt�ڋqID.Text, vbOKOnly, "XXXXX"
    'Debug.Print txt�ڋqID.Text
    
    With �ڋq�}�X�^
        
        .�ڋqID = txt�ڋqID.Text
        .�ڋq�� = txt�ڋq��.Text
        .�t���K�i = txt�t���K�i.Text
        .�� = txt�X�֔ԍ�.Text
        .�Z��1 = txt�Z��_��i.Text
        .�Z��2 = txt�Z��_���i.Text
        .�Z��3 = txt�Z��_���i.Text
        .�d�b�ԍ� = txt�d�b�ԍ�.Text
        .���[�� = txt���[��.Text
        .�y�V���[�� = txt�y�V���[��.Text
        .���[�����M = chk���[�����M.Value
        .�A�[�f���N���u = cmb�A�[�f���N���u.Text
        
        If txt�����.Text = "____/__/__" Then
            .����� = ""
        Else
            .����� = txt�����.Text
        End If
        
        If txt�މ��.Text = "____/__/__" Then
            .�މ�� = ""
        Else
            .�މ�� = txt�މ��.Text
        End If
        
        If txt�a����.Text = "____/__/__" Then
            .�a���� = ""
        Else
            .�a���� = txt�a����.Text
        End If
        
        .���� = IIf(opt�j��.Value = True, "1", "2")
        .���l = txt���l.Text
        .����1 = chk����1.Value
        .����2 = chk����2.Value
        .����3 = chk����3.Value
        .����4 = chk����4.Value
        .����5 = chk����5.Value
        .�폜 = "0"
                
        If .�ڋq�� = "" _
            And .�t���K�i = "" _
            And .�� = "" _
            And .�Z��1 = "" _
            And .�Z��2 = "" _
            And .�Z��3 = "" _
            And .�d�b�ԍ� = "" _
            And .���[�� = "" _
            And .�y�V���[�� = "" _
            And .�A�[�f���N���u = "" _
            And (.����� = "____/__/__" Or .����� = "") _
            And (.�މ�� = "____/__/__" Or .�މ�� = "") _
            And (.�a���� = "____/__/__" Or .�a���� = "") _
            And .���l = "" Then
            Exit Function
        End If
        
        If .�ڋqID <> "" Then
            ' �ڋqID���̔ԍς݂̏ꍇ�́A�ڋq�f�[�^���X�V����
            If �ڋq�}�X�^�X�V(�ڋq�}�X�^) = False Then
                If MsgBox("���̒[���ōX�V����Ă��邽�߁A�X�V�ł��܂���B" + Chr$(13) + Chr$(10) + "�����[�h���܂����H", vbYesNo, "�ڋq�Ǘ�") = vbYes Then
                    Call cmd���o�׈ꗗ_Click
                End If
                Exit Function
            End If
        Else
            ' �ڋqID�����̔Ԃ̏ꍇ�́A�ڋq�f�[�^��V�K�ɓo�^����
            .�ڋqID = �ڋq�}�X�^�o�^(�ڋq�}�X�^)
            txt�ڋqID.Text = .�ڋqID
        End If
        
        row = G_�ڋq���X�g_ROW
        'Call SpreadSetVal(va�ڋq���X�g, row, COL_�`�F�b�N, 0)
        Call SpreadSetVal(va�ڋq���X�g, row, COL_�ڋqID, .�ڋqID)
        Call SpreadSetVal(va�ڋq���X�g, row, COL_�ڋq��, .�ڋq��)
        Call SpreadSetVal(va�ڋq���X�g, row, COL_�t���K�i, .�t���K�i)
        Call SpreadSetVal(va�ڋq���X�g, row, COL_��, .��)
        
        �Z�� = .�Z��1 + .�Z��2 + .�Z��3
        Call SpreadSetVal(va�ڋq���X�g, row, COL_�Z��1, �Z��)
        'Call SpreadSetVal(va�ڋq���X�g, row, COL_�Z��2, .�Z��2)
        'Call SpreadSetVal(va�ڋq���X�g, row, COL_�Z��3, .�Z��3)
        Call SpreadSetVal(va�ڋq���X�g, row, COL_�d�b�ԍ�, .�d�b�ԍ�)
        Call SpreadSetVal(va�ڋq���X�g, row, COL_���[��, .���[��)
        Call SpreadSetVal(va�ڋq���X�g, row, COL_�A�[�f���N���u, .�A�[�f���N���u)
        Call SpreadSetVal(va�ڋq���X�g, row, COL_�����, .�����)
        Call SpreadSetVal(va�ڋq���X�g, row, COL_���l, .���l)
        Call SpreadSetVal(va�ڋq���X�g, row, COL_�y�V���[��, .�y�V���[��)
        txt���ӊ��N.Caption = .���l
            
    End With
              
    Exit Function
    
err:
    Call MsgBox("DB�X�V�G���[�ɂ��ċN�����ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
End Function

'************************************************************************
'�@  �\ :�z�������o�^����B
'************************************************************************
Private Function �z����_�o�^_sub() As Boolean
    
    Dim �ڋqID As String
    Dim �ڋq�}�X�^ As type�ڋq�}�X�^
    Dim row As Integer
    
    On Error GoTo err
    
    �z����_�o�^_sub = False
    
    If va�ڋq���X�g.MaxRows < 1 Then
        �z����_�o�^_sub = False
        Exit Function
    End If
    
    If txt�ڋqID.Text = "" Then
        Call MsgBox("�ڋq��񂪖����͂ł�", vbOKOnly, "�ڋq�Ǘ�")
        �z����_�o�^_sub = True
        Exit Function
    End If
    
'   If txt�ڋq��.Text = "" Then
'       Call MsgBox("�ڋq���������͂ł�", vbOKOnly, "�ڋq�Ǘ�")
'       �z����_�o�^_sub = True
'       Exit Function
'   End If
    
    With �ڋq�}�X�^
        
        .�ڋqID = txt�ڋqID.Text
        .�ڋq�� = txt�ڋq��.Text
        .�t���K�i = txt�t���K�i.Text
        .�� = txt�X�֔ԍ�.Text
        .�Z��1 = txt�Z��_��i.Text
        .�Z��2 = txt�Z��_���i.Text
        .�Z��3 = txt�Z��_���i.Text
        .�d�b�ԍ� = txt�d�b�ԍ�.Text
        .���[�� = txt���[��.Text
        .���� = IIf(opt�j��.Value = True, "1", "2")
        .���l = txt���l.Text
        .�폜 = "0"
        
        If .�ڋqID <> "" Then
            ' �ڋqID���̔ԍς݂̏ꍇ�́A�ڋq�f�[�^���X�V����
            If �z����X�V(�ڋq�}�X�^) = False Then
                If MsgBox("���̒[���ōX�V����Ă��邽�߁A�X�V�ł��܂���B" + Chr$(13) + Chr$(10) + "�����[�h���܂����H", vbYesNo, "�ڋq�Ǘ�") = vbYes Then
                    Call cmd���o�׈ꗗ_Click
                End If
                Exit Function
            End If
        Else
            ' �ڋqID�����̔Ԃ̏ꍇ�́A�ڋq�f�[�^��V�K�ɓo�^����
            Call �z����o�^(�ڋq�}�X�^)
        End If
        
        row = G_�ڋq���X�g_ROW
        Call SpreadSetVal(va�ڋq���X�g, row, COL_���͂��於, .�ڋq��)
        Call SpreadSetVal(va�ڋq���X�g, row, COL_���͂��惁�[��, .���[��)
    
    End With
    
    Exit Function
err:
    Call MsgBox("DB�X�V�G���[�ɂ��ċN�����ĉ������B", vbOKOnly, "�ڋq�Ǘ�")

End Function

'************************************************************************
'�@  �\�@�������X�V����B
'************************************************************************
Private Function ����_�X�V() As Boolean
    
    Dim i               As Long
    Dim row             As Long
    Dim ���㖾��RS      As New ADODB.Recordset
    Dim ����ID          As String
    Dim �ڋqID          As String
    Dim �ڋq��          As String
    Dim �ݐϖ{��        As Integer
    Dim �z�B��]����    As String
    Dim �����}�K���M�\���  As Date
    On Error GoTo err
    
    Dim ���㖾�� As type���㖾��
    
    'MsgBox txt����ID.Text, vbOKOnly, "XXXXX"
    'Debug.Print txt����ID.Text
    
    ����_�X�V = True
    
    If va�ڋq���X�g.MaxRows < 1 Then
        ����_�X�V = False
        Exit Function
    End If
    
    If va�������X�g.MaxRows < 1 Then
        ����_�X�V = False
        Call cmd�ǉ�2_Click
        Exit Function
    End If
    
    ����ID = txt����ID.Text
    
    If ����ID = "" Then
        �ڋqID = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�ڋqID)
        �ڋq�� = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�ڋq��)
        '�ڋqID = txt�ڋqID.Text
        '�ڋq�� = txt�ڋq��.Text
    
        If �ڋqID = "" Then
            Call MsgBox("�ڋq��񂪖����͂ł�", vbOKOnly, "�ڋq�Ǘ�")
            ����_�X�V = False
            Exit Function
        End If
    End If
    
    If cmb�X�e�[�^�X.Text = "�o�׊���" Then
    
        If txt�󒍓�.Text >= "2012/02/15" Then
        
            If txt�o�ד�.Text = "____/__/__" Then
                Call MsgBox("�o�ד��������͂ł�", vbOKOnly, "�ڋq�Ǘ�")
                ����_�X�V = False
                Exit Function
            End If
        
            If txt�����ԍ�.Text = "" Then
                Call MsgBox("�����ԍ��������͂ł�", vbOKOnly, "�ڋq�Ǘ�")
                ����_�X�V = False
                Exit Function
            End If
            
            Select Case cmb������.Text
                Case "�y�V"
                    If Mid(txt�����ԍ�.Text, 7, 1) = "-" And Mid(txt�����ԍ�.Text, 16, 1) = "-" Then
                    Else
                        Call MsgBox("�����ԍ��̌`�������ł�", vbOKOnly, "�ڋq�Ǘ�")
                        ����_�X�V = False
                        Exit Function
                    End If
                Case "Yahoo"
                    If Mid(txt�����ԍ�.Text, 1, 6) = "adele-" Or IsNumeric(txt�����ԍ�.Text) = True Then
                    Else
                        Call MsgBox("�����ԍ��̌`�������ł�", vbOKOnly, "�ڋq�Ǘ�")
                        ����_�X�V = False
                        Exit Function
                    End If
                Case "�����g���b�N�X"
                    If Mid(txt�����ԍ�.Text, 1, 1) = "R" Then
                    Else
                        Call MsgBox("�����ԍ��̌`�������ł�", vbOKOnly, "�ڋq�Ǘ�")
                        ����_�X�V = False
                        Exit Function
                    End If
'                Case "������̂��l�b�g"
'                    If Mid(txt�����ԍ�.Text, 1, 5) = "OCNK-" Then
'                    Else
'                        Call MsgBox("�����ԍ��̌`�������ł�", vbOKOnly, "�ڋq�Ǘ�")
'                        ����_�X�V = False
'                        Exit Function
'                    End If
                Case "�A�}�]��"
                    If Mid(txt�����ԍ�.Text, 4, 1) = "-" And Mid(txt�����ԍ�.Text, 12, 1) = "-" Then
                    Else
                        Call MsgBox("�����ԍ��̌`�������ł�", vbOKOnly, "�ڋq�Ǘ�")
                        ����_�X�V = False
                        Exit Function
                    End If
                Case "�R�}�`"
                    If Mid(txt�����ԍ�.Text, 1, 8) = "KOMACHI-" Then
                    Else
                        Call MsgBox("�����ԍ��̌`�������ł�", vbOKOnly, "�ڋq�Ǘ�")
                        ����_�X�V = False
                        Exit Function
                    End If
                Case Else
                    If Mid(txt�����ԍ�.Text, 1, 4) = "ETC-" Then
                    Else
                        Call MsgBox("�����ԍ��̌`�������ł�", vbOKOnly, "�ڋq�Ǘ�")
                        ����_�X�V = False
                        Exit Function
                    End If
            End Select
            
            If CLng(txt���v���z.Text) < 0 Then
                Call MsgBox("���v���z���}�C�i�X�ɂȂ�Ȃ��悤�ɓ��͂��ĉ�����", vbOKOnly, "�ڋq�Ǘ�")
                ����_�X�V = False
                Exit Function
            End If
            
            If cmb����.Text = "�����" Then
                If �A�[�f������(cmb���i��.Text) = 1 Or �A�[�f������(cmb���i��.Text) = 9 Then
                Else
                    Call MsgBox("���傪����Ă��܂�", vbOKOnly, "�ڋq�Ǘ�")
                    ����_�X�V = False
                    Exit Function
                End If
            End If
            
            If cmb����.Text = "���̑�" Then
                If �A�[�f������(cmb���i��.Text) = 1 Or �A�[�f������(cmb���i��.Text) = 9 Then
                    Call MsgBox("���傪����Ă��܂�", vbOKOnly, "�ڋq�Ǘ�")
                    ����_�X�V = False
                    Exit Function
                End If
            End If
            
            If cmb����.Text = "��ײ�" And CLng(txt���v���z.Text) <> 0 Then
'               If txt�d�����z.Value = 0 Or txt�ב��^��.Value = 0 Then
                If txt�d�����z.Value = 0 Then
                    Call MsgBox("�d�����z�^�ב��^������͂��ĉ�����", vbOKOnly, "�ڋq�Ǘ�")
                    ����_�X�V = False
                    Exit Function
                End If
            End If
            
            If cmb�������@.Text = "��s�U��" Then
                If cmb��s.Text = "" Then
                    Call MsgBox("��s�U���̏ꍇ�A��s������͂��ĉ�����", vbOKOnly, "�ڋq�Ǘ�")
                    ����_�X�V = False
                    Exit Function
                End If
            End If
            
            If cmb�������@.Text = "���i���" Then
                If �A�[�f������(cmb���i��.Text) = 1 Then
                    If cmb��z�Ǝ�.Text = "����}��" Or cmb��z�Ǝ�.Text = "�䂤�p�b�N" Then
                    Else
                        Call MsgBox("���i������̏ꍇ�A����}�� or �䂤�p�b�N����͂��ĉ�����", vbOKOnly, "�ڋq�Ǘ�")
                        ����_�X�V = False
                        Exit Function
                    End If
                End If
            End If
        End If
    End If

    
#If 0 Then
    If cmb������.Text = "" Then
        Call MsgBox("��������I�����ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
        ����_�X�V = False
        Exit Function
    End If
    
    If cmb���i��.Text = "" Then
        Call MsgBox("���i��I�����ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
        ����_�X�V = False
        Exit Function
    End If
#End If
        
    With ���㖾��
        
        If txt�󒍓�.Text = "____/__/__" Then
            .�󒍓� = ""
        Else
            .�󒍓� = txt�󒍓�.Text
        End If
        .�X�e�[�^�X = cmb�X�e�[�^�X.Text
        .���i�� = cmb���i��.Text
        .���� = cmb����.Text
        .�������@ = cmb�������@.Text
        .��s = cmb��s.Text
        .�z�B��]���� = txt�z�B����.Text
        .�z�B��]����2 = txt�z�B����2.Text
        .�d�����z = txt�d�����z.Value
        .�P�� = txt�P��.Value
        .���� = txt����.Value
        .�����敪 = "\"
        .���� = txt����.Value
        .���z = (txt�P��.Value + txt����.Value) * txt����.Value
        .����� = 0
        .���� = txt����.Value
        .�ב��^�� = txt�ב��^��.Value
        .�ԋ� = txt�ԋ�.Value
        .���̑��萔�� = txt���̑��萔��.Value
        .���v���z = CLng(txt���v���z.Text)
        
        If txt������.Text = "____/__/__" Then
            .������ = ""
        Else
            .������ = txt������.Text
        End If
        
        If txt�o�ד�.Text = "____/__/__" Then
            .�o�ד� = ""
        Else
            .�o�ד� = txt�o�ד�.Text
        End If
        
        .���ד� = ""
        .��z�Ǝ� = cmb��z�Ǝ�.Text
        .������ = cmb������.Text
        .Yahoo�����ԍ� = Trim(txt�����ԍ�.Text)
        .�Q�ƌ� = ""
        .�L�[���[�h = ""
        .���̓|�C���g = ""
        .���i�R�[�h = ""
        .���C�����e�B�[ = 0
        .���t���� = ""
        .�ԕi�Ώ� = ""
        .�x���ԍ� = txt�x���ԍ�.Text
        .�⍇�ԍ� = txt�⍇�ԍ�.Text
        .���l1 = txt���l2.Text
        .���l2 = ""
        .���l3 = ""
        .�R�����C�tNO = Trim(txt�R�����C�t.Text)
        
        If txt�o�ח\���.Text = "____/__/__" Then
            .�o�ח\��� = ""
        Else
            .�o�ח\��� = txt�o�ח\���.Text
        End If

        .����URL = txt����URL.Text

        If ����ID <> "" Then .����ID = CLng(����ID) Else .����ID = -1
        .�ڋqID = �ڋqID
        .�ڋq�� = �ڋq��
        .���[�����M = txt���[�����M.Text
        .���㒊�o = "0"
        .�폜 = "0"

        If .����ID <> -1 Then
            ' ����ID���̔ԍς݂̏ꍇ�́A�����f�[�^���X�V����
            If ���㖾�׍X�V(���㖾��) = False Then
                If MsgBox("���̒[���ōX�V����Ă��邽�߁A�X�V�ł��܂���B" + Chr$(13) + Chr$(10) + "�����[�h���܂����H", vbYesNo, "�ڋq�Ǘ�") = vbYes Then
                    Call cmd���o�׈ꗗ_Click
                End If
                Exit Function
            End If
        Else
            ' ����ID�����̔Ԃ̏ꍇ�́A�����f�[�^��V�K�ɓo�^����
            ����ID = ���㖾�דo�^(���㖾��)
        End If
        
        txt����ID.Text = CStr(����ID)
        
        If va�������X�g.MaxRows < 1 Then
            va�������X�g.MaxRows = 1
            G_�������X�g_ROW = 1
        End If
        
        row = G_�������X�g_ROW
        Call SpreadSetVal(va�������X�g, row, COL_�󒍓�, .�󒍓�)
        Call SpreadSetVal(va�������X�g, row, COL_�X�e�[�^�X, .�X�e�[�^�X)
        Call SpreadSetVal(va�������X�g, row, COL_���i��, .���i��)
        Call SpreadSetVal(va�������X�g, row, COL_�������@, .�������@)
        �z�B��]���� = .�z�B��]���� + " " + .�z�B��]����2
        Call SpreadSetVal(va�������X�g, row, COL_�z�B��]����, �z�B��]����)
        Call SpreadSetVal(va�������X�g, row, COL_�P��, .�P��)
        Call SpreadSetVal(va�������X�g, row, COL_����, .����)
        Call SpreadSetVal(va�������X�g, row, COL_����, .����)
        Call SpreadSetVal(va�������X�g, row, COL_���z, .���z)
        Call SpreadSetVal(va�������X�g, row, COL_����, .����)
        Call SpreadSetVal(va�������X�g, row, COL_�ԋ�, .�ԋ�)
        Call SpreadSetVal(va�������X�g, row, COL_���̑��萔��, .���̑��萔��)
        Call SpreadSetVal(va�������X�g, row, COL_���v���z, .���v���z)
        Call SpreadSetVal(va�������X�g, row, COL_������, .������)
        Call SpreadSetVal(va�������X�g, row, COL_�o�ד�, .�o�ד�)
        Call SpreadSetVal(va�������X�g, row, COL_���ד�, .���ד�)
        Call SpreadSetVal(va�������X�g, row, COL_��z�Ǝ�, .��z�Ǝ�)
        Call SpreadSetVal(va�������X�g, row, COL_������, .������)
        Call SpreadSetVal(va�������X�g, row, COL_Yahoo�����ԍ�, .Yahoo�����ԍ�)
        Call SpreadSetVal(va�������X�g, row, COL_�Q�ƌ�, .�Q�ƌ�)
        Call SpreadSetVal(va�������X�g, row, COL_�L�[���[�h, .�L�[���[�h)
        Call SpreadSetVal(va�������X�g, row, COL_���̓|�C���g, .���̓|�C���g)
        Call SpreadSetVal(va�������X�g, row, COL_���i�R�[�h, .���i�R�[�h)
        Call SpreadSetVal(va�������X�g, row, COL_���C�����e�B�[, .���C�����e�B�[)
        Call SpreadSetVal(va�������X�g, row, COL_���t����, .���t����)
        Call SpreadSetVal(va�������X�g, row, COL_�ԕi�Ώ�, .�ԕi�Ώ�)
        Call SpreadSetVal(va�������X�g, row, COL_�x���ԍ�, .�x���ԍ�)
        Call SpreadSetVal(va�������X�g, row, COL_�⍇�ԍ�, .�⍇�ԍ�)
        Call SpreadSetVal(va�������X�g, row, COL_���l1, .���l1)
        Call SpreadSetVal(va�������X�g, row, COL_���l2, .���l2)
        Call SpreadSetVal(va�������X�g, row, COL_���l3, .���l3)
        Call SpreadSetVal(va�������X�g, row, COL_����ID, ����ID)
        Call SpreadSetVal(va�������X�g, row, COL_���[�����M, .���[�����M)
        Call SpreadSetVal(va�������X�g, row, COL_�����敪, "�~")
        Call SpreadSetVal(va�������X�g, row, COL_�o�ח\���, .�o�ח\���)
        Call SpreadSetVal(va�������X�g, row, COL_����URL, .����URL)
        
    End With
    
    ����_�X�V = False
    
    txt�ݐϐ�.Text = �ݐϐ��v�Z()
    
    If cmb�X�e�[�^�X.Text = "�o�׊���" Then
        �ڋqID = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�ڋqID)
        '�����}�K���M�\��� = Format(DateAdd("d", 30, txt�o�ד�.Text), "yyyy/mm/dd")
        'Call �����}�K���sNO�X�V(�ڋqID, 0, "'" + CStr(�����}�K���M�\���) + "'")
    End If
    
    Exit Function
    
err:
    Call MsgBox("DB�X�V�G���[�ɂ��ċN�����ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
    
End Function

'************************************************************************
'�@  �\ :�A�[�f���̗ݐύw�������擾����
'************************************************************************
Private Function �ݐϐ��v�Z() As Long
    
    Dim i           As Long
    Dim �ݐϖ{��    As Long
    Dim �X�e�[�^�X  As String
    Dim ���i��      As String
    Dim ����        As String
    
    �ݐϖ{�� = 0
    �ݐϐ��v�Z = 0
    
    With va�������X�g
        For i = 1 To .MaxRows
            �X�e�[�^�X = SpreadGetVal(va�������X�g, i, COL_�X�e�[�^�X)
            ���i�� = SpreadGetVal(va�������X�g, i, COL_���i��)
            ���� = SpreadGetVal(va�������X�g, i, COL_����)
            
            If �X�e�[�^�X <> "�L�����Z��" And �X�e�[�^�X <> "�ۗ�" And �X�e�[�^�X <> "��������" Then
                If ���i�� = "�A�[�f��" Or ���i�� = "�X�[�p�[�A�[�f��" Or ���i�� = "�A�[�f��(�Z�[��)" Or ���i�� = "�A�[�f���{�V�����v�[" Or ���i�� = "�A�[�f���{�V�����v�[�����i" Then
                    �ݐϖ{�� = �ݐϖ{�� + IIf(IsNumeric(����), CInt(����), 0)
                End If
                
                If ���i�� = "�A�[�f��2�{�Z�b�g" Then
                    �ݐϖ{�� = �ݐϖ{�� + IIf(IsNumeric(����), CInt(����), 0) * 2
                End If
            End If
        Next i
    End With
    
    �ݐϐ��v�Z = �ݐϖ{��

End Function

'************************************************************************
'�@  �\ :�A�[�f���̗ݐύw�������擾����
'************************************************************************
Private Function �ݐϐ��v�Z2(ByVal ID As Long) As Long
    
    Dim i           As Long
    Dim �ݐϖ{��    As Long
    Dim �X�e�[�^�X  As String
    Dim ���i��      As String
    Dim ����        As String
    Dim ����ID      As Long
    
    �ݐϖ{�� = 0
    �ݐϐ��v�Z2 = 0
    
    With va�������X�g
        For i = 1 To .MaxRows
            ����ID = SpreadGetVal2(va�������X�g, i, COL_����ID)
            �X�e�[�^�X = SpreadGetVal(va�������X�g, i, COL_�X�e�[�^�X)
            ���i�� = SpreadGetVal(va�������X�g, i, COL_���i��)
            ���� = SpreadGetVal(va�������X�g, i, COL_����)
            
            If ����ID <> ID Then
                If �X�e�[�^�X <> "�L�����Z��" And �X�e�[�^�X <> "�ۗ�" And �X�e�[�^�X <> "��������" Then
                If ���i�� = "�A�[�f��" Or ���i�� = "�X�[�p�[�A�[�f��" Or ���i�� = "�A�[�f��(�Z�[��)" Or ���i�� = "�A�[�f���{�V�����v�[" Or ���i�� = "�A�[�f���{�V�����v�[�����i" Then
                        �ݐϖ{�� = �ݐϖ{�� + IIf(IsNumeric(����), CInt(����), 0)
                    End If
                    
                    If ���i�� = "�A�[�f��2�{�Z�b�g" Then
                        �ݐϖ{�� = �ݐϖ{�� + IIf(IsNumeric(����), CInt(����), 0) * 2
                    End If
                End If
            End If
        Next i
    End With
    
    �ݐϐ��v�Z2 = �ݐϖ{��

End Function

'************************************************************************
'�@  �\ :�X�v���b�h�V�[�g�Ƀf�[�^��ݒ肷��
'************************************************************************
Public Sub SpreadSetVal(ByVal Spread As vaSpread, ByVal lngRow As Long, ByVal lngCol As Long, ByVal strText As String)
    With Spread
        .row = lngRow
        .Col = lngCol
        .Text = strText
    End With
End Sub

'************************************************************************
'�@  �\ :�X�v���b�h�V�[�g����f�[�^���擾����
'************************************************************************
Public Function SpreadGetVal(ByVal Spread As vaSpread, ByVal lngRow As Long, ByVal lngCol As Long) As String
    With Spread
        .row = lngRow
        .Col = lngCol
        SpreadGetVal = Trim(.Text)
    End With
End Function

'************************************************************************
'�@  �\ :�X�v���b�h�V�[�g����f�[�^���擾����
'************************************************************************
Public Function SpreadGetVal2(ByVal Spread As vaSpread, ByVal lngRow As Long, ByVal lngCol As Long) As Long
    With Spread
        .row = lngRow
        .Col = lngCol
        If Trim(.Text) <> "" Then
            If IsNumeric(Trim(.Text)) Then
                SpreadGetVal2 = CLng(Trim(.Text))
            End If
        Else
            SpreadGetVal2 = 0
        End If
    End With
End Function

'************************************************************************
'�@  �\ :�X�v���b�h�V�[�g�̃Z���ʒu��ݒ肷��
'************************************************************************
Public Sub SpreadSetFocus(ByVal Spread As vaSpread, ByVal lngRow As Long, ByVal lngCol As Long)
        
    With Spread
        .Col = lngCol
        .row = lngRow
        .Position = 6
        .Action = 1
        
        .Col = lngCol
        .row = lngRow
        .Action = 0
        .SetFocus
    End With
    
End Sub

'************************************************************************
'�@  �\ :�`�F�b�N����Ă��錏�����擾����
'************************************************************************
Function �`�F�b�N�����擾() As Integer
    
    Dim i As Integer
    Dim cnt As Integer
    
    cnt = 0
    
    For i = 1 To va�������X�g.MaxRows
        If SpreadGetVal(va�������X�g, i, COL_�`�F�b�N) = "1" Then
            cnt = cnt + 1
        End If
    Next i
    
    �`�F�b�N�����擾 = cnt
    
End Function

'************************************************************************
'�@  �\ :CSV�o�͏������s��
'************************************************************************
Private Sub cmdCSV�o��_Click()

    Dim �ڋqID          As String
    Dim �ڋq��          As String
    Dim ��              As String
    Dim �Z��1           As String
    Dim �Z��2           As String
    Dim �Z��3           As String
    Dim �d�b�ԍ�        As String
    Dim intFileNo       As Integer

    Dim CSV���oRS As New ADODB.Recordset
    Dim �z����RS As New ADODB.Recordset
    
    If MsgBox("�Z���^�b�r�u���o�͂��Ă���낵���ł����H", vbYesNo, "�ڋq�Ǘ�") = vbNo Then
        Exit Sub
    End If

    MousePointer = vbHourglass
    
    intFileNo = FreeFile()
    Open "C:\�ڋq�Ǘ�\����}��.csv" For Output As #intFileNo

    Call CSV���o�͌ڋq�}�X�^�Ǎ�(CSV���oRS)

    If CSV���oRS.EOF Then
        CSV���oRS.Close
        Close #intFileNo
        MousePointer = vbNormal
        Call MsgBox("�V�K�Z���^�f�[�^�͑��݂��܂���", vbOKOnly, "�ڋq�Ǘ�")
        Exit Sub
    End If
    
    With CSV���oRS
        Do Until .EOF
        
        �ڋq�� = !�ڋq��
        �� = ![��]
        �Z��1 = !�Z��1
        �Z��2 = !�Z��2
        �Z��3 = IIf(IsNull(!�Z��3), "", !�Z��3)
        �d�b�ԍ� = !�d�b�ԍ�
        
        Call �z����1���Ǎ�(�z����RS, !�ڋqID)
        
        If Not �z����RS.EOF Then
            If �z����RS!�ڋq�� <> "" Then
                �ڋq�� = �z����RS!�ڋq��
                �� = �z����RS![��]
                �Z��1 = �z����RS!�Z��1
                �Z��2 = �z����RS!�Z��2
                �Z��3 = �z����RS!�Z��3
                �d�b�ԍ� = �z����RS!�d�b�ԍ�
            End If
        End If
        
        �z����RS.Close
        
        Print #intFileNo, CStr(CLng(!�ڋqID)) & "," _
                            & �Z��1 & "," _
                            & �Z��2 & "," _
                            & �Z��3 & "," _
                            & �ڋq�� & "," _
                            & "," _
                            & �d�b�ԍ� & "," _
                            & �� & ",,,,,,,,,,,,,,,000,,,00,,,,,,10,00,,,,,,"
                            

        Call CSV�o�̓t���O�X�V(!�ڋqID)
        CSV���oRS.MoveNext
        Loop
        .Close
    End With
    
    Close #intFileNo
    MousePointer = vbNormal
    Call MsgBox("�uC:\�ڋq�Ǘ�\����}��.csv�v�ɁA�Z���^�f�[�^���o�͂��܂���", vbOKOnly, "�ڋq�Ǘ�")
    
End Sub

'************************************************************************
'�@  �\ :����f�[�^���b�r�u�o�͂���
'************************************************************************
Private Sub cmd����_Click()
    
    Dim ��v            As type��v
    Dim ���o��          As String
    Dim ADF018          As New ADF018

    Dim ���㒊�oRS As New ADODB.Recordset
    
    If MsgBox("����b�r�u���o�͂��Ă���낵���ł����H", vbYesNo, "�ڋq�Ǘ�") = vbNo Then
        Exit Sub
    End If
    
    If MsgBox("�{���ɍ쐬���Ă���낵���ł����H", vbYesNo, "�ڋq�Ǘ�") = vbNo Then
        Exit Sub
    End If
    
    'Call ADF018.Show(1)

    '���o�� = ADF018.���o���擾()

    MousePointer = vbHourglass
    
    Call ����f�[�^�폜
    
    'Call ���o�͔���f�[�^�Ǎ�(���㒊�oRS, ���o��)
    Call ���o�͔���f�[�^�Ǎ�(���㒊�oRS)

    If ���㒊�oRS.EOF Then
        ���㒊�oRS.Close
        MousePointer = vbNormal
        Call MsgBox("�V�K����b�r�u�f�[�^�͑��݂��܂���", vbOKOnly, "�ڋq�Ǘ�")
        Exit Sub
    End If
    
    
    ��v.���ʃt���O = "11"
    ��v.�`�[NO = 0                         '�`�[�ԍ��擾()                    ' """"""
'    ��v.���Z = """"""
    ��v.������� = ""
    
    ��v.�^�C�v = "3"
    ��v.������ = "�U�`"
    
    ' ���㉼����
    Call ���㏈��(��v, ���㒊�oRS, False)
    
    ���㒊�oRS.Close
    
    ' �ؕ����z�Ƒݕ����z���`�F�b�N����
    If ����f�[�^�`�F�b�N() = False Then
        MousePointer = vbNormal
        Call MsgBox("�ؕ����z�Ƒݕ����z�������܂���", vbOKOnly, "�ڋq�Ǘ�")
        Exit Sub
    End If
    
    Call ����f�[�^�폜
    
    'Call ���o�͔���f�[�^�Ǎ�(���㒊�oRS, ���o��)
    Call ���o�͔���f�[�^�Ǎ�(���㒊�oRS)
    
    
    ' ����{����
    Call ���㏈��(��v, ���㒊�oRS, True)
    
    ���㒊�oRS.Close
    
    Call ����f�[�^�폜2
    
    Call ����f�[�^�R�s�[
    
    Call ���ʃt���O�ݒ�
    
    If G_�X�ܖ� = "�g���j�e�B�[�y�V�s��X" Then
        Call �y�V_�����X�e�[�^�X�ύX
    Else
        Call Yahoo_�����X�e�[�^�X�ύX
    End If
    
    Call ��vCSV�o��

    MousePointer = vbNormal

    If G_�X�ܖ� = "�g���j�e�B�[�y�V�s��X" Then
        Call MsgBox("�uC:\�ڋq�Ǘ�\�y�V_����.csv�v�ɁA�f�[�^���o�͂��܂���", vbOKOnly, "�ڋq�Ǘ�")
    Else
        Call MsgBox("�uC:\�ڋq�Ǘ�\Yahoo_����.csv�v�ɁA�f�[�^���o�͂��܂���", vbOKOnly, "�ڋq�Ǘ�")
    End If
    
End Sub

'************************************************************************
'�@  �\ :���㏈��
'************************************************************************
Private Sub ���㏈��(ByRef ��v As type��v, ByVal ���㒊�oRS As ADODB.Recordset, ByVal �{�ԋ敪 As Boolean)
    
    Dim �敪            As Integer
    
    With ���㒊�oRS
    
        Do Until .EOF
            ' �R�����C�t�̏ꍇ�A�P���Â������΂炩����B
            ' �������Ȃ��ƁA�������������������ꍇ�A�d����ƁA�ב��^�����ŏ��Ɋ���Ă��܂�
            '
            If !���� <> "��ײ�" Then
                ��v.�����ԍ� = !Yahoo�����ԍ�
            Else
                ��v.�����ԍ� = !Yahoo�����ԍ� & "#" & !����ID
            End If
            
            ��v.����ID = !����ID
            ��v.������� = !�o�ד�
            
            If !���� = "�����" Then
                �敪 = 1
            ElseIf !���� = "��ײ�" Then
                �敪 = 2
            ElseIf !���� = "���̑�" Then
                �敪 = 3
            Else
                If �A�[�f������(!���i��) = 1 Then
                    �敪 = 1
                Else
                    �敪 = 2
                End If
            End If
            
            '
            ' �A�[�f������o��
            '
            If �敪 = 1 Then
                If !���v���z > 0 Then
                    If !�������@ = "��s�U��" Then
                        Call ����_�ʏ�o��(��v, ���㒊�oRS)
                    ElseIf !�������@ = "�y�V�o���N����" Then
                        Call �y�V�o���N_�ʏ�o��(��v, ���㒊�oRS)
                    Else
                        Call ���|��_�ʏ�o��(��v, ���㒊�oRS)
                    End If
                Else
                    If !���v���z = 0 And !���̑��萔�� < 0 Then
                        Call ���|��_�|�C���g�o��(��v, ���㒊�oRS)
                    End If
                End If
                
            '
            ' �R�����C�t����o��
            '
            ElseIf �敪 = 2 Then
                If !���v���z > 0 Then
                    If !�������@ = "��s�U��" Then
                        Call ����_�R�����C�t�o��(��v, ���㒊�oRS)
                    ElseIf !�������@ = "�y�V�o���N����" Then
                        Call �y�V�o���N_�R�����C�t�o��(��v, ���㒊�oRS)
                    Else
                        Call ���|��_�R�����C�t�o��(��v, ���㒊�oRS)
                    End If
                Else
                    If !���v���z = 0 And !���̑��萔�� < 0 Then
                        Call ���|��_�|�C���g_�R�����C�t�o��(��v, ���㒊�oRS)
                    End If
                End If
                
                If !���z > 0 Then
                    Call ���|��_�o��(��v, ���㒊�oRS)
                    
                    Call �ב��^��_�o��(��v, ���㒊�oRS)
                End If
            '
            ' ���̑��o��
            '
            Else
                If !���v���z > 0 Then
                    If !�������@ = "��s�U��" Then
                        Call ����_���̑��o��(��v, ���㒊�oRS)
                    ElseIf !�������@ = "�y�V�o���N����" Then
                        Call �y�V�o���N_���̑��o��(��v, ���㒊�oRS)
                    Else
                        Call ���|��_���̑��o��(��v, ���㒊�oRS)
                    End If
                Else
                    If !���v���z = 0 And !���̑��萔�� < 0 Then
                        Call ���|��_�|�C���g_���̑��o��(��v, ���㒊�oRS)
                    End If
                End If
                
            End If
            
            If �{�ԋ敪 = True Then
                Call ����o�̓t���O�X�V(!����ID)
            End If
            
            .MoveNext
        Loop
    End With

End Sub

'************************************************************************
'�@  �\ :���|���o�́i�ʏ프��j
'************************************************************************
Private Sub ���|��_�ʏ�o��(ByRef k As type��v, ByVal u As ADODB.Recordset)
    
    k.�ؕ�����Ȗ� = "���|��"
'    If u!������ = "Yahoo" Or u!������ = "�y�V" Or u!������ = "���ЃT�C�g" Or u!������ = "������̂��l�b�g" Or u!������ = "�R�}�`" Then
    If u!������ = "Yahoo" Or u!������ = "�y�V" Or u!������ = "���ЃT�C�g" Or u!������ = "�����g���b�N�X" Or u!������ = "�R�}�`" Then
        k.�ؕ��⏕�Ȗ� = �ؕ��⏕�Ȗڎ擾1(u!�������@, u!��z�Ǝ�)
    Else
        k.�ؕ��⏕�Ȗ� = u!������
    End If
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = "�ΏۊO"
    k.�ؕ����z = u!���v���z - u!����
    k.�ؕ��ŋ��z = 0
    
    k.�ݕ�����Ȗ� = "���㍂"
    If u!������ = "Yahoo" Or u!������ = "�y�V" Or u!������ = "���ЃT�C�g" Then
        k.�ݕ��⏕�Ȗ� = �ݕ��⏕�Ȗڎ擾1(u!���i��)
    Else
        k.�ݕ��⏕�Ȗ� = u!������
    End If
    
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_����
    k.�ݕ����z = u!���v���z + u!���̑��萔�� * -1 - u!����
    k.�ݕ��ŋ��z = ����Ōv�Z(k.�ݕ����z)
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)
    
    If u!���� > 0 Then
        Call �ב��^��_�o��2(k, u)
    End If
    
    If u!�ԋ� < 0 Then
        k.�ؕ�����Ȗ� = "���|��"
        k.�ؕ��⏕�Ȗ� = �ؕ��⏕�Ȗڎ擾1(u!�������@, u!��z�Ǝ�)
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!�ԋ� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = "����"
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = u!�ԋ� * -1
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
    
        Call ��v�o��(k)
    
    End If
    
    If u!���̑��萔�� < 0 Then
        k.�ؕ�����Ȗ� = "���|��"
        k.�ؕ��⏕�Ȗ� = "�|�C���g"
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!���̑��萔�� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = ""
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = 0
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
        
        Call ��v�o��(k)
        
    End If

End Sub

'************************************************************************
'�@  �\ :�����o��
'************************************************************************
Private Sub ����_�ʏ�o��(ByRef k As type��v, ByVal u As ADODB.Recordset)
                    
    k.�ؕ�����Ȗ� = "���ʗa��"
    k.�ؕ��⏕�Ȗ� = u!��s
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = "�ΏۊO"
    k.�ؕ����z = u!���v���z - u!����
    k.�ؕ��ŋ��z = 0
    
    k.�ݕ�����Ȗ� = "���㍂"
    
    If u!������ = "Yahoo" Or u!������ = "�y�V" Or u!������ = "���ЃT�C�g" Then
        k.�ݕ��⏕�Ȗ� = �ݕ��⏕�Ȗڎ擾1(u!���i��)
    Else
        k.�ݕ��⏕�Ȗ� = u!������
    End If
    
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_����
    k.�ݕ����z = u!���v���z + u!���̑��萔�� * -1 - u!����
    k.�ݕ��ŋ��z = ����Ōv�Z(k.�ݕ����z)
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)
    
    If u!���� > 0 Then
        Call �ב��^��_�o��2(k, u)
    End If

    If u!�ԋ� < 0 Then
        k.�ؕ�����Ȗ� = "���ʗa��"
        k.�ؕ��⏕�Ȗ� = u!��s
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!�ԋ� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = "����"
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = u!�ԋ� * -1
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
    
        Call ��v�o��(k)
    
    End If
    
    If u!���̑��萔�� < 0 Then
        k.�ؕ�����Ȗ� = "���|��"
        k.�ؕ��⏕�Ȗ� = "�|�C���g"
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!���̑��萔�� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = ""
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = 0
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
        
        Call ��v�o��(k)
        
    End If

End Sub


'************************************************************************
'�@  �\ :�y�V�o���N�o��
'************************************************************************
Private Sub �y�V�o���N_�ʏ�o��(ByRef k As type��v, ByVal u As ADODB.Recordset)
                    
    k.�ؕ�����Ȗ� = "���ʗa��"
    k.�ؕ��⏕�Ȗ� = "�y�V��s"
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = "�ΏۊO"
    k.�ؕ����z = u!���v���z - u!���� - G_�U�荞�萔��
    k.�ؕ��ŋ��z = 0
    
    k.�ݕ�����Ȗ� = "���㍂"
    
    If u!������ = "Yahoo" Or u!������ = "�y�V" Or u!������ = "���ЃT�C�g" Then
        k.�ݕ��⏕�Ȗ� = �ݕ��⏕�Ȗڎ擾1(u!���i��)
    Else
        k.�ݕ��⏕�Ȗ� = u!������
    End If
    
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_����
    k.�ݕ����z = u!���v���z + u!���̑��萔�� * -1 - u!����
    k.�ݕ��ŋ��z = ����Ōv�Z(k.�ݕ����z)
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)
        
    If u!���� > 0 Then
        Call �ב��^��_�o��2(k, u)
    End If

    If u!�ԋ� < 0 Then
        k.�ؕ�����Ȗ� = "���ʗa��"
        k.�ؕ��⏕�Ȗ� = "�y�V��s"
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!�ԋ� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = "����"
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = u!�ԋ� * -1
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
    
        Call ��v�o��(k)
    
    End If
    
    If u!���̑��萔�� < 0 Then
        k.�ؕ�����Ȗ� = "���|��"
        k.�ؕ��⏕�Ȗ� = "�|�C���g"
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!���̑��萔�� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = ""
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = 0
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
        
        Call ��v�o��(k)
        
    End If

    k.�ؕ�����Ȗ� = "�x���萔��"
    k.�ؕ��⏕�Ȗ� = "�U�荞�萔��"
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = G_�d��
    k.�ؕ����z = G_�U�荞�萔��
    k.�ؕ��ŋ��z = ����Ōv�Z(G_�U�荞�萔��)
    
    k.�ݕ�����Ȗ� = ""
    k.�ݕ��⏕�Ȗ� = ""
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = "�ΏۊO"
    k.�ݕ����z = 0
    k.�ݕ��ŋ��z = 0
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)

End Sub

'************************************************************************
'�@  �\ :���|���o�́i�|�C���g�j
'************************************************************************
Private Sub ���|��_�|�C���g�o��(k As type��v, ByVal u As ADODB.Recordset)
                    
    k.�ؕ�����Ȗ� = "���|��"
    k.�ؕ��⏕�Ȗ� = "�|�C���g"
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = "�ΏۊO"
    k.�ؕ����z = u!���̑��萔�� * -1 - u!����
    k.�ؕ��ŋ��z = 0
    
    k.�ݕ�����Ȗ� = "���㍂"
    
    If u!������ = "Yahoo" Or u!������ = "�y�V" Or u!������ = "���ЃT�C�g" Then
        k.�ݕ��⏕�Ȗ� = �ݕ��⏕�Ȗڎ擾1(u!���i��)
    Else
        k.�ݕ��⏕�Ȗ� = u!������
    End If
    
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_����
    k.�ݕ����z = u!���̑��萔�� * -1 - u!����
    k.�ݕ��ŋ��z = ����Ōv�Z(k.�ݕ����z)
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)
        
    If u!���� > 0 Then
        Call �ב��^��_�o��2(k, u)
    End If

    If u!�ԋ� < 0 Then
        k.�ؕ�����Ȗ� = "���|��"
        k.�ؕ��⏕�Ȗ� = "�|�C���g"
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!�ԋ� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = "����"
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = u!�ԋ� * -1
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
    
        Call ��v�o��(k)
    
    End If

End Sub

'************************************************************************
'�@  �\ :���|���o�́i�R�����C�t����j
'************************************************************************
Private Sub ���|��_�R�����C�t�o��(ByRef k As type��v, ByVal u As ADODB.Recordset)
    
    k.�ؕ�����Ȗ� = "���|��"
    If u!������ = "Yahoo" Or u!������ = "�y�V" Or u!������ = "���ЃT�C�g" Then
        k.�ؕ��⏕�Ȗ� = �ؕ��⏕�Ȗڎ擾2(u!�������@, u!��z�Ǝ�)
    Else
        k.�ؕ��⏕�Ȗ� = u!������
    End If
    
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = "�ΏۊO"
    k.�ؕ����z = u!���v���z '+ u!���̑��萔��
    k.�ؕ��ŋ��z = 0
    
    k.�ݕ�����Ȗ� = "���㍂"
    k.�ݕ��⏕�Ȗ� = "�����炢��"
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_����
    k.�ݕ����z = u!���v���z - u!���� + u!���̑��萔�� * -1
    k.�ݕ��ŋ��z = ����Ōv�Z(k.�ݕ����z)
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)
    
    If u!�ԋ� < 0 Then
        k.�ؕ�����Ȗ� = "���|��"
        k.�ؕ��⏕�Ȗ� = �ؕ��⏕�Ȗڎ擾2(u!�������@, u!��z�Ǝ�)
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!�ԋ� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = "����"
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = u!�ԋ� * -1
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
    
        Call ��v�o��(k)
    
    End If
    
    If u!���̑��萔�� < 0 Then
        k.�ؕ�����Ȗ� = "���|��"
        k.�ؕ��⏕�Ȗ� = "�|�C���g"
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!���̑��萔�� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = ""
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = 0
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
        
        Call ��v�o��(k)
    End If

End Sub

'************************************************************************
'�@  �\ :���|���o�́i�R�����C�t�|�C���g�j
'************************************************************************
Private Sub ���|��_�|�C���g_�R�����C�t�o��(ByRef k As type��v, ByVal u As ADODB.Recordset)
                    
    k.�ؕ�����Ȗ� = "���|��"
    k.�ؕ��⏕�Ȗ� = "�|�C���g"
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = "�ΏۊO"
    k.�ؕ����z = u!���̑��萔�� * -1
    k.�ؕ��ŋ��z = 0
    
    k.�ݕ�����Ȗ� = "���㍂"
    k.�ݕ��⏕�Ȗ� = "�����炢��"
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_����
    k.�ݕ����z = u!���̑��萔�� * -1 - u!����
    k.�ݕ��ŋ��z = ����Ōv�Z(k.�ݕ����z)
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)
    
    If u!�ԋ� < 0 Then
        k.�ؕ�����Ȗ� = "���|��"
        k.�ؕ��⏕�Ȗ� = "�|�C���g"
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!�ԋ� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = "����"
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = u!�ԋ� * -1
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
    
        Call ��v�o��(k)
    
    End If
    
End Sub

'************************************************************************
'�@  �\ :�����o�́i�R�����C�t����j
'************************************************************************
Private Sub ����_�R�����C�t�o��(ByRef k As type��v, ByVal u As ADODB.Recordset)
                    
    k.�ؕ�����Ȗ� = "���ʗa��"
    k.�ؕ��⏕�Ȗ� = u!��s
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = "�ΏۊO"
    k.�ؕ����z = u!���v���z '+ u!���̑��萔��
    k.�ؕ��ŋ��z = 0
    
    k.�ݕ�����Ȗ� = "���㍂"
    k.�ݕ��⏕�Ȗ� = "�����炢��"
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_����
    k.�ݕ����z = u!���v���z - u!���� + u!���̑��萔�� * -1
    k.�ݕ��ŋ��z = ����Ōv�Z(k.�ݕ����z)
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)
    
    If u!�ԋ� < 0 Then
        k.�ؕ�����Ȗ� = "���ʗa��"
        k.�ؕ��⏕�Ȗ� = u!��s
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!�ԋ� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = "����"
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = u!�ԋ� * -1
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
    
        Call ��v�o��(k)
    
    End If
    
    If u!���̑��萔�� < 0 Then
        k.�ؕ�����Ȗ� = "���|��"
        k.�ؕ��⏕�Ȗ� = "�|�C���g"
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!���̑��萔�� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = ""
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = 0
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
        
        Call ��v�o��(k)
    End If

End Sub

'************************************************************************
'�@  �\ :�y�V�o���N�o�́i�R�����C�t����j
'************************************************************************
Private Sub �y�V�o���N_�R�����C�t�o��(ByRef k As type��v, ByVal u As ADODB.Recordset)

    k.�ؕ�����Ȗ� = "���ʗa��"
    k.�ؕ��⏕�Ȗ� = "�y�V��s"
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = "�ΏۊO"
    k.�ؕ����z = u!���v���z - G_�U�荞�萔��
    k.�ؕ��ŋ��z = 0
    
    k.�ݕ�����Ȗ� = "���㍂"
    k.�ݕ��⏕�Ȗ� = "�����炢��"
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_����
    k.�ݕ����z = u!���v���z - u!���� + u!���̑��萔�� * -1
    k.�ݕ��ŋ��z = ����Ōv�Z(k.�ݕ����z)
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)
    
    If u!�ԋ� < 0 Then
        k.�ؕ�����Ȗ� = "���ʗa��"
        k.�ؕ��⏕�Ȗ� = u!��s
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!�ԋ� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = "����"
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = u!�ԋ� * -1
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
    
        Call ��v�o��(k)
    
    End If
    
    If u!���̑��萔�� < 0 Then
        k.�ؕ�����Ȗ� = "���|��"
        k.�ؕ��⏕�Ȗ� = "�|�C���g"
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!���̑��萔�� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = ""
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = 0
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
        
        Call ��v�o��(k)
    End If

    k.�ؕ�����Ȗ� = "�x���萔��"
    k.�ؕ��⏕�Ȗ� = "�U�荞�萔��"
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = G_�d��
    k.�ؕ����z = G_�U�荞�萔��
    k.�ؕ��ŋ��z = ����Ōv�Z(G_�U�荞�萔��)
    
    k.�ݕ�����Ȗ� = ""
    k.�ݕ��⏕�Ȗ� = ""
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = "�ΏۊO"
    k.�ݕ����z = 0
    k.�ݕ��ŋ��z = 0
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)
End Sub
'************************************************************************
'�@  �\ :���|���o�́i���̑�����j
'************************************************************************
Private Sub ���|��_���̑��o��(ByRef k As type��v, ByVal u As ADODB.Recordset)
                    
    k.�ؕ�����Ȗ� = "���|��"
    k.�ؕ��⏕�Ȗ� = �ؕ��⏕�Ȗڎ擾1(u!�������@, u!��z�Ǝ�)
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = "�ΏۊO"
    k.�ؕ����z = u!���v���z - u!����
    k.�ؕ��ŋ��z = 0
    
    k.�ݕ�����Ȗ� = "���㍂"
    If k.�ؕ��⏕�Ȗ� = "���t�I�N" Then
        k.�ݕ��⏕�Ȗ� = "���t�I�N"
    Else
        k.�ݕ��⏕�Ȗ� = "���̑�"
    End If
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_����
    k.�ݕ����z = u!���v���z + u!���̑��萔�� * -1 - u!����
    k.�ݕ��ŋ��z = ����Ōv�Z(k.�ݕ����z)
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)
    
    If u!���� > 0 Then
        Call �ב��^��_�o��2(k, u)
    End If
    
    If u!�ԋ� < 0 Then
        k.�ؕ�����Ȗ� = "���|��"
        k.�ؕ��⏕�Ȗ� = �ؕ��⏕�Ȗڎ擾1(u!�������@, u!��z�Ǝ�)
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!�ԋ� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = "����"
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = u!�ԋ� * -1
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
    
        Call ��v�o��(k)
    
    End If
    
    If u!���̑��萔�� < 0 Then
        k.�ؕ�����Ȗ� = "���|��"
        k.�ؕ��⏕�Ȗ� = "�|�C���g"
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!���̑��萔�� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = ""
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = 0
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
        
        Call ��v�o��(k)
        
    End If

End Sub

'************************************************************************
'�@  �\ :�����o�́i���̑��j
'************************************************************************
Private Sub ����_���̑��o��(ByRef k As type��v, ByVal u As ADODB.Recordset)
                    
    k.�ؕ�����Ȗ� = "���ʗa��"
    k.�ؕ��⏕�Ȗ� = u!��s
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = "�ΏۊO"
    k.�ؕ����z = u!���v���z - u!����
    k.�ؕ��ŋ��z = 0
    
    k.�ݕ�����Ȗ� = "���㍂"
    k.�ݕ��⏕�Ȗ� = "���̑�"
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_����
    k.�ݕ����z = u!���v���z + u!���̑��萔�� * -1 - u!����
    k.�ݕ��ŋ��z = ����Ōv�Z(k.�ݕ����z)
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)
    
    If u!���� > 0 Then
        Call �ב��^��_�o��2(k, u)
    End If

    If u!�ԋ� < 0 Then
        k.�ؕ�����Ȗ� = "���ʗa��"
        k.�ؕ��⏕�Ȗ� = u!��s
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!�ԋ� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = "����"
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = u!�ԋ� * -1
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
    
        Call ��v�o��(k)
    
    End If
    
    If u!���̑��萔�� < 0 Then
        k.�ؕ�����Ȗ� = "���|��"
        k.�ؕ��⏕�Ȗ� = "�|�C���g"
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!���̑��萔�� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = ""
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = 0
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
        
        Call ��v�o��(k)
        
    End If

End Sub


'************************************************************************
'�@  �\ :�y�V�o���N�o�́i���̑��j
'************************************************************************
Private Sub �y�V�o���N_���̑��o��(ByRef k As type��v, ByVal u As ADODB.Recordset)
                    
    k.�ؕ�����Ȗ� = "���ʗa��"
    k.�ؕ��⏕�Ȗ� = "�y�V��s"
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = "�ΏۊO"
    k.�ؕ����z = u!���v���z - u!���� - G_�U�荞�萔��
    k.�ؕ��ŋ��z = 0
    
    k.�ݕ�����Ȗ� = "���㍂"
    k.�ݕ��⏕�Ȗ� = "���̑�"
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_����
    k.�ݕ����z = u!���v���z + u!���̑��萔�� * -1 - u!����
    k.�ݕ��ŋ��z = ����Ōv�Z(k.�ݕ����z)
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)
        
    If u!���� > 0 Then
        Call �ב��^��_�o��2(k, u)
    End If

    If u!�ԋ� < 0 Then
        k.�ؕ�����Ȗ� = "���ʗa��"
        k.�ؕ��⏕�Ȗ� = "�y�V��s"
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!�ԋ� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = "����"
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = u!�ԋ� * -1
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
    
        Call ��v�o��(k)
    
    End If
    
    If u!���̑��萔�� < 0 Then
        k.�ؕ�����Ȗ� = "���|��"
        k.�ؕ��⏕�Ȗ� = "�|�C���g"
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!���̑��萔�� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = ""
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = 0
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
        
        Call ��v�o��(k)
        
    End If

    k.�ؕ�����Ȗ� = "�x���萔��"
    k.�ؕ��⏕�Ȗ� = "�U�荞�萔��"
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = G_�d��
    k.�ؕ����z = G_�U�荞�萔��
    k.�ؕ��ŋ��z = ����Ōv�Z(G_�U�荞�萔��)
    
    k.�ݕ�����Ȗ� = ""
    k.�ݕ��⏕�Ȗ� = ""
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = "�ΏۊO"
    k.�ݕ����z = 0
    k.�ݕ��ŋ��z = 0
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)

End Sub

'************************************************************************
'�@  �\ :���|���o�́i���̑��|�C���g�j
'************************************************************************
Private Sub ���|��_�|�C���g_���̑��o��(k As type��v, ByVal u As ADODB.Recordset)
                    
    k.�ؕ�����Ȗ� = "���|��"
    k.�ؕ��⏕�Ȗ� = "�|�C���g"
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = "�ΏۊO"
    k.�ؕ����z = u!���̑��萔�� * -1 - u!����
    k.�ؕ��ŋ��z = 0
    
    k.�ݕ�����Ȗ� = "���㍂"
    k.�ݕ��⏕�Ȗ� = "���̑�"
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_����
    k.�ݕ����z = u!���̑��萔�� * -1 - u!����
    k.�ݕ��ŋ��z = ����Ōv�Z(k.�ݕ����z)
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)
        
    If u!���� > 0 Then
        Call �ב��^��_�o��2(k, u)
    End If

    If u!�ԋ� < 0 Then
        k.�ؕ�����Ȗ� = "���|��"
        k.�ؕ��⏕�Ȗ� = "�|�C���g"
        k.�ؕ����� = "�S��"
        k.�ؕ��ŋ敪 = "�ΏۊO"
        k.�ؕ����z = u!�ԋ� * -1
        k.�ؕ��ŋ��z = 0
        
        k.�ݕ�����Ȗ� = "����"
        k.�ݕ��⏕�Ȗ� = ""
        k.�ݕ����� = "�S��"
        k.�ݕ��ŋ敪 = "�ΏۊO"
        k.�ݕ����z = u!�ԋ� * -1
        k.�ݕ��ŋ��z = 0
        
        k.�E�v = u!�ڋq��
    
        Call ��v�o��(k)
    
    End If

End Sub

'************************************************************************
'�@  �\ :���|���o�́i�C���t�H�g�b�v�j
'************************************************************************
Private Sub ���|��_�C���t�H�g�b�v�o��(ByRef k As type��v, ByVal u As ADODB.Recordset)
                    
    k.�ؕ�����Ȗ� = "���|��"
    k.�ؕ��⏕�Ȗ� = "�C���t�H�g�b�v"
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = "�ΏۊO"
    k.�ؕ����z = u!���v���z + u!���� * -1
    k.�ؕ��ŋ��z = 0
    
    k.�ݕ�����Ȗ� = "���㍂"
    k.�ݕ��⏕�Ȗ� = �ݕ��⏕�Ȗڎ擾1(u!���i��)
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_����
    k.�ݕ����z = u!���v���z + u!���̑��萔�� * -1 + u!���� * -1
    k.�ݕ��ŋ��z = ����Ōv�Z(k.�ݕ����z)
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)
    
    k.�ؕ�����Ȗ� = "���|��"
    k.�ؕ��⏕�Ȗ� = "�C���t�H�g�b�v"
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = "�ΏۊO"
    k.�ؕ����z = u!����
    k.�ؕ��ŋ��z = 0
    
    k.�ݕ�����Ȗ� = "�ב��^��������"
    k.�ݕ��⏕�Ȗ� = ""
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_�d��
    k.�ݕ����z = u!����
    k.�ݕ��ŋ��z = ����Ōv�Z(u!����)
    
    k.�E�v = u!�ڋq��

    Call ��v�o��(k)

End Sub

'************************************************************************
'�@  �\ :���|���o��
'************************************************************************
Private Sub ���|��_�����g���b�N�X�o��(ByRef k As type��v, ByVal u As ADODB.Recordset)
                    
    k.�ؕ�����Ȗ� = "���|��"
    k.�ؕ��⏕�Ȗ� = "�����g���b�N�X"
'    k.�ؕ��⏕�Ȗ� = "������̂��l�b�g"
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = "�ΏۊO"
    k.�ؕ����z = u!���v���z + u!���� * -1
    k.�ؕ��ŋ��z = 0
    
    k.�ݕ�����Ȗ� = "���㍂"
    k.�ݕ��⏕�Ȗ� = �ݕ��⏕�Ȗڎ擾1(u!���i��)
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_����
    k.�ݕ����z = u!���v���z + u!���̑��萔�� * -1 + u!���� * -1
    k.�ݕ��ŋ��z = ����Ōv�Z(k.�ݕ����z)
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)
    
    k.�ؕ�����Ȗ� = "���|��"
    k.�ؕ��⏕�Ȗ� = "�����g���b�N�X"
'    k.�ؕ��⏕�Ȗ� = "������̂��l�b�g"
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = "�ΏۊO"
    k.�ؕ����z = u!����
    k.�ؕ��ŋ��z = 0
    
    k.�ݕ�����Ȗ� = "�ב��^��������"
    k.�ݕ��⏕�Ȗ� = ""
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_�d��
    k.�ݕ����z = u!����
    k.�ݕ��ŋ��z = ����Ōv�Z(u!����)
    
    k.�E�v = u!�ڋq��

    Call ��v�o��(k)

End Sub


'************************************************************************
'�@  �\ :���|���o��
'************************************************************************
Private Sub ���|��_�o��(ByRef k As type��v, ByVal u As ADODB.Recordset)
                    
    k.�ؕ�����Ȗ� = "�d����"
    If G_�X�ܖ� = "�g���j�e�B�[�y�V�s��X" Then
        k.�ؕ��⏕�Ȗ� = "�����炢��"
    Else
        k.�ؕ��⏕�Ȗ� = "�R�����C�t"
    End If
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = G_�d��
    k.�ؕ����z = u!�d�����z
    k.�ؕ��ŋ��z = ����Ōv�Z(k.�ؕ����z)
    
    k.�ݕ�����Ȗ� = "���|��"
    k.�ݕ��⏕�Ȗ� = "�����炢��"
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = "�ΏۊO"
    k.�ݕ����z = u!�d�����z + u!�ב��^��
    k.�ݕ��ŋ��z = 0
    
    k.�E�v = u!�ڋq��
    
    Call ��v�o��(k)

End Sub

'************************************************************************
'�@  �\ :�ב��^���o��
'************************************************************************
Private Sub �ב��^��_�o��(ByRef k As type��v, ByVal u As ADODB.Recordset)
                    
    k.�ؕ�����Ȗ� = "�ב��^��������"
    k.�ؕ��⏕�Ȗ� = ""
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = G_�d��
    k.�ؕ����z = u!�ב��^��
    k.�ؕ��ŋ��z = ����Ōv�Z(k.�ؕ����z)
    
    k.�ݕ�����Ȗ� = "�ב��^��������"
    k.�ݕ��⏕�Ȗ� = ""
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_�d��
    k.�ݕ����z = u!����
    k.�ݕ��ŋ��z = ����Ōv�Z(k.�ݕ����z)
    
    k.�E�v = ""
    
    Call ��v�o��(k)

End Sub

'************************************************************************
'�@  �\ :�ב��^���o��
'************************************************************************
Private Sub �ב��^��_�o��2(ByRef k As type��v, ByVal u As ADODB.Recordset)
                    
'    k.�ؕ�����Ȗ� = "�ב��^��������"
'    k.�ؕ��⏕�Ȗ� = ""
    k.�ؕ����� = "�S��"
    k.�ؕ��ŋ敪 = "�ΏۊO"
    k.�ؕ����z = u!����
    k.�ؕ��ŋ��z = 0
    
    k.�ݕ�����Ȗ� = "�ב��^��������"
    k.�ݕ��⏕�Ȗ� = ""
    k.�ݕ����� = "�S��"
    k.�ݕ��ŋ敪 = G_�d��
    k.�ݕ����z = u!����
    k.�ݕ��ŋ��z = ����Ōv�Z(k.�ݕ����z)
    
    k.�E�v = ""
    
    Call ��v�o��(k)

End Sub


'************************************************************************
'�@  �\ :����Ōv�Z
'************************************************************************
Private Function ����Ōv�Z(ByVal ���z As Long)
    
    Dim ���z2       As Double
    Dim ���i���    As Long
    Dim �����      As Long
   
    ���z2 = ���z
    ���i��� = CLng(Format(CStr((���z2 / (G_����� + 1))), "0000000000"))
    
    ����Ōv�Z = ���z - ���i���
    
End Function

'************************************************************************
'�@  �\ :�ؕ��⏕�Ȗڎ擾
'************************************************************************
Private Function �ؕ��⏕�Ȗڎ擾1(ByVal �������@ As String, ByVal ��z�Ǝ� As String) As String
    
    �ؕ��⏕�Ȗڎ擾1 = "���̑�"
    
    ' �������@���u�N���W�b�g�v�̏ꍇ
    If �������@ = "�N���W�b�g" Then
        �ؕ��⏕�Ȗڎ擾1 = "�N���W�b�g"
        Exit Function
    End If
    
    ' �������@���u�����N���W�b�g�v�̏ꍇ
    If �������@ = "�����N���W�b�g" Then
        �ؕ��⏕�Ȗڎ擾1 = "�����N���W�b�g"
        Exit Function
    End If
                
    ' �������@���u���i������v�̏ꍇ
    If �������@ = "���i���" Then
        If ��z�Ǝ� = "����}��" Then
            �ؕ��⏕�Ȗڎ擾1 = "����}��"
        Else
            �ؕ��⏕�Ȗڎ擾1 = "�䂤�p�b�N"
        End If
        Exit Function
    End If
    
    ' �������@���u�|�C���g�v�̏ꍇ
    If �������@ = "�|�C���g" Then
        �ؕ��⏕�Ȗڎ擾1 = "�|�C���g"
        Exit Function
    End If
    
    ' �������@���u�㕥���v�̏ꍇ
    If �������@ = "�㕥��" Then
        �ؕ��⏕�Ȗڎ擾1 = "�㕥��"
        Exit Function
    End If
    
    ' �������@���u�y�C�W�[�v�̏ꍇ
    If �������@ = "�y�C�W�[" Then
        �ؕ��⏕�Ȗڎ擾1 = "�y�C�W�["
        Exit Function
    End If
    
    ' �������@���u�R���r�j�v�̏ꍇ
    If �������@ = "�R���r�j" Then
        �ؕ��⏕�Ȗڎ擾1 = "�R���r�j"
        Exit Function
    End If
    
    ' �������@���u�g�ь��ρv�̏ꍇ
    If �������@ = "�g�ь���" Then
        �ؕ��⏕�Ȗڎ擾1 = "�g�ь���"
        Exit Function
    End If
    
    ' �������@���u���t�I�N�v�̏ꍇ
    If �������@ = "���t�I�N" Then
        �ؕ��⏕�Ȗڎ擾1 = "���t�I�N"
        Exit Function
    End If
    
    ' �������@���u�d�q�}�l�[�v�̏ꍇ
    If �������@ = "�d�q�}�l�[" Then
        �ؕ��⏕�Ȗڎ擾1 = "�d�q�}�l�["
        Exit Function
    End If
    
End Function

'************************************************************************
'�@  �\ :�ؕ��⏕�Ȗڎ擾
'************************************************************************
Private Function �ؕ��⏕�Ȗڎ擾2(ByVal �������@ As String, ByVal ��z�Ǝ� As String) As String
    
    �ؕ��⏕�Ȗڎ擾2 = "���̑�"
    
    ' �������@���u�N���W�b�g�v�̏ꍇ
    If �������@ = "�N���W�b�g" Then
        �ؕ��⏕�Ȗڎ擾2 = "�N���W�b�g"
        Exit Function
    End If
                
    ' �������@���u�����N���W�b�g�v�̏ꍇ
    If �������@ = "�����N���W�b�g" Then
        �ؕ��⏕�Ȗڎ擾2 = "�����N���W�b�g"
        Exit Function
    End If
                
    ' �������@���u���i������v�̏ꍇ
    If �������@ = "���i���" Then
        �ؕ��⏕�Ȗڎ擾2 = "�����炢��"
        Exit Function
    End If
    
    ' �������@���u�|�C���g�v�̏ꍇ
    If �������@ = "�|�C���g" Then
        �ؕ��⏕�Ȗڎ擾2 = "�|�C���g"
        Exit Function
    End If
    
    ' �������@���u�㕥���v�̏ꍇ
    If �������@ = "�㕥��" Then
        �ؕ��⏕�Ȗڎ擾2 = "�㕥��"
        Exit Function
    End If
    
    ' �������@���u�y�C�W�[�v�̏ꍇ
    If �������@ = "�y�C�W�[" Then
        �ؕ��⏕�Ȗڎ擾2 = "�y�C�W�["
        Exit Function
    End If
    
    ' �������@���u�R���r�j�v�̏ꍇ
    If �������@ = "�R���r�j" Then
        �ؕ��⏕�Ȗڎ擾2 = "�R���r�j"
        Exit Function
    End If
    
    ' �������@���u�g�ь��ρv�̏ꍇ
    If �������@ = "�g�ь���" Then
        �ؕ��⏕�Ȗڎ擾2 = "�g�ь���"
        Exit Function
    End If
    
    ' �������@���u���t�I�N�v�̏ꍇ
    If �������@ = "���t�I�N" Then
        �ؕ��⏕�Ȗڎ擾2 = "���t�I�N"
        Exit Function
    End If
    
    ' �������@���u�d�q�}�l�[�v�̏ꍇ
    If �������@ = "�d�q�}�l�[" Then
        �ؕ��⏕�Ȗڎ擾2 = "�d�q�}�l�["
        Exit Function
    End If
    
End Function

'************************************************************************
'�@  �\ :�ؕ��⏕�Ȗڎ擾
'************************************************************************
Private Function �ݕ��⏕�Ȗڎ擾1(ByVal ���i�� As String) As String

    �ݕ��⏕�Ȗڎ擾1 = ���i��
    
    If ���i�� = "�A�[�f���{�V�����v�[" Then
        If G_�X�ܖ� = "�g���j�e�B�[�y�V�s��X" Then
            �ݕ��⏕�Ȗڎ擾1 = "�Z�b�g��"
        Else
            �ݕ��⏕�Ȗڎ擾1 = "�Z�b�g"
        End If
    End If
    
    If ���i�� = "�A�[�f��2�{�Z�b�g" Then
        �ݕ��⏕�Ȗڎ擾1 = "�A�[�f��"
    End If
    
    If ���i�� = "�V�����v�[2�{�Z�b�g" Then
        �ݕ��⏕�Ȗڎ擾1 = "�V�����v�["
    End If
    
    If ���i�� = "�A�[�f�����V�����v�[�����i" Then
        �ݕ��⏕�Ȗڎ擾1 = "�����i"
    End If
    
    If ���i�� = "���C�X�g���b�` �N�����W���O" Or _
       ���i�� = "���C�X�g���b�` �E�H�b�V���O" Or _
       ���i�� = "���C�X�g���b�` ���[�V����" Or _
       ���i�� = "���C�X�g���b�` �W�F��" Or _
       ���i�� = "���C�X�g���b�` ���C�����G�b�Z���X" Or _
       ���i�� = "���C�X�g���b�` ��b���ϕi�Z�b�g" Then
       �ݕ��⏕�Ȗڎ擾1 = "Ӳ��د�"
    End If

End Function

'************************************************************************
'�@  �\ :�A�[�f�����i����
'************************************************************************
Private Function �A�[�f������(ByVal ���i�� As String) As Integer
    
    �A�[�f������ = 2
    
    If ���i�� = "�A�[�f��" Or ���i�� = "�A�[�f��2�{�Z�b�g" Or _
       ���i�� = "�A�[�f���{�V�����v�[" Or _
       ���i�� = "�V�u�X�^" Or _
       ���i�� = "�V�u�X�^�{�V�����v�[" Or _
       ���i�� = "�u�[�X�^�[" Or _
       ���i�� = "�u�[�X�^�[�i�v���ь��ԁj" Or _
       ���i�� = "�u�[�X�^�[�{�V�����v�[" Or _
       ���i�� = "�V�n�C�u���b�^�[" Or _
       ���i�� = "�V�n�C�u���b�^�[�{�V�����v�[" Or _
       ���i�� = "�n�C�u���b�h" Or _
       ���i�� = "�n�C�u���b�h�{�V�����v�[" Or _
       ���i�� = "�i�C�X���f�B�[" Or _
       ���i�� = "�i�C�X���f�B�[�{�V�����v�[" Or _
       ���i�� = "�n�C�u���b�h�i�v���[���g�j" Or _
       ���i�� = "�V�����v�[" Or _
       ���i�� = "�V�����v�[2�{�Z�b�g" Or _
       ���i�� = "�V�����v�[�i�v���[���g�j" Or _
       ���i�� = "�V�����v�[�{�g���[�g�����g" Or _
       ���i�� = "�g���[�g�����g" Or _
       ���i�� = "�g���[�g�����g�i�v���[���g�j" Or _
       ���i�� = "�A�[�f�����V�����v�[�����i" Or _
       ���i�� = "�A�[�f�������i" Or _
       ���i�� = "�V�����v�[�����i" Then
       
       �A�[�f������ = 1
       
    End If
    
    If ���i�� = "�A�[�f�����p�E�}�j���A���i�v���[���g�j" Or _
       ���i�� = "�����̐ςݏd�˂���؂ł��E�}�j���A���i�v���[���g�j" Or _
       ���i�� = "�h�N�^�[�A�[�f���E��тc�u�c�i�v���[���g�j" Or _
       ���i�� = "��тƉ^���E�}�j���A���i�v���[���g�j" Or _
       ���i�� = "��сE���у}�j���A���i�v���[���g�j" Then
       
       �A�[�f������ = 9
       
    End If
    
End Function

'************************************************************************
'�@  �\ :����f�[�^��CSV���o�͂���
'************************************************************************
Private Sub ��vCSV�o��()
    
    Dim intFileNo       As Integer
    Dim ����f�[�^RS As New ADODB.Recordset
    
    intFileNo = FreeFile()
    
    Call ����f�[�^�Ǎ�(����f�[�^RS)
    
    If G_�X�ܖ� = "�g���j�e�B�[�y�V�s��X" Then
        Open "C:\�ڋq�Ǘ�\�y�V_����.csv" For Output As #intFileNo
    Else
        Open "C:\�ڋq�Ǘ�\Yahoo_����.csv" For Output As #intFileNo
    End If
    
    With ����f�[�^RS
        Do Until .EOF
    
'        Print #intFileNo, !���ʃt���O & "," & !�`�[NO & "," & !���Z & "," & !������� & "," & !�ؕ�����Ȗ� & "," & !�ؕ��⏕�Ȗ� & "," & !�ؕ����� & "," & _
'                            !�ؕ��ŋ敪 & "," & !�ؕ����z & "," & !�ؕ��ŋ��z & "," & !�ݕ�����Ȗ� & "," & !�ݕ��⏕�Ȗ� & "," & _
'                            !�ݕ����� & "," & !�ݕ��ŋ敪 & "," & !�ݕ����z & "," & !�ݕ��ŋ��z & "," & !�E�v & "," & _
'                            !�ԍ� & "," & !���� & "," & !�^�C�v & "," & !������ & "," & !�d������ & "," & !�t�1 & "," & !�t�2 & "," & !����
        
        Print #intFileNo, !���ʃt���O1 & !���ʃt���O2 & !���ʃt���O3 & "," & !�`�[NO & "," & !������� & "," & _
                            !�ؕ�����Ȗ� & "," & !�ؕ��⏕�Ȗ� & "," & !�ؕ����� & "," & !�ؕ��ŋ敪 & "," & !�ؕ����z & "," & _
                            !�ݕ�����Ȗ� & "," & !�ݕ��⏕�Ȗ� & "," & !�ݕ����� & "," & !�ݕ��ŋ敪 & "," & !�ݕ����z & "," & _
                            !�E�v & "," & !�^�C�v & "," & !������ & "," & "0" & "," & "0" & "," & !�ؕ��ŋ��z & "," & !�ݕ��ŋ��z & "," & "no" & "," & "no" & "," & "no" & "," & """"""
        .MoveNext
        Loop
        
        .Close
    End With
    
    Close #intFileNo
    
End Sub


'************************************************************************
'�@  �\ :Yahoo�����X�e�[�^�X�ύX
'************************************************************************
Private Sub �y�V_�����X�e�[�^�X�ύX()
    
    Dim intFileNo1      As Integer
    Dim intFileNo2      As Integer
    Dim ����f�[�^RS    As New ADODB.Recordset
    Dim �����f�[�^RS    As New ADODB.Recordset
    Dim �����ԍ�        As String
    Dim �����ԍ�w       As String
    Dim �ʒu            As Integer
    Dim �z����          As String
    Dim �t���O1         As Boolean
    Dim �t���O2         As Boolean
    Dim ����ID          As String
    
    �t���O1 = False
    �t���O2 = False
    
    ' FileSystemObject (FSO) �̐V�����C���X�^���X�𐶐�����
    Dim cFso As FileSystemObject
    Set cFso = New FileSystemObject

    ' �t�@�C�����폜����
    On Error Resume Next
    Call cFso.DeleteFile("C:\�ڋq�Ǘ�\rakuten_status_001.csv")
    On Error Resume Next
    Call cFso.DeleteFile("C:\�ڋq�Ǘ�\rakuten_status_002.csv")

    ' �s�v�ɂȂ������_�ŎQ�Ƃ�������� (Terminate �C�x���g�𑁂߂ɋN����)
    Set cFso = Nothing
    
    Call ����f�[�^�Ǎ�(����f�[�^RS)
    
    intFileNo1 = FreeFile()
    Open "C:\�ڋq�Ǘ�\rakuten_status_001.csv" For Output As #intFileNo1
    
    intFileNo2 = FreeFile()
    Open "C:\�ڋq�Ǘ�\rakuten_status_002.csv" For Output As #intFileNo2
    
    Print #intFileNo1, """�󒍔ԍ�""" + "," + """�󒍃X�e�[�^�X""" + "," + """�z����""" + "," + """���ו��`�[�ԍ�"""
    
    Print #intFileNo2, """�����w���󒍔ԍ�""" + "," + """�󒍃X�e�[�^�X""" + "," + """�z����""" + "," + """���ו��`�[�ԍ�"""
    
    �����ԍ�w = ""
    
    With ����f�[�^RS
        Do Until .EOF
            If !�ؕ�����Ȗ� = "���|��" Or !�ؕ�����Ȗ� = "���ʗa��" Then
                
                �ʒu = InStr(!�����ԍ�, "#")
                
                If �ʒu > 0 Then
                    �����ԍ� = Left(!�����ԍ�, �ʒu - 1)
                Else
                    �����ԍ� = Trim(!�����ԍ�)
                End If
                
                ����ID = !����ID
                
                If �����ԍ� <> �����ԍ�w Then
                    Call �����ԍ�����(����ID, �����f�[�^RS)
                    
                    �z���� = """" + Mid(�����f�[�^RS!�o�ד�, 1, 4) + "-" + Mid(�����f�[�^RS!�o�ד�, 6, 2) + "-" + Mid(�����f�[�^RS!�o�ד�, 9, 2) + """"
                    
                    If Not �����f�[�^RS.EOF Then
                        If �����f�[�^RS!������ = "�y�V" Then
                            
                            If InStr(�����ԍ�, "-g") > 0 Then
                                ' �����w��
                                Print #intFileNo2, """" + �����ԍ� + """" + "," + """������""" + "," + �z���� + "," + """" + Trim(�����f�[�^RS!�⍇�ԍ�) + """"
                                �t���O2 = True
                            Else
                                ' �ʏ�w��
                                Print #intFileNo1, """" + �����ԍ� + """" + "," + """������""" + "," + �z���� + "," + """" + Trim(�����f�[�^RS!�⍇�ԍ�) + """"
                                �t���O1 = True
                            End If
                        End If
                    End If
                    
                    �����f�[�^RS.Close
                    �����ԍ�w = �����ԍ�
                
                End If
            End If
            .MoveNext
        Loop
        
        .Close
    End With
    
    Close #intFileNo1
    Close #intFileNo2
    
    ' FileSystemObject (FSO) �̐V�����C���X�^���X�𐶐�����
    Set cFso = New FileSystemObject
    
    ' �t�@�C�����폜����
    If �t���O1 = False Then
        On Error Resume Next
        Call cFso.DeleteFile("C:\�ڋq�Ǘ�\rakuten_status_001.csv")
    End If
    
    If �t���O2 = False Then
        On Error Resume Next
        Call cFso.DeleteFile("C:\�ڋq�Ǘ�\rakuten_status_002.csv")
    End If
    
    Set cFso = Nothing
    
End Sub

'************************************************************************
'�@  �\ :Yahoo�����X�e�[�^�X�ύX
'************************************************************************
Private Sub Yahoo_�����X�e�[�^�X�ύX()
    
    Dim intFileNo       As Integer
    Dim ����f�[�^RS    As New ADODB.Recordset
    Dim �����f�[�^RS    As New ADODB.Recordset
    Dim �����ԍ�        As String
    Dim �����ԍ�w       As String
    Dim �ʒu            As Integer
    Dim ����ID          As String
    
    ' FileSystemObject (FSO) �̐V�����C���X�^���X�𐶐�����
    Dim cFso As FileSystemObject
    Set cFso = New FileSystemObject

    ' �t�@�C�����폜����
    On Error Resume Next
    Call cFso.DeleteFile("C:\�ڋq�Ǘ�\Yahoo_status.csv")

    ' �s�v�ɂȂ������_�ŎQ�Ƃ�������� (Terminate �C�x���g�𑁂߂ɋN����)
    Set cFso = Nothing

    intFileNo = FreeFile()
    
    Call ����f�[�^�Ǎ�(����f�[�^RS)
    
    Open "C:\�ڋq�Ǘ�\Yahoo_status.csv" For Output As #intFileNo
    
    Print #intFileNo, """OrderID""" + "," + """Status""" + "," + """Quantity1""" + "," + """Shipping1""" + "," + """Paymentcharge1""" + "," + """Gift Wrap1""" + "," + """Discount1"""
    
    �����ԍ�w = ""
    
    With ����f�[�^RS
        Do Until .EOF
            If !�ؕ�����Ȗ� = "���|��" Or !�ؕ�����Ȗ� = "���ʗa��" Then
                
                �ʒu = InStr(!�����ԍ�, "#")
                
                If �ʒu > 0 Then
                    �����ԍ� = Left(!�����ԍ�, �ʒu - 1)
                Else
                    �����ԍ� = Trim(!�����ԍ�)
                End If
                  
                ����ID = !����ID
                
                If �����ԍ� <> �����ԍ�w Then
                    Call �����ԍ�����(����ID, �����f�[�^RS)
                    
                    If Not �����f�[�^RS.EOF Then
                        If �����f�[�^RS!������ = "Yahoo" Then
                            Print #intFileNo, �����ԍ� + "," + """����""" + "," + "" + "," + "" + "," + "" + "," + "" + "," + ""
                        End If
                    End If
                    
                    �����f�[�^RS.Close
                    �����ԍ�w = �����ԍ�
                
                End If
            End If
            .MoveNext
        Loop
        
        .Close
    End With
    
    Close #intFileNo
    
End Sub


'************************************************************************
'�@  �\ :�I�[�g�V�b�v��ʂ�\������
'************************************************************************
Private Sub cmd�I�[�g�V�b�v_Click()

    Dim ADF016      As New ADF016

    Call ADF016.Show(1)
    
End Sub

'************************************************************************
'�@  �\ :�ꊇ��������
'************************************************************************
Private Sub cmd�ꊇ����_Click()

    Dim ADF019      As New ADF019

    Call ADF019.Show(1)

    'Call cmd����_Click
    
    Call cmd���o�׈ꗗ_Click
    
End Sub

'************************************************************************
'�@  �\ :�X�֔ԍ�����Z����\������
'************************************************************************
Private Sub �X�֔ԍ�����Z����ϊ�����()
    
    Dim �X�֔ԍ�        As String
    Dim �X�֔ԍ�����RS  As New ADODB.Recordset
    Dim ADF014          As New ADF014
    Dim ����            As Integer
    Dim �Z��_��i       As String
    Dim �Z��_���i       As String
    
    If txt�Z��_��i = "" Then
        
        �X�֔ԍ� = txt�X�֔ԍ�.Text
        
        If Len(�X�֔ԍ�) = 7 Then
            �X�֔ԍ� = Mid(�X�֔ԍ�, 1, 3) & "-" & Mid(�X�֔ԍ�, 4, 4)
            txt�X�֔ԍ�.Text = �X�֔ԍ�
        End If
        
        If Len(�X�֔ԍ�) = 8 Then
            ���� = �Z����������(�X�֔ԍ�)
            
            If ���� > 1 Then
                Call ADF014.SET_�X�֔ԍ�(�X�֔ԍ�)
                Call ADF014.Show(1)
                Call ADF014.GET_�Z��(�Z��_��i, �Z��_���i)
                txt�Z��_��i.Text = �Z��_��i
                txt�Z��_���i.Text = �Z��_���i
            Else
                Call �Z������(�X�֔ԍ�, �X�֔ԍ�����RS)
                If Not �X�֔ԍ�����RS.EOF Then
                    txt�Z��_��i.Text = �X�֔ԍ�����RS!�s���{���� + �X�֔ԍ�����RS!�s�撬����
                    txt�Z��_���i.Text = �X�֔ԍ�����RS!���於
                End If
            
                �X�֔ԍ�����RS.Close
            End If
        End If
    End If
    
End Sub

'************************************************************************
'�@  �\ :�ڋq���N���A����B
'************************************************************************
Private Sub �ڋq���N���A()
        
    txt�ڋqID.Text = ""
    txt�ڋq��.Text = ""
    txt�t���K�i.Text = ""
    txt�X�֔ԍ�.Text = ""
    txt�Z��_��i.Text = ""
    txt�Z��_���i.Text = ""
    txt�Z��_���i.Text = ""
    txt�d�b�ԍ�.Text = ""
    txt���[��.Text = ""
    txt�y�V���[��.Text = ""
    cmb�A�[�f���N���u.ListIndex = 0
    txt�����.Text = "____/__/__"
    txt�މ��.Text = "____/__/__"
    opt�j��.Value = True
    opt����.Value = False
    txt���l.Text = ""
    chk���[�����M = 1
    txt�a����.Text = "____/__/__"

End Sub

'************************************************************************
'�@  �\ :�������N���A����B
'************************************************************************
Private Sub �������N���A()
        
    txt�󒍓�.Text = "____/__/__"
    txt����ID.Text = ""
    txt�����ԍ�.Text = ""
    cmb�X�e�[�^�X.ListIndex = 0
    cmb���i��.ListIndex = 0
    cmb�������@.ListIndex = 0
    txt�z�B����.Text = ""
    txt�o�ד�.Text = "____/__/__"
    cmb��z�Ǝ�.ListIndex = 0
    txt�x���ԍ�.Text = ""
    txt�⍇�ԍ�.Text = ""
    txt�P��.Value = 0
    txt����.Value = 0
    txt����.Value = 0
    txt����.Value = 0
    txt�ԋ�.Value = 0
    txt���̑��萔��.Value = 0
    txt���v���z.Text = 0
    txt���[�����M.Text = ""
    cmb������.Text = ""
    txt���l2.Text = ""
    txt�R�����C�t.Text = ""
    txt�o�ח\���.Text = "____/__/__"
    txt����URL.Text = ""

End Sub

'************************************************************************
'�@  �\ :�ڋq����
'************************************************************************
Private Sub cmd����_Click()

    Dim �����l As String
    Dim �ڋq�}�X�^RS As New ADODB.Recordset

    If cmb��������.Text = "" Then Exit Sub
    
    MousePointer = vbHourglass
    
    If cmb��������.Text = "�o�ד�" Then
        ' ���͂��ꂽ�o�ד������ɁA�ڋq�������s�Ȃ�
        Call �ڋq����2(txt��������.Text, �ڋq�}�X�^RS)
    ElseIf cmb��������.Text = "����ID" Then
    
        ' ���͂��ꂽ����ID�������s��
        Call �ڋq����3(txt��������.Text, �ڋq�}�X�^RS)
    ElseIf cmb��������.Text = "�R�����C�tNO" Then
    
        ' ���͂��ꂽ�R�����C�tNO�������s��
        Call �ڋq����4(txt��������.Text, �ڋq�}�X�^RS)
    ElseIf cmb��������.Text = "�����ԍ�" Then
    
        ' ���͂��ꂽ�����ԍ��������s��
        Call �ڋq����5(txt��������.Text, �ڋq�}�X�^RS)
        
    ElseIf cmb��������.Text = "��" Then
    
        ' ���͂��ꂽ���������s��
        �����l = txt��������.Text
        Call �ڋq����(�����l, �ڋq�}�X�^RS, "[��]")
        
    ElseIf cmb��������.Text = "�⍇�ԍ�" Then
        If InStr(txt��������.Text, "-") <= 0 Then
            �����l = �⍇�ԍ��ҏW(txt��������.Text)
        Else
            �����l = txt��������.Text
        End If
            
        ' ���͂��ꂽ�ڋq�������Ƀ��C���h�J�[�h�������s��
        Call �ڋq����6(�����l, �ڋq�}�X�^RS)
    Else
        �����l = "%" & txt��������.Text & "%"
        
        ' ���͂��ꂽ�ڋq�������Ƀ��C���h�J�[�h�������s��
        Call �ڋq����(�����l, �ڋq�}�X�^RS, cmb��������.Text)
    End If
    
    ' �ڋq���X�g��\������
    Call �ڋq���X�g�\��(�ڋq�}�X�^RS)
    
    MousePointer = vbNormal

End Sub

'************************************************************************
'�@  �\ :�⍇�ԍ���ҏW����
'************************************************************************
Private Function �⍇�ԍ��ҏW(ByVal �⍇�ԍ� As String) As String
    
    Dim �⍇�ԍ�1 As String
    Dim �⍇�ԍ�2 As String
    Dim �⍇�ԍ�3 As String
    
    �⍇�ԍ�1 = Mid(�⍇�ԍ�, 1, 4)
    �⍇�ԍ�2 = Mid(�⍇�ԍ�, 5, 4)
    �⍇�ԍ�3 = Mid(�⍇�ԍ�, 9, 4)
    
    �⍇�ԍ��ҏW = �⍇�ԍ�1 & "-" & �⍇�ԍ�2 & "-" & �⍇�ԍ�3
    
End Function

'************************************************************************
'�@  �\ :���o�ׂ���������
'************************************************************************
Private Sub cmd���o�׈ꗗ_Click()

    Dim �ڋq�}�X�^RS As New ADODB.Recordset
    
    MousePointer = vbHourglass
        
    Call �g�����U�N�V�����f�[�^�̍X�V
    
    ' ���o�׈ꗗ���擾����
    Call ���o�׌���(�ڋq�}�X�^RS)

    ' �ڋq���X�g��\������
    Call �ڋq���X�g�\��(�ڋq�}�X�^RS)
        
    MousePointer = vbNormal

End Sub

'************************************************************************
'�@  �\ :����������������
'************************************************************************
Private Sub cmd������_Click()

    Dim �ڋq�}�X�^RS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call �g�����U�N�V�����f�[�^�̍X�V
    
    ' ���o�׈ꗗ���擾����
    Call ����������(�ڋq�}�X�^RS)

    ' �ڋq���X�g��\������
    Call �ڋq���X�g�\��(�ڋq�}�X�^RS)
    
    MousePointer = vbNormal

End Sub

'************************************************************************
'�@  �\ :�o�ח\��ꗗ���o�͂���
'************************************************************************
Private Sub cmd�o�ח\��ꗗ_Click()

    Dim �o�ח\��ꗗRS As New ADODB.Recordset
    Dim ADF012 As New ADF012
    
    Call �g�����U�N�V�����f�[�^�̍X�V
    
    MousePointer = vbHourglass
    
    ' ���o�׈ꗗ���擾����
    Call �o�ח\��ꗗ(�o�ח\��ꗗRS)
    
    MousePointer = vbNormal
    
   
    ' �m�F���b�Z�[�W��\������
    'If MsgBox("�[�i����������Ă�낵���ł����H", vbYesNo, "�ڋq�Ǘ�") <> vbYes Then Exit Sub
    
    If �o�ח\��ꗗRS.EOF Then
        Call MsgBox("�o�ח\�肪����܂���", vbOKOnly, "�ڋq�Ǘ�")
        Exit Sub
    End If
    
    Set G_�o�ח\�胊�X�g = Nothing
    Set G_�o�ח\�胊�X�g = New �o�ח\�胊�X�g
    Call G_�o�ח\�胊�X�g.Database.SetDataSource(�o�ח\��ꗗRS)
    Call ADF012.�����ݒ�("�o�ח\�胊�X�g")
    Call ADF012.Show(vbModal)
    �o�ח\��ꗗRS.Close

End Sub

'************************************************************************
'�@  �\ :�A�[�f���w����
'************************************************************************
Private Sub cmd�A�[�f���w����_Click()

    Dim �ڋq�}�X�^RS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call �g�����U�N�V�����f�[�^�̍X�V
    
    ' �A�[�f���w���҂��擾����
    Call �A�[�f���w���Ҍ���(�ڋq�}�X�^RS)

    ' �ڋq���X�g��\������
    Call �ڋq���X�g�\��(�ڋq�}�X�^RS)
    
    MousePointer = vbNormal

End Sub

'************************************************************************
'�@  �\ :�A�[�f���N���u���
'************************************************************************
Private Sub cmd�N���u����_Click()

    Dim �ڋq�}�X�^RS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call �g�����U�N�V�����f�[�^�̍X�V
    
    ' �A�[�f���N���u������擾����
    Call �A�[�f���N���u�������(�ڋq�}�X�^RS)

    ' �ڋq���X�g��\������
    Call �ڋq���X�g�\��(�ڋq�}�X�^RS)
    
    MousePointer = vbNormal

End Sub

'************************************************************************
'�@  �\ :�A�[�f���N���u����������
'************************************************************************
Private Sub cmd�N���u������_Click()

    Dim i As Integer
    Dim �ڋq�}�X�^RS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call �g�����U�N�V�����f�[�^�̍X�V
    
    ' �A�[�f���N���u���[�������M�ڋq���擾����
    Call ���[������3(�ڋq�}�X�^RS)

    ' �ڋq���X�g��\������
    Call �ڋq���X�g�\��(�ڋq�}�X�^RS)

    MousePointer = vbNormal

End Sub


'************************************************************************
'�@  �\ :�A�[�f����������
'************************************************************************
Private Sub cmd��������_Click()

    Dim �ڋq�}�X�^RS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call �g�����U�N�V�����f�[�^�̍X�V
    
    ' �A�[�f�������������s��
    Call ��������(�ڋq�}�X�^RS)

    ' �ڋq���X�g��\������
    Call �ڋq���X�g�\��(�ڋq�}�X�^RS)

    MousePointer = vbNormal

End Sub

'************************************************************************
'�@  �\ :�R�����C�t�̎d�����z�A�ב��^���̌v�Z���s��
'************************************************************************
Private Sub cmd�v�Z_Click()
    
    Dim �d�����z    As Long
    Dim �ב��^��    As Long
    
    Dim ADF017      As New ADF017

    Call ADF017.Show(1)

    Call ADF017.�d�����z_�ב��^���擾(�d�����z, �ב��^��)
    
    txt�d�����z.Value = �d�����z
    txt�ב��^��.Value = �ב��^��
    
    Call ����_�X�V
    
End Sub

'************************************************************************
'�@  �\ :�V�K�����̌������s��
'************************************************************************
Private Sub cmd�V�K����_Click()

    Dim �ڋq�}�X�^RS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call �g�����U�N�V�����f�[�^�̍X�V

    ' �V�K�����̌������s��
    Call �X�e�[�^�X����("�V�K����", �ڋq�}�X�^RS)

    ' �ڋq���X�g��\������
    Call �ڋq���X�g�\��(�ڋq�}�X�^RS)

    MousePointer = vbNormal

End Sub

'************************************************************************
'�@  �\ :�����҂��̌������s��
'************************************************************************
Private Sub cmd�����҂�_Click()

    Dim �ڋq�}�X�^RS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call �g�����U�N�V�����f�[�^�̍X�V
    
    ' �����҂��̌������s��
    Call �X�e�[�^�X����("�����҂�", �ڋq�}�X�^RS)

    ' �ڋq���X�g��\������
    Call �ڋq���X�g�\��(�ڋq�}�X�^RS)

    MousePointer = vbNormal

End Sub

'************************************************************************
'�@  �\ :�o�׏������̌������s��
'************************************************************************
Private Sub cmd�o�׏�����_Click()

    Dim �ڋq�}�X�^RS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call �g�����U�N�V�����f�[�^�̍X�V
    
    ' �o�׏������̌������s��
    Call �X�e�[�^�X����("�o�׏���", �ڋq�}�X�^RS)

    ' �ڋq���X�g��\������
    Call �ڋq���X�g�\��(�ڋq�}�X�^RS)

    MousePointer = vbNormal

End Sub

'************************************************************************
'�@  �\ :�o�׍ς݂̌������s��
'************************************************************************
Private Sub cmd�o�׍ς�_Click()

    Dim �ڋq�}�X�^RS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call �g�����U�N�V�����f�[�^�̍X�V
    
    ' �o�׍ς݂̌������s��
    Call �X�e�[�^�X����("�o�׊���", �ڋq�}�X�^RS)

    ' �ڋq���X�g��\������
    Call �ڋq���X�g�\��(�ڋq�}�X�^RS)

    MousePointer = vbNormal

End Sub

'************************************************************************
'�@  �\ :�R�����C�t�������s��
'************************************************************************
Private Sub cmd�R�����C�t_Click()

    Dim �ڋq�}�X�^RS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call �g�����U�N�V�����f�[�^�̍X�V
    
    ' �ۗ����̌������s��
    Call �X�e�[�^�X����("�R�����C�t", �ڋq�}�X�^RS)

    ' �ڋq���X�g��\������
    Call �ڋq���X�g�\��(�ڋq�}�X�^RS)

    MousePointer = vbNormal


End Sub

'************************************************************************
'�@  �\ :�ۗ����̌������s��
'************************************************************************
Private Sub cmd�ۗ�������_Click()

    Dim �ڋq�}�X�^RS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call �g�����U�N�V�����f�[�^�̍X�V
    
    ' �ۗ����̌������s��
    Call �X�e�[�^�X����("�ۗ�", �ڋq�}�X�^RS)

    ' �ڋq���X�g��\������
    Call �ڋq���X�g�\��(�ڋq�}�X�^RS)

    MousePointer = vbNormal

End Sub

'************************************************************************
'�@  �\ :�L�����Z���̌������s��
'************************************************************************
Private Sub cmd�L�����Z������_Click()

    Dim �ڋq�}�X�^RS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call �g�����U�N�V�����f�[�^�̍X�V
    
    ' �L�����Z���̌������s��
    Call �X�e�[�^�X����("�L�����Z��", �ڋq�}�X�^RS)

    ' �ڋq���X�g��\������
    Call �ڋq���X�g�\��(�ڋq�}�X�^RS)

    MousePointer = vbNormal

End Sub

'************************************************************************
'�@  �\ :�S�ڋq�I��
'************************************************************************
Private Sub cmd�S�I��_Click()

    Dim i As Integer
    
    MousePointer = vbHourglass
    
    Call �g�����U�N�V�����f�[�^�̍X�V
    
    For i = 1 To va�ڋq���X�g.MaxRows
        Call SpreadSetVal(va�ڋq���X�g, i, COL_�`�F�b�N, "1")
    Next i
    
    MousePointer = vbNormal

End Sub

'************************************************************************
'�@  �\ :�S�ڋq�I������
'************************************************************************
Private Sub cmd�S����_Click()

    Dim i As Integer
    
    MousePointer = vbHourglass
    
    Call �g�����U�N�V�����f�[�^�̍X�V
    
    For i = 1 To va�ڋq���X�g.MaxRows
        Call SpreadSetVal(va�ڋq���X�g, i, COL_�`�F�b�N, "0")
    Next i
    
    MousePointer = vbNormal

End Sub

'************************************************************************
'�@  �\ :�[�i�����������B
'************************************************************************
Private Sub cmd�[�i��_Click()
    
    Dim i As Integer
    Dim ������ As String
        
    Call �g�����U�N�V�����f�[�^�̍X�V

    If �`�F�b�N�����擾() <= 0 Then
        Call MsgBox("�[�i����������閾�ׂɃ`�F�b�N��t���ĉ�����", vbOKOnly, "�ڋq�Ǘ�")
        Exit Sub
    End If
    
    For i = 1 To va�������X�g.MaxRows
        If SpreadGetVal(va�������X�g, i, COL_�`�F�b�N) = "1" Then
            ������ = SpreadGetVal(va�������X�g, i, COL_������)
            
            If ������ <> "�������" Then
                Call cmd�[�i��_sub1
                Exit For
            Else
                Call cmd�[�i��_sub2
                Exit For
            End If
        End If
    Next i
    
    'Call Sleep(3000)

End Sub

'************************************************************************
'�@  �\ :�[�i�����������i�L���b�g�n���h�p�j
'************************************************************************
Private Sub cmd�[�i��_sub1()

    Dim i As Integer
    Dim �ڋqID As String
    Dim ����ID As String
    Dim ADF012 As New ADF012
    Dim �[�i��RS As New ADODB.Recordset
   
    ' �m�F���b�Z�[�W��\������
    'If MsgBox("�[�i����������Ă�낵���ł����H", vbYesNo, "�ڋq�Ǘ�") <> vbYes Then Exit Sub
    
    �ڋqID = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�ڋqID)
    ����ID = ""
    
    For i = 1 To va�������X�g.MaxRows
        If SpreadGetVal(va�������X�g, i, COL_�`�F�b�N) = "1" Then
            ����ID = ����ID & SpreadGetVal2(va�������X�g, i, COL_����ID) & ","
        End If
    Next i
        
    If ����ID <> "" Then
        ����ID = Left(����ID, Len(����ID) - 1)
        Call �[�i�f�[�^�擾(�ڋqID, ����ID, �[�i��RS)
        If Not �[�i��RS.EOF Then
            Set G_�[�i�� = Nothing
            Set G_�[�i�� = New �[�i��
            Call G_�[�i��.Database.SetDataSource(�[�i��RS)
            Call ADF012.�����ݒ�("�[�i��")
            Call ADF012.Show(vbModal)
        End If
        �[�i��RS.Close
    End If

End Sub
'************************************************************************
'�@  �\ :�[�i�����������i�A�[�f�����p�j
'************************************************************************
Private Sub cmd�[�i��_sub2()

    Dim i As Integer
    Dim �ڋqID As String
    Dim ����ID As String
    Dim ADF012 As New ADF012
    Dim �[�i��RS As New ADODB.Recordset
   
    ' �m�F���b�Z�[�W��\������
    'If MsgBox("�[�i����������Ă�낵���ł����H", vbYesNo, "�ڋq�Ǘ�") <> vbYes Then Exit Sub
    
    �ڋqID = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�ڋqID)
    ����ID = ""
    
    For i = 1 To va�������X�g.MaxRows
        If SpreadGetVal(va�������X�g, i, COL_�`�F�b�N) = "1" Then
            ����ID = ����ID & SpreadGetVal2(va�������X�g, i, COL_����ID) & ","
        End If
    Next i
        
    If ����ID <> "" Then
        ����ID = Left(����ID, Len(����ID) - 1)
        Call �[�i�f�[�^�擾(�ڋqID, ����ID, �[�i��RS)
        If Not �[�i��RS.EOF Then
            Set G_�[�i��2 = Nothing
            Set G_�[�i��2 = New �[�i��2
            Call G_�[�i��2.Database.SetDataSource(�[�i��RS)
            Call ADF012.�����ݒ�("�[�i��2")
            Call ADF012.Show(vbModal)
        End If
        �[�i��RS.Close
    End If
    
End Sub

'************************************************************************
'�@  �\ �������������B
'************************************************************************
Private Sub cmd���_Click()
    Dim i As Integer
    Dim ������ As String
    
    Call �g�����U�N�V�����f�[�^�̍X�V

    If �`�F�b�N�����擾() <= 0 Then
        Call MsgBox("������������閾�ׂɃ`�F�b�N��t���ĉ�����", vbOKOnly, "�ڋq�Ǘ�")
        Exit Sub
    End If
    
    For i = 1 To va�������X�g.MaxRows
        If SpreadGetVal(va�������X�g, i, COL_�`�F�b�N) = "1" Then
            ������ = SpreadGetVal(va�������X�g, i, COL_������)
            
            If ������ <> "�������" Then
                Call cmd���_sub1(i)
            Else
                Call cmd���_sub2(i)
            End If
        End If
    Next i

End Sub

'************************************************************************
'�@  �\ �������������i�L���b�g�n���h�p�j
'************************************************************************
Private Sub cmd���_sub1(ByVal i As Integer)

    Dim �ڋq�� As String
    Dim ���i�� As String
    Dim ADF012 As New ADF012
    Dim �����RS As New ADODB.Recordset
    Dim ���ӎ���RS As New ADODB.Recordset
    Dim �~�j�܂�RS As New ADODB.Recordset
    
    If SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_���͂��於) <> "" Then
        �ڋq�� = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_���͂��於)
    Else
        �ڋq�� = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�ڋq��)
    End If
    
    ���i�� = SpreadGetVal(va�������X�g, i, COL_���i��)
        
    ' �m�F���b�Z�[�W��\������
    'If MsgBox("������������Ă�낵���ł����H", vbYesNo, "�ڋq�Ǘ�") <> vbYes Then Exit Sub
    
    If Left(���i��, 4) = "�A�[�f��" Then
        
        Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
        
        If Not �����RS.EOF Then
            Set G_����� = Nothing
            Set G_����� = New �����
            Call G_�����.Database.SetDataSource(�����RS)
            Call ADF012.�����ݒ�("�����")
            Call ADF012.Show(vbModal)
        End If
        
        �����RS.Close
        
        If ���i�� <> "�A�[�f������" Then
            
            Set G_���ӎ��� = Nothing
            Set G_���ӎ��� = New ���ӎ���
            Call ADF012.�����ݒ�("���ӎ���")
            Call ADF012.Show(vbModal)
            
        End If
        
        If InStr(1, ���i��, "�V�����v�[") > 0 Then
    
            Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
            
            If Not �����RS.EOF Then
                Set G_�����3 = Nothing
                Set G_�����3 = New �����3
                Call G_�����3.Database.SetDataSource(�����RS)
                Call ADF012.�����ݒ�("�����3")
                Call ADF012.Show(vbModal)
            End If
            
            �����RS.Close
        End If
    
    ElseIf Left(���i��, 13) = "�V�����v�[�{�g���[�g�����g" Then
    
        Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
        
        If Not �����RS.EOF Then
            Set G_�����3 = Nothing
            Set G_�����3 = New �����3
            Call G_�����3.Database.SetDataSource(�����RS)
            Call ADF012.�����ݒ�("�����3")
            Call ADF012.Show(vbModal)
        End If
        
        �����RS.Close
        
        Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
        
        If Not �����RS.EOF Then
            Set G_�����7 = Nothing
            Set G_�����7 = New �����7
            Call G_�����7.Database.SetDataSource(�����RS)
            Call ADF012.�����ݒ�("�����7")
            Call ADF012.Show(vbModal)
        End If
        
        �����RS.Close
        
    ElseIf Left(���i��, 5) = "�V�����v�[" Then
    
        Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
        
        If Not �����RS.EOF Then
            Set G_�����3 = Nothing
            Set G_�����3 = New �����3
            Call G_�����3.Database.SetDataSource(�����RS)
            Call ADF012.�����ݒ�("�����3")
            Call ADF012.Show(vbModal)
        End If
        
        �����RS.Close
        
    ElseIf Left(���i��, 7) = "�g���[�g�����g" Then
    
        Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
        
        If Not �����RS.EOF Then
            Set G_�����7 = Nothing
            Set G_�����7 = New �����7
            Call G_�����7.Database.SetDataSource(�����RS)
            Call ADF012.�����ݒ�("�����7")
            Call ADF012.Show(vbModal)
        End If
        
        �����RS.Close
        
    ElseIf Left(���i��, 5) = "�u�[�X�^�[" Then
    
        Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
        
        If ���i�� = "�u�[�X�^�[�i�v���ь��ԁj" Then
            If Not �����RS.EOF Then
                Set G_�����11 = Nothing
                Set G_�����11 = New �����11
                Call G_�����11.Database.SetDataSource(�����RS)
                Call ADF012.�����ݒ�("�����11")
                Call ADF012.Show(vbModal)
            End If
        Else
            If Not �����RS.EOF Then
                Set G_�����4 = Nothing
                Set G_�����4 = New �����4
                Call G_�����4.Database.SetDataSource(�����RS)
                Call ADF012.�����ݒ�("�����4")
                Call ADF012.Show(vbModal)
            End If
        End If
        
        �����RS.Close
    
        
        If InStr(1, ���i��, "�V�����v�[") > 0 Then
    
            Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
            
            If Not �����RS.EOF Then
                Set G_�����3 = Nothing
                Set G_�����3 = New �����3
                Call G_�����3.Database.SetDataSource(�����RS)
                Call ADF012.�����ݒ�("�����3")
                Call ADF012.Show(vbModal)
            End If
            
            �����RS.Close
        End If
    
    ElseIf Left(���i��, 6) = "�n�C�u���b�h" Then
    
        Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
        
        If Not �����RS.EOF Then
            Set G_�����5 = Nothing
            Set G_�����5 = New �����5
            Call G_�����5.Database.SetDataSource(�����RS)
            Call ADF012.�����ݒ�("�����5")
            Call ADF012.Show(vbModal)
        End If
        
        �����RS.Close
    
        
        If InStr(1, ���i��, "�V�����v�[") > 0 Then
    
            Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
            
            If Not �����RS.EOF Then
                Set G_�����3 = Nothing
                Set G_�����3 = New �����3
                Call G_�����3.Database.SetDataSource(�����RS)
                Call ADF012.�����ݒ�("�����3")
                Call ADF012.Show(vbModal)
            End If
            
            �����RS.Close
        End If
    
    ElseIf Left(���i��, 7) = "�i�C�X���f�B�[" Then
    
        Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
        
        If Not �����RS.EOF Then
            Set G_�����6 = Nothing
            Set G_�����6 = New �����6
            Call G_�����6.Database.SetDataSource(�����RS)
            Call ADF012.�����ݒ�("�����6")
            Call ADF012.Show(vbModal)
        End If
        
        �����RS.Close
    
        
        If InStr(1, ���i��, "�V�����v�[") > 0 Then
    
            Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
            
            If Not �����RS.EOF Then
                Set G_�����3 = Nothing
                Set G_�����3 = New �����3
                Call G_�����3.Database.SetDataSource(�����RS)
                Call ADF012.�����ݒ�("�����3")
                Call ADF012.Show(vbModal)
            End If
            
            �����RS.Close
        End If
    
    ElseIf Left(���i��, 4) = "�V�u�X�^" Then
    
        Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
        
        If Not �����RS.EOF Then
            Set G_�����8 = Nothing
            Set G_�����8 = New �����8
            Call G_�����8.Database.SetDataSource(�����RS)
            Call ADF012.�����ݒ�("�����8")
            Call ADF012.Show(vbModal)
        End If
        
        �����RS.Close
    
        
        If InStr(1, ���i��, "�V�����v�[") > 0 Then
    
            Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
            
            If Not �����RS.EOF Then
                Set G_�����3 = Nothing
                Set G_�����3 = New �����3
                Call G_�����3.Database.SetDataSource(�����RS)
                Call ADF012.�����ݒ�("�����3")
                Call ADF012.Show(vbModal)
            End If
            
            �����RS.Close
        End If
    
    ElseIf Left(���i��, 8) = "�V�n�C�u���b�^�[" Then
    
        Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
        
        If Not �����RS.EOF Then
            Set G_�����10 = Nothing
            Set G_�����10 = New �����10
            Call G_�����10.Database.SetDataSource(�����RS)
            Call ADF012.�����ݒ�("�����10")
            Call ADF012.Show(vbModal)
        End If
        
        �����RS.Close
    
        
        If InStr(1, ���i��, "�V�����v�[") > 0 Then
    
            Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
            
            If Not �����RS.EOF Then
                Set G_�����3 = Nothing
                Set G_�����3 = New �����3
                Call G_�����3.Database.SetDataSource(�����RS)
                Call ADF012.�����ݒ�("�����3")
                Call ADF012.Show(vbModal)
            End If
            
            �����RS.Close
        End If
    
    
    ElseIf ���i�� = "�~�j�܂�" Then
    
        Call �~�j�܂��f�[�^�擾(�ڋq��, �~�j�܂�RS)
        
        If Not �~�j�܂�RS.EOF Then
            Set G_�~�j�܂� = Nothing
            Set G_�~�j�܂� = New �~�j�܂�
            Call G_�~�j�܂�.Database.SetDataSource(�~�j�܂�RS)
            Call ADF012.�����ݒ�("�~�j�܂�")
            Call ADF012.Show(vbModal)
        End If
        
        �~�j�܂�RS.Close
    Else
        Call MsgBox("�w�肵�����i�̂����̓T�|�[�g����Ă��܂���B", vbOK, "�ڋq�Ǘ�")

    End If
    
End Sub

'************************************************************************
'�@  �\ �������������i�A�[�f�������p�j
'************************************************************************
Private Sub cmd���_sub2(ByVal i As Integer)

    Dim �ڋq�� As String
    Dim ���i�� As String
    Dim ADF012 As New ADF012
    Dim �����RS As New ADODB.Recordset
    Dim ���ӎ���RS As New ADODB.Recordset

    If SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_���͂��於) <> "" Then
        �ڋq�� = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_���͂��於)
    Else
        �ڋq�� = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�ڋq��)
    End If
    
    ���i�� = SpreadGetVal(va�������X�g, i, COL_���i��)
        
    ' �m�F���b�Z�[�W��\������
    'If MsgBox("������������Ă�낵���ł����H", vbYesNo, "�ڋq�Ǘ�") <> vbYes Then Exit Sub
    
    If ���i�� = "�A�[�f��" Or ���i�� = "�A�[�f��2�{�Z�b�g" Or ���i�� = "�A�[�f�������i" Or ���i�� = "�A�[�f��(�Z�[��)" Or ���i�� = "�A�[�f�����V�����v�[�����i" Or ���i�� = "�A�[�f������" Then
        
        Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
        
        If Not �����RS.EOF Then
            Set G_�����2 = Nothing
            Set G_�����2 = New �����2
            Call G_�����2.Database.SetDataSource(�����RS)
            Call ADF012.�����ݒ�("�����2")
            Call ADF012.Show(vbModal)
        End If
        
        �����RS.Close
        
        If ���i�� <> "�A�[�f������" Then
            
            Set G_���ӎ��� = Nothing
            Set G_���ӎ��� = New ���ӎ���
            Call ADF012.�����ݒ�("���ӎ���")
            Call ADF012.Show(vbModal)
            
        End If
    ElseIf ���i�� = "�A�[�f���V�����v�[" Or ���i�� = "�A�[�f���V�����v�[2�{�Z�b�g" Or ���i�� = "�A�[�f���V�����v�[�����i" Or ���i�� = "�A�[�f�����V�����v�[�����i" Then
    
        Call �����f�[�^�擾(�ڋq��, ���i��, �����RS)
        
        If Not �����RS.EOF Then
            Set G_�����3 = Nothing
            Set G_�����3 = New �����3
            Call G_�����3.Database.SetDataSource(�����RS)
            Call ADF012.�����ݒ�("�����3")
            Call ADF012.Show(vbModal)
        End If
        
        �����RS.Close
    
    End If
    
End Sub

'************************************************************************
'�@  �\ ���[�����s���B
'************************************************************************
Private Sub cmd���[��_Click()
    
    Dim �ڋq��      As String
    Dim ���[��ID    As String
    Dim ADF015      As New ADF015
    Dim i           As Integer
        
    Call �g�����U�N�V�����f�[�^�̍X�V

    If �`�F�b�N�����擾() < 1 Then
        Call MsgBox("���[�����閾�ׂɂP���ȏ�`�F�b�N��t���ĉ�����", vbOKOnly, "�ڋq�Ǘ�")
        Exit Sub
    End If
    
    For i = 1 To va�������X�g.MaxRows
        If SpreadGetVal(va�������X�g, i, COL_�`�F�b�N) = "1" Then
        
            G_����ROW = i
            
            ' ���������o�^�̏ꍇ�G���[���b�Z�[�W��\������
            If SpreadGetVal(va�������X�g, G_����ROW, COL_����ID) = "-1" Then
                Call MsgBox("�悸������o�^���ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
                Exit Sub
            End If
            
            If SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_���͂��於) <> "" Then
                �ڋq�� = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_���͂��於)
            Else
                �ڋq�� = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�ڋq��)
            End If
            
            If SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�y�V���[��) <> "" Then
                ���[��ID = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_�y�V���[��)
            Else
                ���[��ID = SpreadGetVal(va�ڋq���X�g, G_�ڋq���X�g_ROW, COL_���[��)
            End If
        
            If �ڋq�� = "" Then
                Call MsgBox("�ڋq������͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
                Exit Sub
            End If
        
            If ���[��ID = "" Then
                Call MsgBox("���[��ID����͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
                Exit Sub
            End If
            
        End If
    Next i
            
    Call ADF015.Show(1)
    
End Sub

'************************************************************************
'�@  �\ �ʃ��[�����s���B
'************************************************************************
Private Sub cmd�ʃ��[��_Click()
    
    Dim ADF020      As New ADF020

    Call ADF020.Show(1)

End Sub

'************************************************************************
'�@  �\ �A�[�f���N���u���[�����s���B
'************************************************************************
Private Sub cmd�A�[�f���N���u_Click()
    
    Dim cnt             As Integer
    Dim �ڋqID          As String
    Dim �ڋq��          As String
    Dim ���[��ID        As String
    Dim row             As Integer
    Dim ���[���{��RS    As New ADODB.Recordset
    Dim ���[�����e      As String
    Dim �T�[�o          As String
    Dim ����            As String
    Dim ���M��          As String
    Dim ����            As String
    Dim ret             As String
    
    If MsgBox("�A�[�f���N���u���[���𑗐M���Ă���낵���ł����H", vbYesNo, "�ڋq�Ǘ�") = vbNo Then
        Exit Sub
    End If

    MousePointer = vbHourglass
        
    Call �g�����U�N�V�����f�[�^�̍X�V

    cnt = 0
    For row = 1 To va�ڋq���X�g.MaxRows
    
        If SpreadGetVal(va�ڋq���X�g, row, COL_�`�F�b�N) = "1" Then
            
            �ڋq�� = SpreadGetVal(va�ڋq���X�g, row, COL_�ڋq��)
            If SpreadGetVal(va�ڋq���X�g, row, COL_�y�V���[��) <> "" Then
                ���[��ID = SpreadGetVal(va�ڋq���X�g, row, COL_�y�V���[��)
            Else
                ���[��ID = SpreadGetVal(va�ڋq���X�g, row, COL_���[��)
            End If
        
            If �ڋq�� = "" Then
                MousePointer = vbNormal
                Call MsgBox("�ڋq������͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
                Exit Sub
            End If
        
            'If ���[��ID = "" Then
            '    MousePointer = vbNormal
            '    Call MsgBox("���[��ID����͂��ĉ������B", vbOKOnly, "�ڋq�Ǘ�")
            '    Exit Sub
            'End If
            
            cnt = cnt + 1
        End If
    Next
        
    If cnt <= 0 Then
        MousePointer = vbNormal
        Call MsgBox("���[��������ڋq�ɂP���ȏ�`�F�b�N��t���ĉ�����", vbOKOnly, "�ڋq�Ǘ�")
        Exit Sub
    End If
    
    Call ���[���{������(8, ���[���{��RS)
    
    For row = 1 To va�ڋq���X�g.MaxRows
    
        If SpreadGetVal(va�ڋq���X�g, row, COL_�`�F�b�N) = "1" Then
            
            �ڋqID = SpreadGetVal(va�ڋq���X�g, row, COL_�ڋqID)
            �ڋq�� = SpreadGetVal(va�ڋq���X�g, row, COL_�ڋq��)
            If SpreadGetVal(va�ڋq���X�g, row, COL_�y�V���[��) <> "" Then
                ���[��ID = SpreadGetVal(va�ڋq���X�g, row, COL_�y�V���[��)
            Else
                ���[��ID = SpreadGetVal(va�ڋq���X�g, row, COL_���[��)
            End If
            
            If ���[��ID <> "" Then
                ���� = ���[��ID ' + Chr(9) + "info@cathand.jp"    ' ����
                ���� = ���[���{��RS!����                        ' ����
                ���[�����e = ""
                ���[�����e = ���[�����e + �ڋq�� + "�l" + Chr$(13) + Chr$(10)
                ���[�����e = ���[�����e + Chr$(13) + Chr$(10)
                ���[�����e = ���[�����e + ���[���{��RS!����1 + Chr$(13) + Chr$(10)
                '���[�����e = ���[�����e + "�����[�����s�v�ȏꍇ�A�u�����O�v�𖾋L�̏�A�u���[���s�v�v�Ƃ��ĕԐM�������B" + Chr$(13) + Chr$(10)
            
                ' ���[�����M
                ret = SendMail(G_�T�[�o, ����, G_���M��, ����, ���[�����e, "")
                                
                If Len(ret) <> 0 Then
                   'Call MsgBox("���[�����M�G���[�F" & ret, vbOKOnly, "�ڋq�Ǘ�")
                End If
                
                Sleep (1000 * 3)
                
                'If ���[�����e <> "" Then
                '    Shell "..\bin\sendmail " + "|" + ���[�����e + "|"
                '    Call ���[�����M�ғo�^3(�ڋqID)
                'End If
            End If
        End If
    Next
    
    Call MsgBox("���[���𑗐M���܂���", vbOKOnly, "�ڋq�Ǘ�")
    
    If ���[���{��RS.State <> adStateClosed Then
        ���[���{��RS.Close
    End If
    
    MousePointer = vbNormal
    
End Sub


'************************************************************************
'�@  �\ �����}�K�𔭍s����B
'************************************************************************
Private Sub cmd�����}�K���s_Click()
    
    Dim �ڋq��          As String
    Dim ���[�����e      As String
    Dim ����            As String
    Dim ����            As String
    Dim �����}�K���M�\���  As String
    Dim ret             As String
    Dim �����}�KRS      As New ADODB.Recordset
    Dim �ڋq�}�X�^RS    As New ADODB.Recordset
    
    If MsgBox("�����}�K�𑗐M���Ă���낵���ł����H", vbYesNo, "�ڋq�Ǘ�") = vbNo Then
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    
    Call �g�����U�N�V�����f�[�^�̍X�V
    
    ' �ڋq�}�X�^��S�����[�h����
    Call �ڋq�}�X�^�Ǎ�2(�ڋq�}�X�^RS)
    
    With �ڋq�}�X�^RS
        Do Until .EOF
            If (!���[�� <> "" Or !�y�V���[�� <> "") And !�����}�KNO >= 0 Then
                
#If 0 Then
                 If Format(!�����}�K���M�\���, "yyyy/mm/dd") <= Format(Now, "yyyy/mm/dd") Or IsNull(!�����}�K���M�\���) Then
                'If Format(!�����}�K���M�\���, "yyyy/mm/dd") <= Format(Now, "yyyy/mm/dd") Then
                    
                    Call �����}�K�{������(IIf(!�����}�KNO <= 0, 1, !�����}�KNO), �����}�KRS)
                    'Call �����}�K�{������(0, �����}�KRS)
                    
                    If Not �����}�KRS.EOF Then
                        
                        If !�y�V���[�� <> "" Then
                            ���� = !�y�V���[�� ' + Chr(9) + "info@cathand.jp"    ' ����
                        Else
                            ���� = !���[�� ' + Chr(9) + "info@cathand.jp"    ' ����
                        End If
                        '���� = !�ڋq�� + "�l " + "���[���}�K�W�� [" & IIf(!�����}�KNO <= 0, 1, !�����}�KNO) & "] ��"            ' ����
                        '���� = !�ڋq�� + "�l " + "���������p���肪�Ƃ��������܂�"            ' ����
                        ���� = !�ڋq�� + "�l " + "��сE���эU���@�������v���[���g"            ' ����
                        
                        ���[�����e = ""
                        ���[�����e = ���[�����e + !�ڋq�� + "�l" + Chr$(13) + Chr$(10)
                        ���[�����e = ���[�����e + Chr$(13) + Chr$(10)
                        ���[�����e = ���[�����e + �����}�KRS!�����}�K + Chr$(13) + Chr$(10)
                        
                        ' ���[�����M
                        ret = SendMail(G_�T�[�o, ����, G_���M��, ����, ���[�����e, "")
                    
                        If Len(ret) <> 0 Then
                           'Call MsgBox("���[�����M�G���[�F" & ret, vbOKOnly, "�ڋq�Ǘ�")
                        End If
                        
                        �����}�K���M�\��� = Format(DateAdd("d", 7, Now), "yyyy/mm/dd")
                        Call �����}�K���sNO�X�V(!�ڋqID, IIf(!�����}�KNO <= 0, 2, !�����}�KNO + 1), "'" + �����}�K���M�\��� + "'")
                        'Call �����}�K���sNO�X�V(!�ڋqID, -1, "NULL")
                    
                        Sleep (1000 * 3)
                    
                    End If
                    
                    �����}�KRS.Close
                    
                End If
#Else
                    
                Call �����}�K�{������(0, �����}�KRS)
                
                If Not �����}�KRS.EOF Then
                    
                    If !�y�V���[�� <> "" Then
                        ���� = !�y�V���[�� ' + Chr(9) + "info@cathand.jp"    ' ����
                    Else
                        ���� = !���[�� ' + Chr(9) + "info@cathand.jp"    ' ����
                    End If
                    ���� = !�ڋq�� + "�l " + "��э܃A�[�f���I�v���[���g����t���I"            ' ����
                    
                    ���[�����e = ""
                    ���[�����e = ���[�����e + !�ڋq�� + "�l" + Chr$(13) + Chr$(10)
                    ���[�����e = ���[�����e + Chr$(13) + Chr$(10)
                    ���[�����e = ���[�����e + �����}�KRS!�����}�K + Chr$(13) + Chr$(10)
                    ���[�����e = Replace(���[�����e, "##########", !�ڋqID)
                    
                    ' ���[�����M
                    ret = SendMail(G_�T�[�o, ����, G_���M��, ����, ���[�����e, "")
                
                    If Len(ret) <> 0 Then
                       'Call MsgBox("���[�����M�G���[�F" & ret, vbOKOnly, "�ڋq�Ǘ�")
                    End If
                    
                    Sleep (1000 * 3)
                
                End If
                
                �����}�KRS.Close

#End If

            End If
            
            .MoveNext
        Loop
        
        .Close
    End With
    
    Call MsgBox("���[���𑗐M���܂���", vbOKOnly, "�ڋq�Ǘ�")
        
    MousePointer = vbNormal
    
End Sub

'************************************************************************
'�@  �\�F�e���v���[�g����
'************************************************************************
Private Sub cmd�e���v���[�g_Click()
    
    Dim �e���v���[�g As String
    Dim �\��    As String
    
    �e���v���[�g = cmb�e���v���[�g.Text
    
    If �e���v���[�g = "�A�[�f���V�K" Then
        �\�� = "�A�[�f��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�P�O�������p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�ԋ��p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f������" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�A�[�f�����V�����v�[�Z�b�g�V�K" Then
        �\�� = "�A�[�f�����V�����v�[�Z�b�g" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�P�O�������p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�ԋ��p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�`���V" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f������" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�V�u�X�^�V�K" Then
        �\�� = "�V�u�X�^" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�ԋ��p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f������" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�V�u�X�^���V�����v�[�Z�b�g�V�K" Then
        �\�� = "�V�u�X�^" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�V�����v�[" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�ԋ��p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f������" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�u�[�X�^�[�V�K" Then
        �\�� = "�u�[�X�^�[" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�ԋ��p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f������" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�u�[�X�^�[���V�����v�[�Z�b�g�V�K" Then
        �\�� = "�u�[�X�^�[" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�V�����v�[" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�ԋ��p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f������" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�V�n�C�u���b�^�[�V�K" Then
        �\�� = "�V�n�C�u���b�^�[" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�ԋ��p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f������" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�V�n�C�u���b�^�[���V�����v�[�Z�b�g�V�K" Then
        �\�� = "�V�n�C�u���b�^�[" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�V�����v�[" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�ԋ��p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f������" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�n�C�u���b�h�V�K" Then
        �\�� = "�n�C�u���b�h" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�ԋ��p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f������" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�n�C�u���b�h���V�����v�[�Z�b�g�V�K" Then
        �\�� = "�n�C�u���b�h" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�V�����v�[" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�ԋ��p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f������" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�i�C�X���f�B�[�V�K" Then
        �\�� = "�i�C�X���f�B�[" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�ԋ��p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f������" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�i�C�X���f�B�[���V�����v�[�Z�b�g�V�K" Then
        �\�� = "�i�C�X���f�B�[" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�V�����v�[" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�ԋ��p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f������" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�V�����v�[�Q�{�Z�b�g�V�K" Then
        �\�� = "�V�����v�[�Q�{" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�ԋ��p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f������" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�V�����v�[�V�K" Then
        �\�� = "�V�����v�[" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�ԋ��p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f������" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�V�����v�[���g���[�g�����g�V�K" Then
        �\�� = "�V�����v�[���g���[�g�����g" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�ԋ��p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f������" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�g���[�g�����g�V�K" Then
        �\�� = "�g���[�g�����g" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�ԋ��p��" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f������" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�����i�V�K" Then
        �\�� = "�����i�Z�b�g" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f������" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�A�[�f�����p" Then
        �\�� = "" 'Chr$(13) + Chr$(10)
        �\�� = �\�� + "�`���V" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�A�[�f�����p�E�}�j���A��" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "�����̐ςݏd�˂���؂ł�" Then
        �\�� = "" 'Chr$(13) + Chr$(10)
        �\�� = �\�� + "�`���V" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�����̐ςݏd�˂���؂ł��E�}�j���A��" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "��тc�u�c" Then
        �\�� = "" 'Chr$(13) + Chr$(10)
        �\�� = �\�� + "�`���V" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "�h�N�^�[�A�[�f���E��тc�u�c" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "��тƉ^��" Then
        �\�� = "" 'Chr$(13) + Chr$(10)
        �\�� = �\�� + "�`���V" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "��тƉ^���E�}�j���A��" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If
    
    If �e���v���[�g = "��сE����" Then
        �\�� = "" 'Chr$(13) + Chr$(10)
        �\�� = �\�� + "�`���V" + Chr$(13) + Chr$(10)
        �\�� = �\�� + "��сE���у}�j���A��" + Chr$(13) + Chr$(10)
        txt���l2.Text = txt���l2.Text + �\��
    End If

    Call ����_�X�V

End Sub

'************************************************************************
'�@  �\�F�]���{�^��
'************************************************************************
Private Sub cmd�]��_Click()
    
    Call �]���X�V(txt�ڋqID.Text)
    
End Sub

'************************************************************************
'�@  �\�F�g�����U�N�V�����f�[�^�̍X�V
'************************************************************************
Private Sub �g�����U�N�V�����f�[�^�̍X�V()
    
#If 0 Then
    Select Case G_�^�uNO
        Case 1
            Call �ڋq���_�o�^
        Case 2
            Call �ڋq���_�o�^
        Case 3
            Call ����_�X�V
    End Select
#End If
End Sub

'************************************************************************
'�@  �\�F���̓`�F�b�N
'************************************************************************
Private Function ���̓`�F�b�N(ByVal ���͒l As String) As Boolean

    Dim �ʒu As Integer
    
    �ʒu = InStr(���͒l, "'")
    
    If �ʒu > 0 Then
        ���̓`�F�b�N = True
    Else
        ���̓`�F�b�N = False
    End If
    
End Function

'************************************************************************
'�@  �\�FIME �I��/�I�t �؂�ւ�
'************************************************************************
Private Sub psubIMEOnOff(ByVal hwnd As Long, ByVal booOnOff As Boolean)
    Dim himc As Long    'IME�n���h��
    'IME�n���h���擾
    himc = ImmGetContext(hwnd)
    'IME�؂�ւ�
    Call ImmSetOpenStatus(himc, booOnOff)
    'IME�n���h�����
    Call ImmReleaseContext(hwnd, himc)
End Sub

'************************************************************************
'�@  �\�FIME���[�h�̐؂�ւ�
'************************************************************************
Private Sub ImeMode(Index As Integer)
    Dim himc As Long            'IME�n���h��
    Dim lngConversion As Long   '���̓��[�h
    Dim lngSentence As Long     '���[�h��
    'IME�n���h���擾
    himc = ImmGetContext(Me.hwnd)
    'IME�X�e�[�^�X�擾
    If Not ImmGetOpenStatus(himc) Then
        'IME�؂�ւ�
        Call ImmSetOpenStatus(himc, 1)
    End If
    'IME���̓��[�h�擾
    Call ImmGetConversionStatus(himc, lngConversion, lngSentence)
    'IME���̓��[�h�ݒ�
    Select Case Index
    Case 0  '�S�p�Ђ炪��
        lngConversion = MY_IME_CHMODE_ZEN_HIRA
    Case 1  '�S�p�J�^�J�i
        lngConversion = MY_IME_CHMODE_ZEN_KATA
    Case 2  '�S�p�p��
        lngConversion = MY_IME_CHMODE_ZEN_EISU
    Case 3  '���p�J�^�J�i
        lngConversion = MY_IME_CHMODE_HAN_KATA
    Case 4  '���p�p��
        lngConversion = MY_IME_CHMODE_HAN_EISU
    End Select
    Call ImmSetConversionStatus(himc, lngConversion, lngSentence)
    'IME�n���h�����
    Call ImmReleaseContext(Me.hwnd, himc)
End Sub

'************************************************************************
'�@  �\ :�����f���X�V
'************************************************************************
Private Sub cmd�X�V_Click()
    
    Dim row             As Long
    Dim ������          As String
    Dim ��������        As String
    Dim �N���W�b�g      As String
    Dim �����N���W�b�g  As String
    Dim ���i���        As String
    Dim �R���r�j        As String
    Dim ��s�U��        As String
    Dim �y�V�o���N����  As String
    Dim �y�C�W�[        As String
    Dim �㕥            As String
    Dim �|�C���g        As String
    Dim �g�ь���        As String
    Dim �d�q�}�l�[      As String
    Dim ���t�I�N        As String
    
    If G_�X�ܖ� = "�g���j�e�B�[�y�V�s��X" Then
        ������ = "�y�V"
    Else
        ������ = "Yahoo"
    End If
    
    row = 1
    Call ���������擾(������, ��������, �N���W�b�g, �����N���W�b�g, ���i���, �R���r�j, ��s�U��, �y�V�o���N����, �y�C�W�[, �㕥, �|�C���g, �g�ь���, �d�q�}�l�[, ���t�I�N)
    
    Call SpreadSetVal(va�����f����, row, 2, ��������)
    Call SpreadSetVal(va�����f����, row, 3, �N���W�b�g)
    Call SpreadSetVal(va�����f����, row, 4, �����N���W�b�g)
    Call SpreadSetVal(va�����f����, row, 5, ���i���)
    Call SpreadSetVal(va�����f����, row, 6, �R���r�j)
    Call SpreadSetVal(va�����f����, row, 7, ��s�U��)
    Call SpreadSetVal(va�����f����, row, 8, �y�V�o���N����)
    Call SpreadSetVal(va�����f����, row, 9, �y�C�W�[)
    Call SpreadSetVal(va�����f����, row, 10, �㕥)
    Call SpreadSetVal(va�����f����, row, 11, �|�C���g)
    Call SpreadSetVal(va�����f����, row, 12, �g�ь���)
    Call SpreadSetVal(va�����f����, row, 13, �d�q�}�l�[)
    Call SpreadSetVal(va�����f����, row, 14, ���t�I�N)
    
    ' ����
    row = row + 1
    Call ���������擾("���ЃT�C�g", ��������, �N���W�b�g, �����N���W�b�g, ���i���, �R���r�j, ��s�U��, �y�V�o���N����, �y�C�W�[, �㕥, �|�C���g, �g�ь���, �d�q�}�l�[, ���t�I�N)
    
    Call SpreadSetVal(va�����f����, row, 2, ��������)
    Call SpreadSetVal(va�����f����, row, 3, �N���W�b�g)
    Call SpreadSetVal(va�����f����, row, 4, �����N���W�b�g)
    Call SpreadSetVal(va�����f����, row, 5, ���i���)
    Call SpreadSetVal(va�����f����, row, 6, �R���r�j)
    Call SpreadSetVal(va�����f����, row, 7, ��s�U��)
    Call SpreadSetVal(va�����f����, row, 8, �y�V�o���N����)
    Call SpreadSetVal(va�����f����, row, 9, �y�C�W�[)
    Call SpreadSetVal(va�����f����, row, 10, �㕥)
    Call SpreadSetVal(va�����f����, row, 11, �|�C���g)
    Call SpreadSetVal(va�����f����, row, 12, �g�ь���)
    Call SpreadSetVal(va�����f����, row, 13, �d�q�}�l�[)
    Call SpreadSetVal(va�����f����, row, 14, ���t�I�N)
        
    ' �A�}�]��
    row = row + 1
    Call ���������擾("�A�}�]��", ��������, �N���W�b�g, �����N���W�b�g, ���i���, �R���r�j, ��s�U��, �y�V�o���N����, �y�C�W�[, �㕥, �|�C���g, �g�ь���, �d�q�}�l�[, ���t�I�N)
    
    Call SpreadSetVal(va�����f����, row, 2, ��������)
    Call SpreadSetVal(va�����f����, row, 3, �N���W�b�g)
    Call SpreadSetVal(va�����f����, row, 4, �����N���W�b�g)
    Call SpreadSetVal(va�����f����, row, 5, ���i���)
    Call SpreadSetVal(va�����f����, row, 6, �R���r�j)
    Call SpreadSetVal(va�����f����, row, 7, ��s�U��)
    Call SpreadSetVal(va�����f����, row, 8, �y�V�o���N����)
    Call SpreadSetVal(va�����f����, row, 9, �y�C�W�[)
    Call SpreadSetVal(va�����f����, row, 10, �㕥)
    Call SpreadSetVal(va�����f����, row, 11, �|�C���g)
    Call SpreadSetVal(va�����f����, row, 12, �g�ь���)
    Call SpreadSetVal(va�����f����, row, 13, �d�q�}�l�[)
    Call SpreadSetVal(va�����f����, row, 14, ���t�I�N)
    
    ' �����g���b�N�X
    row = row + 1
    Call ���������擾("�����g���b�N�X", ��������, �N���W�b�g, �����N���W�b�g, ���i���, �R���r�j, ��s�U��, �y�V�o���N����, �y�C�W�[, �㕥, �|�C���g, �g�ь���, �d�q�}�l�[, ���t�I�N)
'    Call ���������擾("������̂��l�b�g", ��������, �N���W�b�g, ���i���, �R���r�j, ��s�U��, �y�V�o���N����, �y�C�W�[, �㕥, �|�C���g, �g�ь���)
    
    Call SpreadSetVal(va�����f����, row, 2, ��������)
    Call SpreadSetVal(va�����f����, row, 3, �N���W�b�g)
    Call SpreadSetVal(va�����f����, row, 4, �����N���W�b�g)
    Call SpreadSetVal(va�����f����, row, 5, ���i���)
    Call SpreadSetVal(va�����f����, row, 6, �R���r�j)
    Call SpreadSetVal(va�����f����, row, 7, ��s�U��)
    Call SpreadSetVal(va�����f����, row, 8, �y�V�o���N����)
    Call SpreadSetVal(va�����f����, row, 9, �y�C�W�[)
    Call SpreadSetVal(va�����f����, row, 10, �㕥)
    Call SpreadSetVal(va�����f����, row, 11, �|�C���g)
    Call SpreadSetVal(va�����f����, row, 12, �g�ь���)
    Call SpreadSetVal(va�����f����, row, 13, �d�q�}�l�[)
    Call SpreadSetVal(va�����f����, row, 14, ���t�I�N)
    
    ' ���t�I�N
    row = row + 1
    Call ���������擾("���t�I�N", ��������, �N���W�b�g, �����N���W�b�g, ���i���, �R���r�j, ��s�U��, �y�V�o���N����, �y�C�W�[, �㕥, �|�C���g, �g�ь���, �d�q�}�l�[, ���t�I�N)
    
    Call SpreadSetVal(va�����f����, row, 2, ��������)
    Call SpreadSetVal(va�����f����, row, 3, �N���W�b�g)
    Call SpreadSetVal(va�����f����, row, 4, �����N���W�b�g)
    Call SpreadSetVal(va�����f����, row, 5, ���i���)
    Call SpreadSetVal(va�����f����, row, 6, �R���r�j)
    Call SpreadSetVal(va�����f����, row, 7, ��s�U��)
    Call SpreadSetVal(va�����f����, row, 8, �y�V�o���N����)
    Call SpreadSetVal(va�����f����, row, 9, �y�C�W�[)
    Call SpreadSetVal(va�����f����, row, 10, �㕥)
    Call SpreadSetVal(va�����f����, row, 11, �|�C���g)
    Call SpreadSetVal(va�����f����, row, 12, �g�ь���)
    Call SpreadSetVal(va�����f����, row, 13, �d�q�}�l�[)
    Call SpreadSetVal(va�����f����, row, 14, ���t�I�N)

End Sub

'************************************************************************
'�@  �\ :�X�֔ԍ�����
'************************************************************************
Private Sub cmd�X�֔ԍ�_Click()
    
    Dim ADF025      As New ADF025
    Dim i           As Integer
    Dim �X�֔ԍ�    As String
    
    Call ADF025.Show(1)
    
    �X�֔ԍ� = ADF025.get�X�֔ԍ�()
    If �X�֔ԍ� <> "" Then
        txt�X�֔ԍ�.Text = �X�֔ԍ�
    End If
    
End Sub

'************************************************************************
'�@  �\ :�⍇�ԍ��Z�b�g
'************************************************************************
Private Sub cmd�⍇�ԍ�_Click()

    Dim cCsvReader  As CsvReader
    Set cCsvReader = New CsvReader
    Dim �ڋqID      As String
    Dim �o�ד�      As String
    Dim �⍇�ԍ�    As String
    Dim �폜�敪    As String
    
    MousePointer = vbHourglass
    
    ' �w�肵�� CSV �t�@�C�����J��
    If cCsvReader.OpenStream("c:\�ڋq�Ǘ�\�z������.csv") = False Then
        MousePointer = vbNormal
        Call MsgBox("�z������������܂���B", vbOK, "�ڋq�Ǘ�")
        Exit Sub
    End If
    
    ' �ŏ��̍s���w�b�_�Ƃ��ēǂݍ���
    Call cCsvReader.ReadHeader

    ' CSV �t�@�C���̒��g�����ׂĎ擾����
    Dim cTable As Collection
    Set cTable = cCsvReader.ReadToEnd()

    ' ���ׂĂ̒��g (Table) ���� �s (Row) ��񋓂��Ď��o��
    Dim cRow As Collection
    
    For Each cRow In cTable
        ' �s����J���������g���Ċe Item ���o�͂���
        On Error GoTo skip
        �ڋqID = cRow("�Z���^�R�[�h")
        If �ڋqID = "" Then
            Exit For
        End If
        �ڋqID = Format(CLng(�ڋqID), "00000")
        �o�ד� = cRow("�o�ד���")
        �⍇�ԍ� = �⍇�ԍ��ҏW(cRow("���⍇�������"))
        �폜�敪 = cRow("�폜�敪")
        If �폜�敪 = "0" Then
            Call �₢���킹�ԍ��X�V(�ڋqID, �⍇�ԍ�)
        End If
    Next
    
skip:
   
    Call MsgBox("�⍇�ԍ���ǂݍ��݂܂����B", vbOKOnly, "�ڋq�Ǘ�")
    
    MousePointer = vbNormal

End Sub

'************************************************************************
'�@  �\ :�y�V��Yahoo�ڍs
'************************************************************************
Private Sub cmd�ڍs_Click()
    
    Dim �ڋqID      As String
    Dim �ڋq��      As String
    Dim �ڋq�}�X�^RS As New ADODB.Recordset
    Dim ADF027      As New ADF027

    Call ADF027.Show(1)
    
    Call ADF027.�ڋqID�擾(�ڋqID, �ڋq��)
    
    cmb��������.Text = "�ڋq��"
    txt��������.Text = �ڋq��
    
    MousePointer = vbHourglass
    
    Call �ڋq����7(�ڋqID, �ڋq�}�X�^RS)
    
    Call �ڋq���X�g�\��(�ڋq�}�X�^RS)
    
    MousePointer = vbNormal
    
End Sub

