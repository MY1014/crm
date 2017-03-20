VERSION 5.00
Object = "{A4B55B03-8129-101D-836D-3E0683BCA07A}#1.0#0"; "TEXT50S.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "SS32X25.OCX"
Object = "{604A59D5-2409-101D-97D5-C6626B63EF2D}#1.0#0"; "NUM50S.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{FE1D09E3-6FC7-101D-836D-3E0683BCA07A}#1.0#0"; "DATE50S.OCX"
Begin VB.Form ADF010 
   Caption         =   "キャットハンド顧客管理"
   ClientHeight    =   13530
   ClientLeft      =   2700
   ClientTop       =   915
   ClientWidth     =   15900
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
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
   Begin FPSpread.vaSpread va注文リスト 
      Height          =   4215
      Left            =   9000
      OleObjectBlob   =   "ADF010.frx":0000
      TabIndex        =   22
      Top             =   1560
      Width           =   6495
   End
   Begin FPSpread.vaSpread va顧客リスト 
      Height          =   4215
      Left            =   240
      OleObjectBlob   =   "ADF010.frx":360A
      TabIndex        =   21
      Top             =   1560
      Width           =   8775
   End
   Begin FPSpread.vaSpread va注文掲示板 
      Height          =   1095
      Left            =   9000
      OleObjectBlob   =   "ADF010.frx":4B2E
      TabIndex        =   126
      Top             =   120
      Width           =   6375
   End
   Begin VB.CommandButton cmd移行 
      Caption         =   "移行"
      Height          =   375
      Left            =   5400
      TabIndex        =   129
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmd問合番号 
      Caption         =   "問合番号"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.CommandButton cmd更新 
      Caption         =   "更新"
      Height          =   375
      Left            =   14400
      TabIndex        =   127
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdTOOL 
      Caption         =   "TOOL"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.CommandButton cmd個別メール 
      Caption         =   "個別メール"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.CommandButton cmd閉じる 
      Caption         =   "閉じる"
      CausesValidation=   0   'False
      Height          =   495
      Left            =   14760
      TabIndex        =   119
      Top             =   12840
      Width           =   1095
   End
   Begin VB.CommandButton cmd一括完了 
      Caption         =   "一括完了"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.CommandButton cmd売上 
      Caption         =   "売上出力"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.CommandButton cmdオートシップ 
      Cancel          =   -1  'True
      Caption         =   "オートシップ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.CommandButton cmdコモライフ 
      Caption         =   "コモライフ"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
   Begin VB.Frame frmメール履歴 
      Height          =   5535
      Left            =   600
      TabIndex        =   104
      Top             =   7080
      Width           =   14775
      Begin FPSpread.vaSpread vaメール履歴 
         Height          =   4935
         Left            =   480
         OleObjectBlob   =   "ADF010.frx":519C
         TabIndex        =   105
         Top             =   240
         Width           =   8895
      End
      Begin ImTextCtrl.ImText txtメール本文 
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
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.CommandButton cmdアーデル購入者 
      Caption         =   "アーデル購入者"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.CommandButton cmdメルマガ発行 
      Caption         =   "メルマガ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.Frame frm注文 
      Height          =   5535
      Left            =   600
      TabIndex        =   58
      Top             =   7080
      Width           =   14775
      Begin VB.CheckBox chk資料1_1 
         Caption         =   "効果的利用"
         Height          =   255
         Left            =   7080
         TabIndex        =   125
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CheckBox chk資料2_1 
         Caption         =   "毎日の積重"
         Height          =   255
         Left            =   8640
         TabIndex        =   124
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CheckBox chk資料3_1 
         Caption         =   "ＤＶＤ"
         Height          =   255
         Left            =   10320
         TabIndex        =   123
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CheckBox chk資料4_1 
         Caption         =   "運動と育毛"
         Height          =   255
         Left            =   11400
         TabIndex        =   122
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CheckBox chk資料5_1 
         Caption         =   "秘密の育毛"
         Height          =   255
         Left            =   13080
         TabIndex        =   121
         Top             =   5160
         Width           =   1455
      End
      Begin VB.ComboBox cmb部門 
         Height          =   345
         IMEMode         =   4  '全角ひらがな
         Left            =   4560
         TabIndex        =   66
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox cmb銀行 
         Height          =   345
         IMEMode         =   4  '全角ひらがな
         Left            =   4560
         TabIndex        =   70
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmd計算 
         Caption         =   "コモライフ計算"
         Height          =   375
         Left            =   6840
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CommandButton cmdテンプレート 
         Caption         =   "検索"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin VB.ComboBox cmbテンプレート 
         Height          =   345
         Left            =   8880
         TabIndex        =   101
         Top             =   2280
         Width           =   4575
      End
      Begin ImTextCtrl.ImText txt決済URL 
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
         Caption         =   "決済ID"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   72
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin VB.CommandButton cmd割引5 
         Caption         =   "定価"
         Height          =   375
         Left            =   12840
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   1320
         Width           =   735
      End
      Begin ImDateCtrl.ImDate txt出荷予定日 
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
         Caption         =   "出荷予定日"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   88
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin VB.CommandButton cmd割引4 
         Caption         =   "-20%"
         Height          =   375
         Left            =   11880
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmd割引3 
         Caption         =   "-10%"
         Height          =   375
         Left            =   10920
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmd割引2 
         Caption         =   "-4600"
         Height          =   375
         Left            =   9840
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox txt合計金額 
         Alignment       =   1  '右揃え
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
      Begin VB.CommandButton cmd割引 
         Caption         =   "-770"
         Height          =   375
         Left            =   8880
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmd本日2 
         Caption         =   "本日"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin VB.CommandButton cmd本日1 
         Caption         =   "本日"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImDateCtrl.ImDate txt入金日 
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
         Caption         =   "入金日"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin VB.ComboBox cmb注文元 
         Height          =   345
         IMEMode         =   4  '全角ひらがな
         ItemData        =   "ADF010.frx":5507
         Left            =   3240
         List            =   "ADF010.frx":550E
         TabIndex        =   63
         Top             =   840
         Width           =   2535
      End
      Begin ImTextCtrl.ImText txt注文ID 
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
         Caption         =   "注文ID"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin VB.ComboBox cmb宅配業者 
         Height          =   345
         IMEMode         =   4  '全角ひらがな
         Left            =   3120
         TabIndex        =   75
         Top             =   2760
         Width           =   1935
      End
      Begin ImTextCtrl.ImText txtメール送信 
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
         Caption         =   "メール送信"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   80
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImTextCtrl.ImText txt備考2 
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
         Caption         =   "備考"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   0
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImTextCtrl.ImText txt問合番号 
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
         Caption         =   "問合番号"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   72
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImTextCtrl.ImText txt支払番号 
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
         Caption         =   "支払番号"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   72
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImDateCtrl.ImDate txt出荷日 
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
         Caption         =   "出荷日"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImNumberCtrl.ImNumber txtその他手数料 
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
         Caption         =   "ポイント利用"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   88
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImNumberCtrl.ImNumber txt返金 
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
         Caption         =   "返金"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   40
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImNumberCtrl.ImNumber txt送料 
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
         Caption         =   "送料"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   40
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImNumberCtrl.ImNumber txt数量 
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
         Caption         =   "数量"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   40
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImNumberCtrl.ImNumber txt割引 
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
         Caption         =   "割引"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   40
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImNumberCtrl.ImNumber txt単価 
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
         Caption         =   "単価"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   40
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImTextCtrl.ImText txt配達日時 
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
         Caption         =   "配達日時"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   72
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin VB.ComboBox cmb注文方法 
         Height          =   345
         IMEMode         =   4  '全角ひらがな
         Left            =   1320
         TabIndex        =   68
         Top             =   1800
         Width           =   2535
      End
      Begin VB.ComboBox cmb商品名 
         Height          =   345
         IMEMode         =   4  '全角ひらがな
         Left            =   1320
         TabIndex        =   65
         Top             =   1320
         Width           =   3255
      End
      Begin VB.ComboBox cmbステータス 
         Height          =   345
         IMEMode         =   4  '全角ひらがな
         ItemData        =   "ADF010.frx":5684
         Left            =   1320
         List            =   "ADF010.frx":568B
         TabIndex        =   62
         Top             =   840
         Width           =   1815
      End
      Begin ImDateCtrl.ImDate txt受注日 
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
         Caption         =   "受注日"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImTextCtrl.ImText txt注文番号 
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
         Caption         =   "注文番号"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   64
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImTextCtrl.ImText txtコモライフ 
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
         Caption         =   "ｺﾓﾗｲﾌNO"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   72
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImNumberCtrl.ImNumber txt荷造運賃 
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
         Caption         =   "荷造運賃"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   64
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImNumberCtrl.ImNumber txt仕入金額 
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
         Caption         =   "仕入金額"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   72
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImTextCtrl.ImText txt配達日時2 
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
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin VB.Label lbl銀行 
         Caption         =   "銀行"
         Height          =   375
         Left            =   3960
         TabIndex        =   69
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lb合計金額 
         Caption         =   "合計金額"
         Height          =   375
         Left            =   5550
         TabIndex        =   95
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label lb注文方法 
         Caption         =   "注文方法"
         Height          =   255
         Left            =   240
         TabIndex        =   67
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lb商品名 
         Caption         =   "商品名"
         Height          =   255
         Left            =   480
         TabIndex        =   64
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lbステータス 
         Caption         =   "ステータス"
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame frm顧客 
      Height          =   5415
      Left            =   600
      TabIndex        =   30
      Top             =   7200
      Width           =   14775
      Begin VB.CommandButton cmd郵便番号 
         Caption         =   "〒"
         Height          =   375
         Left            =   2880
         TabIndex        =   0
         Top             =   1320
         Width           =   735
      End
      Begin VB.CheckBox chk資料5 
         Caption         =   "秘密の育毛"
         Height          =   255
         Left            =   13320
         TabIndex        =   55
         Top             =   4800
         Width           =   1455
      End
      Begin VB.CheckBox chk資料4 
         Caption         =   "運動と育毛"
         Height          =   255
         Left            =   11640
         TabIndex        =   54
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CheckBox chk資料3 
         Caption         =   "ＤＶＤ"
         Height          =   255
         Left            =   10560
         TabIndex        =   53
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CheckBox chk資料2 
         Caption         =   "毎日の積重"
         Height          =   255
         Left            =   8880
         TabIndex        =   52
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CheckBox chk資料1 
         Caption         =   "効果的利用"
         Height          =   255
         Left            =   7320
         TabIndex        =   51
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton cmd転居 
         Caption         =   "転居"
         Height          =   375
         Left            =   5520
         TabIndex        =   46
         Top             =   4800
         Width           =   1095
      End
      Begin ImDateCtrl.ImDate txt誕生日 
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
         Caption         =   "誕生日"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   64
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImDateCtrl.ImDate txt退会日 
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
         Caption         =   "退会日"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImTextCtrl.ImText txt楽天メール 
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
         Caption         =   "楽天メール"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   88
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin VB.CheckBox chkメール送信 
         Height          =   255
         Left            =   1680
         TabIndex        =   44
         Top             =   4920
         Width           =   495
      End
      Begin VB.CommandButton cmd転記 
         Caption         =   "顧客情報転記"
         Height          =   495
         Left            =   2880
         TabIndex        =   32
         Top             =   240
         Width           =   1815
      End
      Begin VB.Frame frm男女 
         BorderStyle     =   0  'なし
         Height          =   495
         Left            =   4080
         TabIndex        =   120
         Top             =   1200
         Width           =   3375
         Begin VB.OptionButton opt女性 
            Caption         =   "女性"
            Height          =   345
            Left            =   1560
            TabIndex        =   37
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton opt男性 
            Caption         =   "男性"
            Height          =   375
            Left            =   480
            TabIndex        =   36
            Top             =   120
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin ImTextCtrl.ImText txt備考 
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
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImDateCtrl.ImDate txt入会日 
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
         Caption         =   "入会日"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   64
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin VB.ComboBox cmbアーデルクラブ 
         Height          =   345
         IMEMode         =   4  '全角ひらがな
         ItemData        =   "ADF010.frx":57C9
         Left            =   9840
         List            =   "ADF010.frx":57CB
         TabIndex        =   47
         Top             =   360
         Width           =   2175
      End
      Begin ImTextCtrl.ImText txtメール 
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
         Caption         =   "メール"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImTextCtrl.ImText txt電話番号 
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
         Caption         =   "電話番号"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   72
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImTextCtrl.ImText txt住所_下段 
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
         Caption         =   "住所_下段"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   80
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImTextCtrl.ImText txt住所_上段 
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
         Caption         =   "住所_上段"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   80
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImTextCtrl.ImText txt郵便番号 
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
         Caption         =   "〒"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   24
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImTextCtrl.ImText txtフリガナ 
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
         Caption         =   "フリガナ"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImTextCtrl.ImText txt顧客名 
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
         Caption         =   "顧客名"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImTextCtrl.ImText txt顧客ID 
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
         Caption         =   "顧客ID"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   56
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin ImTextCtrl.ImText txt住所_中段 
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
         Caption         =   "住所_中段"
         CaptionAlignment=   3
         CaptionColor    =   0
         CaptionWidth    =   80
         CaptionPosition =   0
         CaptionSpacing  =   3
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
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
            Name            =   "ＭＳ Ｐゴシック"
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
      Begin VB.Label lbメール送信 
         Caption         =   "メール送信"
         Height          =   375
         Left            =   480
         TabIndex        =   56
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label lbアーデルクラブ 
         Caption         =   "アーデルクラブ"
         Height          =   255
         Left            =   8160
         TabIndex        =   57
         Top             =   405
         Width           =   1455
      End
   End
   Begin MSComctlLib.TabStrip tab情報 
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
            Caption         =   "顧客"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "配送先"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "注文"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "メール履歴"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdキャンセル検索 
      Caption         =   "ｷｬﾝｾﾙ検索"
      Height          =   375
      Left            =   10920
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd保留中検索 
      Caption         =   "保留中検索"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmd出荷済み 
      Caption         =   "出荷済み"
      Height          =   375
      Left            =   9120
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd出荷処理中 
      Caption         =   "出荷処理中"
      Height          =   375
      Left            =   11640
      TabIndex        =   15
      Top             =   5760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd入金待ち 
      Caption         =   "入金待ち"
      Height          =   375
      Left            =   7920
      TabIndex        =   13
      Top             =   6360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd新規注文 
      Caption         =   "新規注文"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdCSV出力 
      Caption         =   "e飛伝出力"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.CommandButton cmdアーデルクラブ 
      Caption         =   "クラブメール"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.CommandButton cmd出荷予定一覧 
      Caption         =   "出荷予定一覧"
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   5880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd未入金 
      Caption         =   "未入金検索"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox cmb検索条件 
      Height          =   345
      Left            =   240
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmd全解除 
      Caption         =   "全解除"
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmd全選択 
      Caption         =   "全選択"
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdクラブ未加入 
      Caption         =   "クラブ未加入"
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdクラブ検索 
      Caption         =   "クラブ検索"
      Height          =   375
      Left            =   9840
      TabIndex        =   19
      Top             =   6000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdモモ検索 
      Caption         =   "モモ検索"
      Height          =   375
      Left            =   11640
      TabIndex        =   20
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmd未出荷一覧 
      Caption         =   "未出荷検索"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdメール 
      Caption         =   "メール"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin ImNumberCtrl.ImNumber txt累積数 
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
         Name            =   "ＭＳ Ｐゴシック"
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
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.CommandButton cmd検索 
      Caption         =   "検索"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmd礼状 
      Caption         =   "お礼状"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.CommandButton cmd納品書 
      Caption         =   "納品書"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.CommandButton cmd削除2 
      Caption         =   "注文削除"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.CommandButton cmd追加2 
      Caption         =   "新規注文"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.CommandButton cmd削除1 
      Caption         =   "顧客削除"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.CommandButton cmd追加1 
      Caption         =   "新規顧客"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin ImTextCtrl.ImText txt検索条件 
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
         Name            =   "ＭＳ Ｐゴシック"
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
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.Label lbl注意喚起 
      Alignment       =   2  '中央揃え
      BackColor       =   &H000000FF&
      Caption         =   "注意喚起"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   130
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lbl矢印2 
      Caption         =   "→"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.Label txt注意喚起 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
   Begin VB.Label lb累積本数 
      Caption         =   "累積本数"
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

Private G_行番号        As Integer
Private G_フラグ        As Boolean
Private G_ROW           As Long
Private G_タブNO        As Integer
Public G_顧客リスト_ROW As Long
Public G_注文リスト_ROW As Long
Public G_注文元         As String
Public G_商品名         As String

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SendMail Lib "bsmtp" _
      (szServer As String, szTo As String, szFrom As String, _
      szSubject As String, szBody As String, szFile As String) As String
'************************
'オリジナル入力モード定数
'************************
'全角ひらがな入力
Private Const MY_IME_CHMODE_ZEN_HIRA = IME_CMODE_ROMAN Or IME_CMODE_JAPANESE Or IME_CMODE_FULLSHAPE
'全角カタカナ入力
Private Const MY_IME_CHMODE_ZEN_KATA = IME_CMODE_ROMAN Or IME_CMODE_JAPANESE Or IME_CMODE_KATAKANA Or IME_CMODE_FULLSHAPE
'全角英数入力
Private Const MY_IME_CHMODE_ZEN_EISU = IME_CMODE_ROMAN Or IME_CMODE_FULLSHAPE
'半角カタカナ入力
Private Const MY_IME_CHMODE_HAN_KATA = IME_CMODE_ROMAN Or IME_CMODE_JAPANESE Or IME_CMODE_KATAKANA Or IME_CMODE_LANGUAGE
'半角英数入力
Private Const MY_IME_CHMODE_HAN_EISU = IME_CMODE_ROMAN

'************************************************************************
'機  能 :フォームロード
'************************************************************************
Private Sub Form_Load()
    
    Dim i As Integer
    Dim 店舗マスタRS As New ADODB.Recordset
    
    Call コネクション

    If va顧客リスト.MaxRows >= 1 Then
        Call 注文表示(1)
    End If
        
    ' 店舗マスタをリードする
    G_消費税 = 0.08
    G_仕内 = "仕内8%"
    G_売内 = "売内8%"
    
    Call 店舗マスタ取得(店舗マスタRS)
    If Not 店舗マスタRS.EOF Then
        G_店舗名 = 店舗マスタRS!店舗名
        G_店舗略称 = 店舗マスタRS!店舗略称
        G_店舗色 = 店舗マスタRS!店舗色
        G_メール = 店舗マスタRS!メール
        G_サーバ = 店舗マスタRS!サーバ                                          ' mail.cathand.jp:587
        G_消費税 = CDbl(店舗マスタRS!消費税)                                   ' = 1.05
        G_仕内 = 店舗マスタRS!仕内                                              ' = "仕内5%"
        G_売内 = 店舗マスタRS!売内                                              ' = "売内5%"

        
        If 店舗マスタRS!送信元2 <> "" Then
            G_送信元 = 店舗マスタRS!送信元1 & vbTab & 店舗マスタRS!送信元2          ' info@cathand.jp & vbTab & info@cathand.jp:info
        Else
            G_送信元 = 店舗マスタRS!送信元1
        End If
        
        If G_店舗名 = "トリニティー楽天市場店" Then
            G_送信元 = G_送信元 & vbTab & "CRAM-MD5"
        End If
        G_ユーザ = 店舗マスタRS!ユーザ                                          ' order2@cathand.jp
        G_パスワード = 店舗マスタRS!パスワード                                  ' order2@cathand.jp
    End If
    
    店舗マスタRS.Close

    Call 楽天_店舗マスタ取得(店舗マスタRS)
    If Not 店舗マスタRS.EOF Then
        G_サーバ2 = 店舗マスタRS!サーバ                                         ' sub.fw.rakuten.ne.jp:587
        
        If 店舗マスタRS!送信元2 <> "" Then
            G_送信元2 = 店舗マスタRS!送信元1 & vbTab & 店舗マスタRS!送信元2     ' 251377:IwK93MZNj0
        Else
            G_送信元2 = 店舗マスタRS!送信元1
        End If
        
        G_メール2 = 店舗マスタRS!メール
        G_送信元2 = G_送信元2 & vbTab & "CRAM-MD5"
    End If
    
    店舗マスタRS.Close


    Call cmb検索条件.Clear
    Call cmb検索条件.AddItem("顧客名")
    Call cmb検索条件.AddItem("お届け先名")
    Call cmb検索条件.AddItem("フリガナ")
    Call cmb検索条件.AddItem("電話番号")
    Call cmb検索条件.AddItem("メール")
    Call cmb検索条件.AddItem("楽天メール")
    Call cmb検索条件.AddItem("〒")
    Call cmb検索条件.AddItem("住所1")
    Call cmb検索条件.AddItem("住所2")
    Call cmb検索条件.AddItem("住所3")
    Call cmb検索条件.AddItem("注文番号")
    Call cmb検索条件.AddItem("問合番号")
    Call cmb検索条件.AddItem("コモライフNO")
    Call cmb検索条件.AddItem("決済ID")
    Call cmb検索条件.AddItem("出荷日")
    cmb検索条件.ListIndex = 0
    
    '住所2の列を非表示にする
    va顧客リスト.Col = COL_住所2
    va顧客リスト.ColHidden = True
    
    '住所3の列を非表示にする
    va顧客リスト.Col = COL_住所3
    va顧客リスト.ColHidden = True
    
    'チェックボックスの列を非表示にする
    'va顧客リスト.col = COL_チェック
    'va顧客リスト.ColHidden = True
    
    '消費税の列を非表示にする
    va注文リスト.Col = COL_消費税
    va注文リスト.ColHidden = True
    
    '注文IDの列を非表示にする
    va注文リスト.Col = COL_注文ID
    va注文リスト.ColHidden = True
    
    '顧客IDの列を非表示にする
    va注文リスト.Col = COL_顧客ID2
    va注文リスト.ColHidden = True
    
    '顧客名の列を非表示にする
    va注文リスト.Col = COL_顧客名2
    va注文リスト.ColHidden = True
    
    ' 参照元を非表示にする
    va注文リスト.Col = COL_参照元
    va注文リスト.ColHidden = True
    
    ' キーワードを非表示にする
    va注文リスト.Col = COL_キーワード
    va注文リスト.ColHidden = True
    
    ' 入力ポイントを非表示にする
    va注文リスト.Col = COL_入力ポイント
    va注文リスト.ColHidden = True
    
    ' 送付資料を非表示にする
    va注文リスト.Col = COL_送付資料
    va注文リスト.ColHidden = True
    
    ' 返品対象を非表示にする
    va注文リスト.Col = COL_返品対象
    va注文リスト.ColHidden = True
    
    ' ロイヤリティーを非表示にする
    va注文リスト.Col = COL_ロイヤリティー
    va注文リスト.ColHidden = True
    
    For i = 2 To va顧客リスト.MaxCols
        va顧客リスト.Col = i
        va顧客リスト.row = -1
        va顧客リスト.Protect = True
        va顧客リスト.Lock = True
    Next i
    
    For i = 2 To va注文リスト.MaxCols
        va注文リスト.Col = i
        va注文リスト.row = -1
        va注文リスト.Protect = True
        va注文リスト.Lock = True
    Next i
    
    For i = 1 To vaメール履歴.MaxCols
        vaメール履歴.Col = i
        vaメール履歴.row = -1
        vaメール履歴.Protect = True
        vaメール履歴.Lock = True
    Next i
           
    ' アーデルクラブ
    Call cmbアーデルクラブ.Clear
    Call cmbアーデルクラブ.AddItem("")
    Call cmbアーデルクラブ.AddItem("アーデルクラブ")
'    Call cmbアーデルクラブ.AddItem("アーデル３ヶ月")
'    Call cmbアーデルクラブ.AddItem("アーデル６ヶ月")
'    Call cmbアーデルクラブ.AddItem("新ブスタ３ヶ月")
'    Call cmbアーデルクラブ.AddItem("新ブスタ６ヶ月")
'    Call cmbアーデルクラブ.AddItem("新ハイブリッター３ヶ月")
'    Call cmbアーデルクラブ.AddItem("新ハイブリッター６ヶ月")
'    Call cmbアーデルクラブ.AddItem("------------------------------")
'    Call cmbアーデルクラブ.AddItem("ブースター３ヶ月")
'    Call cmbアーデルクラブ.AddItem("ブースター６ヶ月")
'    Call cmbアーデルクラブ.AddItem("ハイブリッド３ヶ月")
'    Call cmbアーデルクラブ.AddItem("ハイブリッド６ヶ月")
    Call cmbアーデルクラブ.AddItem("なし")
    
    ' ステータス
    Call cmbステータス.Clear
    Call cmbステータス.AddItem("新規注文")
    Call cmbステータス.AddItem("処理中")
    Call cmbステータス.AddItem("入金処理")
    Call cmbステータス.AddItem("クレジット処理")
    Call cmbステータス.AddItem("出荷処理")
    Call cmbステータス.AddItem("出荷完了")
    Call cmbステータス.AddItem("キャンセル")
    Call cmbステータス.AddItem("コモライフ")
    Call cmbステータス.AddItem("保留")
    
    ' 商品名
    Call cmb商品名.Clear
    Call cmb商品名.AddItem("")
    Call cmb商品名.AddItem("アーデル")
    Call cmb商品名.AddItem("アーデル2本セット")
    Call cmb商品名.AddItem("アーデル＋シャンプー")
    
    Call cmb商品名.AddItem("------------------------------")
    Call cmb商品名.AddItem("新ブスタ")
    Call cmb商品名.AddItem("新ハイブリッター")
    Call cmb商品名.AddItem("新ブスタ＋シャンプー")
    Call cmb商品名.AddItem("新ハイブリッター＋シャンプー")
    Call cmb商品名.AddItem("------------------------------")
    Call cmb商品名.AddItem("ナイスレディー")
    Call cmb商品名.AddItem("ナイスレディー＋シャンプー")
    Call cmb商品名.AddItem("------------------------------")
    Call cmb商品名.AddItem("ブースター")
    Call cmb商品名.AddItem("ブースター（Ｗ発毛月間）")
    Call cmb商品名.AddItem("ハイブリッド")
    Call cmb商品名.AddItem("ブースター＋シャンプー")
    Call cmb商品名.AddItem("ハイブリッド＋シャンプー")
    
    Call cmb商品名.AddItem("------------------------------")
    Call cmb商品名.AddItem("シャンプー")
    Call cmb商品名.AddItem("シャンプー2本セット")
    Call cmb商品名.AddItem("シャンプー（プレゼント）")
    Call cmb商品名.AddItem("シャンプー＋トリートメント")
    Call cmb商品名.AddItem("------------------------------")
    Call cmb商品名.AddItem("トリートメント")
    Call cmb商品名.AddItem("トリートメント（プレゼント）")
    Call cmb商品名.AddItem("------------------------------")
    Call cmb商品名.AddItem("ブスタ５０％OFF券")
    Call cmb商品名.AddItem("ハイブリッター５０％OFF券")
    
    Call cmb商品名.AddItem("------------------------------")
'    Call cmb商品名.AddItem("アーデル活用・マニュアル（プレゼント）")
'    Call cmb商品名.AddItem("毎日の積み重ねが大切です・マニュアル（プレゼント）")
'    Call cmb商品名.AddItem("ドクターアーデル・育毛ＤＶＤ（プレゼント）")
'    Call cmb商品名.AddItem("育毛と運動・マニュアル（プレゼント）")
'    Call cmb商品名.AddItem("育毛・発毛マニュアル（プレゼント）")
'    Call cmb商品名.AddItem("------------------------------")
    Call cmb商品名.AddItem("アーデル＆シャンプー試供品")
    Call cmb商品名.AddItem("アーデル試供品")
    Call cmb商品名.AddItem("シャンプー試供品")
    
'    Call cmb商品名.AddItem("------------------------------")
'    Call cmb商品名.AddItem("モイストリッチ クレンジング")
'    Call cmb商品名.AddItem("モイストリッチ ウォッシング")
'    Call cmb商品名.AddItem("モイストリッチ ローション")
'    Call cmb商品名.AddItem("モイストリッチ ジェル")
'    Call cmb商品名.AddItem("モイストリッチ ロイヤルエッセンス")
'    Call cmb商品名.AddItem("モイストリッチ 基礎化粧品セット")
    
    Call cmb商品名.AddItem("------------------------------")
    Call cmb商品名.AddItem("アーデル資料")
    Call cmb商品名.AddItem("ミニまぐ")
    
    ' 部門
    Call cmb部門.Clear
    Call cmb部門.AddItem("ｱｰﾃﾞﾙ")
    Call cmb部門.AddItem("ｺﾓﾗｲﾌ")
    Call cmb部門.AddItem("その他")
    
    ' 注文方法
    Call cmb注文方法.Clear
    Call cmb注文方法.AddItem("")
    Call cmb注文方法.AddItem("クレジット")
    Call cmb注文方法.AddItem("東京クレジット")
    Call cmb注文方法.AddItem("商品代引")
    Call cmb注文方法.AddItem("コンビニ")
    Call cmb注文方法.AddItem("銀行振込")
    Call cmb注文方法.AddItem("楽天バンク決済")
    Call cmb注文方法.AddItem("ペイジー")
    Call cmb注文方法.AddItem("後払い")
    Call cmb注文方法.AddItem("ポイント")
    Call cmb注文方法.AddItem("携帯決済")
    Call cmb注文方法.AddItem("電子マネー")
    Call cmb注文方法.AddItem("ヤフオク")
    Call cmb注文方法.AddItem("－")
    
    ' 銀行
    Call cmb銀行.Clear
    Call cmb銀行.AddItem("")
    Call cmb銀行.AddItem("みずほ")
    Call cmb銀行.AddItem("楽天銀行")
    Call cmb銀行.AddItem("郵便振替口座")
    
    ' 宅配業者
    Call cmb宅配業者.Clear
    Call cmb宅配業者.AddItem("佐川急便")
    Call cmb宅配業者.AddItem("クロネコヤマト")
    Call cmb宅配業者.AddItem("ゆうパック")
    Call cmb宅配業者.AddItem("レターパック")
    Call cmb宅配業者.AddItem("ペリカン")
    
    ' 注文元
    Call cmb注文元.Clear
    Call cmb注文元.AddItem("")
    Call cmb注文元.AddItem(G_店舗略称)
    Call cmb注文元.AddItem("自社サイト")
'    Call cmb注文元.AddItem("レントラックス")
'    Call cmb注文元.AddItem("おちゃのこネット")
    Call cmb注文元.AddItem("アマゾン")
'    Call cmb注文元.AddItem("コマチ")
    Call cmb注文元.AddItem("ヤフオク")
'    Call cmb注文元.AddItem("インフォトップ")
'    Call cmb注文元.AddItem("メール")
'    Call cmb注文元.AddItem("FAX")
'    Call cmb注文元.AddItem("電話")
'    Call cmb注文元.AddItem("野口さん")
    Call cmb注文元.AddItem("その他")
    If G_店舗名 <> "トリニティー楽天市場店" Then
        Call cmb注文元.AddItem("楽天")
    End If
    
    ' テンプレート
    Call cmbテンプレート.Clear
    Call cmbテンプレート.AddItem("")
    Call cmbテンプレート.AddItem("アーデル新規")
    Call cmbテンプレート.AddItem("-------------------------------")
    Call cmbテンプレート.AddItem("新ブスタ新規")
    Call cmbテンプレート.AddItem("新ハイブリッター新規")
    Call cmbテンプレート.AddItem("-------------------------------")
    Call cmbテンプレート.AddItem("ブースター新規")
    Call cmbテンプレート.AddItem("ハイブリッド新規")
    Call cmbテンプレート.AddItem("ナイスレディー新規")
    Call cmbテンプレート.AddItem("-------------------------------")
    Call cmbテンプレート.AddItem("シャンプー２本セット新規")
    Call cmbテンプレート.AddItem("シャンプー新規")
    Call cmbテンプレート.AddItem("-------------------------------")
    Call cmbテンプレート.AddItem("シャンプー＆トリートメント新規")
    Call cmbテンプレート.AddItem("トリートメント新規")
    Call cmbテンプレート.AddItem("-------------------------------")
    Call cmbテンプレート.AddItem("アーデル＆シャンプーセット新規")
    Call cmbテンプレート.AddItem("新ブスタ＆シャンプーセット新規")
    Call cmbテンプレート.AddItem("新ハイブリッター＆シャンプーセット新規")
    Call cmbテンプレート.AddItem("ブースター＆シャンプーセット新規")
    Call cmbテンプレート.AddItem("ハイブリット＆シャンプーセット新規")
    Call cmbテンプレート.AddItem("ナイスレディー＆シャンプーセット新規")
    Call cmbテンプレート.AddItem("試供品新規")
    Call cmbテンプレート.AddItem("-------------------------------")
    Call cmbテンプレート.AddItem("アーデル活用")
    Call cmbテンプレート.AddItem("毎日の積み重ねが大切です")
    Call cmbテンプレート.AddItem("育毛ＤＶＤ")
    Call cmbテンプレート.AddItem("育毛と運動")
    Call cmbテンプレート.AddItem("育毛・発毛")

    G_行番号 = 0
    G_タブNO = 1
    G_注文元 = ""
    G_商品名 = ""
    G_顧客マスタ_排他フラグ = False
    
    cmb注文元.BackColor = vbRed

#If 0 Then

    If G_店舗名 = "トリニティー楽天市場店" Then
        txt楽天メール.Enabled = True
    Else
        txt楽天メール.Enabled = False
    End If
    
#End If
    
    If G_店舗名 = "トリニティー楽天市場店" Then
        cmd移行.Visible = False
    End If
    
End Sub

'************************************************************************
'機  能 :顧客マスタを表示する
'************************************************************************
Private Sub Form_Activate()

    If G_フラグ = False Then
        
        Dim ADF016      As New ADF016
    
        Dim 顧客マスタRS As New ADODB.Recordset
        
        ADF010.Caption = G_店舗名
        ADF010.BackColor = Val(G_店舗色)
        
        Call MsgBox(G_店舗名 & "用の顧客管理です。間違えないように注意して下さい！", vbOKOnly, "顧客管理")
          
        Call cmd未出荷一覧_Click
        
        If ADF016.一括起票_件数確認() > 0 Then
            If MsgBox("オートシップで確定データがあります" & vbCr & vbLf & "確定しますか？", vbYesNo, "顧客管理") = vbYes Then
                If MsgBox("プリンタの準備はOKですか？", vbYesNo, "顧客管理") = vbYes Then
                    If ADF016.一括起票() > 0 Then
                        Call MsgBox("アーデルクラブの確定を行いました！", vbOKOnly, "顧客管理")
                    End If
                End If
            End If
        End If
        
#If 0 Then
        ' 顧客マスタを全件リードする
        Call 顧客マスタ読込(顧客マスタRS)
    
        ' 顧客リストを表示する
        Call 顧客リスト表示(顧客マスタRS)
        
        If va注文リスト.MaxRows > 0 Then
            
            ' 最終行の背景色を変更する
            Call va注文リスト_Click(1, va注文リスト.MaxRows)
            
            ' セルのフォーカスを最終行に設定する
            Call SpreadSetFocus(va注文リスト, va注文リスト.MaxRows, COL_ステータス)
            
        End If
#End If
        G_フラグ = True
        
    End If
    
    On Error Resume Next
    
    DoEvents
    
    Select Case G_タブNO
        Case 1
                txt顧客名.SetFocus
        Case 2
                txt顧客名.SetFocus
        Case 3
                txt受注日.SetFocus
    End Select

    Call SpreadSetVal(va注文掲示板, 1, 1, G_店舗名)

End Sub

'************************************************************************
'機  能 :UNLOAD
'************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    
    ' 確認メッセージを表示する
    If MsgBox("終了してよろしいですか？", vbYesNo, "顧客管理") <> vbYes Then
        Cancel = 1
        Exit Sub
    End If
    
    Cancel = 0
    End
    
End Sub

'************************************************************************
'機  能 :閉じるボタン
'************************************************************************
Private Sub cmd閉じる_Click()

    ' 確認メッセージを表示する
    If MsgBox("終了してよろしいですか？", vbYesNo, "顧客管理") <> vbYes Then Exit Sub

    End
    
End Sub

'************************************************************************
'機  能 :顧客リストを表示する
'************************************************************************
Private Sub 顧客リスト表示(ByRef 顧客マスタRS As ADODB.Recordset)

    Dim row As Integer
    Dim 住所 As String
    
    row = 1
    G_行番号 = 0
    va顧客リスト.ReDraw = False
    va顧客リスト.MaxRows = 0
    
    ' 検索した顧客リストを表示する
    With 顧客マスタRS
        Do Until .EOF
            va顧客リスト.MaxRows = row
            
            Call SpreadSetVal(va顧客リスト, row, COL_チェック, 0)
          
            If Not IsNull(!顧客ID) Then
                Call SpreadSetVal(va顧客リスト, row, COL_顧客ID, !顧客ID)
            End If
            
            If Not IsNull(!顧客名) Then
                Call SpreadSetVal(va顧客リスト, row, COL_顧客名, !顧客名)
            End If
            
            If Not IsNull(!フリガナ) Then
                Call SpreadSetVal(va顧客リスト, row, COL_フリガナ, !フリガナ)
            End If
            
            If Not IsNull(![〒]) Then
                Call SpreadSetVal(va顧客リスト, row, COL_〒, ![〒])
            End If
            
            住所 = ""
            If Not IsNull(!住所1) Then
'               Call SpreadSetVal(va顧客リスト, row, COL_住所1, !住所1)
                住所 = 住所 + !住所1
            End If
            
            If Not IsNull(!住所2) Then
'               Call SpreadSetVal(va顧客リスト, row, COL_住所2, !住所2)
                住所 = 住所 + !住所2
            End If
            
            If Not IsNull(!住所3) Then
'               Call SpreadSetVal(va顧客リスト, row, COL_住所3, !住所3)
                住所 = 住所 + !住所3
            End If
            
            Call SpreadSetVal(va顧客リスト, row, COL_住所1, 住所)
            
            If Not IsNull(!電話番号) Then
                Call SpreadSetVal(va顧客リスト, row, COL_電話番号, !電話番号)
            End If
            
            If Not IsNull(!メール) Then
                Call SpreadSetVal(va顧客リスト, row, COL_メール, !メール)
            End If
            
            If Not IsNull(!アーデルクラブ) Then
                Call SpreadSetVal(va顧客リスト, row, COL_アーデルクラブ, !アーデルクラブ)
            End If
            
            If Not IsNull(!入会日) Then
                Call SpreadSetVal(va顧客リスト, row, COL_入会日, !入会日)
            End If
            
            If Not IsNull(!備考) Then
                Call SpreadSetVal(va顧客リスト, row, COL_備考, !備考)
                txt注意喚起.Caption = !備考
            End If
            
            If Not IsNull(!お届け先名) Then
                Call SpreadSetVal(va顧客リスト, row, COL_お届け先名, !お届け先名)
            End If
            
            If Not IsNull(!お届け先メール) Then
                Call SpreadSetVal(va顧客リスト, row, COL_お届け先メール, !お届け先メール)
            End If
            
            If Not IsNull(!楽天メール) Then
                Call SpreadSetVal(va顧客リスト, row, COL_楽天メール, !楽天メール)
            End If

            Call .MoveNext
            row = row + 1
        Loop
    End With
    
    顧客マスタRS.Close
    
    va顧客リスト.ReDraw = True
    
    If va顧客リスト.MaxRows > 0 Then
        
        ' セルのフォーカスを最終行に設定する
        Call SpreadSetFocus(va顧客リスト, va顧客リスト.MaxRows, COL_顧客名)

        ' 最終行の背景色を変更する
        Call va顧客リスト_Click(1, va顧客リスト.MaxRows)
                
        ' 先頭行の注文データを表示する
        If va顧客リスト.MaxRows >= 1 Then
            Call 注文表示(va顧客リスト.MaxRows)
        End If
    Else
        va注文リスト.MaxRows = 0
    End If
End Sub

'************************************************************************
'機  能 :顧客リストに１行追加する。
'************************************************************************
Private Sub cmd追加1_Click()
    
    Call トランザクションデータの更新
    
    G_タブNO = 1
    tab情報.Tabs(G_タブNO).Selected = True
    
    ' 顧客リストに１行追加する
    va顧客リスト.MaxRows = va顧客リスト.MaxRows + 1
    
    ' セルのフォーカスを追加した行に設定する
    Call SpreadSetFocus(va顧客リスト, va顧客リスト.MaxRows, COL_顧客名)
        
    ' 追加した行の背景色を変更する
    Call va顧客リスト_Click(1, va顧客リスト.MaxRows)
    
    ' 注文リストを消去する
    va注文リスト.MaxRows = 0
    
    Call 顧客情報クリア
    Call 注文情報クリア
    
End Sub

'************************************************************************
'機  能 :顧客リストで選択されている行を削除する。
'************************************************************************
Private Sub cmd削除1_Click()
    
    Dim i As Integer
    Dim row As Integer
    Dim 件数 As Integer
    Dim 顧客ID As String
    
    Call トランザクションデータの更新

    ' 確認メッセージを表示する
    If MsgBox("削除してよろしいですか？", vbYesNo, "顧客管理") <> vbYes Then Exit Sub
    
    件数 = 0
    For i = 1 To va顧客リスト.MaxRows
        If SpreadGetVal(va顧客リスト, i, COL_チェック) = "1" Then
            顧客ID = SpreadGetVal(va顧客リスト, i, COL_顧客ID)
    
            Call 顧客マスタ削除(顧客ID)
        
        End If
    Next i
    
    ' 顧客リストを表示し直す
    If txt顧客名.Text <> "" Then
        Call cmd検索_Click
    Else
        G_フラグ = False
        Call Form_Activate
    End If
    
    ' 先頭行の注文データを表示する
    If va顧客リスト.MaxRows >= 1 Then
        Call 注文表示(va顧客リスト.MaxRows)
    End If
    
    Call MsgBox("顧客データを削除しました", vbOKOnly, "顧客管理")
    
End Sub

'************************************************************************
'機  能 :伝票消込
'************************************************************************
Private Sub cmdTOOL_Click()

    Dim ADF022      As New ADF022
            
    Call ADF022.Show(1)
    
End Sub

'************************************************************************
'機  能 :顧客リストの行が変わったら注文を表示し直す
'************************************************************************
Private Sub va顧客リスト_Click(ByVal Col As Long, ByVal row As Long)
    
    Dim 電話番号        As String
    Dim 顧客マスタRS    As New ADODB.Recordset
    
    If row < 1 Then Exit Sub
    
    va顧客リスト.ReDraw = False
    
    With va顧客リスト
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

    va顧客リスト.ReDraw = True
    
    G_顧客リスト_ROW = row
    
    Call Tab情報_Click
    
    txt注意喚起.Caption = SpreadGetVal(va顧客リスト, row, COL_備考)
    
    電話番号 = SpreadGetVal(va顧客リスト, row, COL_電話番号)
    Call 楽天_電話番号検索(顧客マスタRS, 電話番号)
    
    lbl注意喚起.Visible = False
    
    If G_店舗名 = "トリニティー楽天市場店" Then
        If 顧客マスタRS!件数 > 0 Then
            lbl注意喚起.Visible = True
            lbl注意喚起.Caption = "Yahoo顧客"
        Else
            lbl注意喚起.Visible = False
        End If
    Else
        If 顧客マスタRS!件数 > 0 Then
            lbl注意喚起.Visible = True
            lbl注意喚起.Caption = "楽天顧客"
        Else
            lbl注意喚起.Visible = False
        End If
    End If
    
    顧客マスタRS.Close
    
    Call 注文表示(row)
    
End Sub

'************************************************************************
'機  能 :注文を表示する
'************************************************************************
Private Sub 注文表示(ByVal 行 As Integer)

    Dim row             As Integer
    Dim 累積本数        As Integer
    Dim 売上明細RS      As New ADODB.Recordset
    Dim 顧客ID          As String
    Dim 配達希望日時    As String
    
    顧客ID = SpreadGetVal(va顧客リスト, 行, COL_顧客ID)
    
    ' 選択された顧客の注文データを取得する
    Call 注文検索(顧客ID, 売上明細RS)

    累積本数 = 0
    
    row = 1
    
    va注文リスト.ReDraw = False
    va注文リスト.MaxRows = 0
    
    ' 選択された顧客の注文データを表示する
    With 売上明細RS
        Do Until .EOF
            va注文リスト.MaxRows = row
            
            Call SpreadSetVal(va注文リスト, row, 1, 0)
          
            If Not IsNull(!受注日) Then
                Call SpreadSetVal(va注文リスト, row, COL_受注日, !受注日)
            End If
          
            If Not IsNull(!ステータス) Then
                Call SpreadSetVal(va注文リスト, row, COL_ステータス, !ステータス)
            End If
          
            If Not IsNull(!商品名) Then
                Call SpreadSetVal(va注文リスト, row, COL_商品名, !商品名)
            End If
          
            If Not IsNull(!注文方法) Then
                Call SpreadSetVal(va注文リスト, row, COL_注文方法, !注文方法)
            End If
            
            配達希望日時 = ""
            
            If Not IsNull(!配達希望日時) Then
                配達希望日時 = !配達希望日時
            End If
            
            If Not IsNull(!配達希望日時2) Then
                配達希望日時 = 配達希望日時 + " " + !配達希望日時2
            End If
          
            Call SpreadSetVal(va注文リスト, row, COL_配達希望日時, 配達希望日時)

            If Not IsNull(!単価) Then
                Call SpreadSetVal(va注文リスト, row, COL_単価, !単価)
            End If
          
            If Not IsNull(!割引) Then
                Call SpreadSetVal(va注文リスト, row, COL_割引, !割引)
            End If
          
            If Not IsNull(!数量) Then
                Call SpreadSetVal(va注文リスト, row, COL_数量, !数量)
            End If
          
            If Not IsNull(!金額) Then
                Call SpreadSetVal(va注文リスト, row, COL_金額, !金額)
            End If
          
            If Not IsNull(!消費税) Then
                Call SpreadSetVal(va注文リスト, row, COL_消費税, !消費税)
            End If
          
            If Not IsNull(!送料) Then
                Call SpreadSetVal(va注文リスト, row, COL_送料, !送料)
            End If
          
            If Not IsNull(!返金) Then
                Call SpreadSetVal(va注文リスト, row, COL_返金, !返金)
            End If
          
            If Not IsNull(!その他手数料) Then
                Call SpreadSetVal(va注文リスト, row, COL_その他手数料, !その他手数料)
            End If
          
            If Not IsNull(!合計金額) Then
                Call SpreadSetVal(va注文リスト, row, COL_合計金額, !合計金額)
            End If
          
            If Not IsNull(!入金日) Then
                Call SpreadSetVal(va注文リスト, row, COL_入金日, !入金日)
            End If
          
            If Not IsNull(!出荷日) Then
                Call SpreadSetVal(va注文リスト, row, COL_出荷日, !出荷日)
            End If
          
            If Not IsNull(!着荷日) Then
                Call SpreadSetVal(va注文リスト, row, COL_着荷日, !着荷日)
            End If
          
            If Not IsNull(!宅配業者) Then
                Call SpreadSetVal(va注文リスト, row, COL_宅配業者, !宅配業者)
            End If
            
            If Not IsNull(!注文元) Then
                Call SpreadSetVal(va注文リスト, row, COL_注文元, !注文元)
            End If
            
            If Not IsNull(!Yahoo注文番号) Then
                Call SpreadSetVal(va注文リスト, row, COL_Yahoo注文番号, !Yahoo注文番号)
            End If
            
            If Not IsNull(!参照元) Then
                Call SpreadSetVal(va注文リスト, row, COL_参照元, !参照元)
            End If
            
            If Not IsNull(!キーワード) Then
                Call SpreadSetVal(va注文リスト, row, COL_キーワード, !キーワード)
            End If
            
            If Not IsNull(!入力ポイント) Then
                Call SpreadSetVal(va注文リスト, row, COL_入力ポイント, !入力ポイント)
            End If
            
            If Not IsNull(!商品コード) Then
                Call SpreadSetVal(va注文リスト, row, COL_商品コード, !商品コード)
            End If
            
            If Not IsNull(!ロイヤリティー) Then
                Call SpreadSetVal(va注文リスト, row, COL_ロイヤリティー, !ロイヤリティー)
            End If
            
            If Not IsNull(!送付資料) Then
                Call SpreadSetVal(va注文リスト, row, COL_送付資料, !送付資料)
            End If
            
            If Not IsNull(!返品対象) Then
                Call SpreadSetVal(va注文リスト, row, COL_返品対象, !返品対象)
            End If
            
            If Not IsNull(!支払番号) Then
                Call SpreadSetVal(va注文リスト, row, COL_支払番号, !支払番号)
            End If
            
            If Not IsNull(!問合番号) Then
                Call SpreadSetVal(va注文リスト, row, COL_問合番号, !問合番号)
            End If
            
            If Not IsNull(!備考1) Then
                Call SpreadSetVal(va注文リスト, row, COL_備考1, !備考1)
            End If
            
            If Not IsNull(!備考2) Then
                Call SpreadSetVal(va注文リスト, row, COL_備考2, !備考2)
            End If
            
            If Not IsNull(!備考3) Then
                Call SpreadSetVal(va注文リスト, row, COL_備考3, !備考3)
            End If
            
            If Not IsNull(!注文ID) Then
                Call SpreadSetVal(va注文リスト, row, COL_注文ID, !注文ID)
            End If
            
            If Not IsNull(!顧客ID) Then
                Call SpreadSetVal(va注文リスト, row, COL_顧客ID2, !顧客ID)
            End If
            
            If Not IsNull(!顧客名) Then
                Call SpreadSetVal(va注文リスト, row, COL_顧客名2, !顧客名)
            End If
            
            If Not IsNull(!メール送信) Then
                Call SpreadSetVal(va注文リスト, row, COL_メール送信, !メール送信)
            End If

            If !割引区分 = "%" Then
                Call SpreadSetVal(va注文リスト, row, COL_割引区分, "％")
            Else
                Call SpreadSetVal(va注文リスト, row, COL_割引区分, "円")
            End If
            
            If Not IsNull(!出荷予定日) Then
                Call SpreadSetVal(va注文リスト, row, COL_出荷予定日, !出荷予定日)
            End If
            
            If Not IsNull(!決済URL) Then
                Call SpreadSetVal(va注文リスト, row, COL_決済URL, !決済URL)
            End If
            
            G_売上_更新日時 = IIf(IsNull(!更新日時), Now, !更新日時)
            
            Call .MoveNext
            row = row + 1
        Loop
    End With
    
    売上明細RS.Close
    
    If va注文リスト.MaxRows > 0 Then
        
        ' 最終行の背景色を変更する
        Call va注文リスト_Click(1, va注文リスト.MaxRows)
        
        ' セルのフォーカスを最終行に設定する
        'Call SpreadSetFocus(va注文リスト, va注文リスト.MaxRows, COL_ステータス)
    Else
        'Call 注文情報クリア
    End If
    
    va注文リスト.ReDraw = True
    
    ' 累積本数を表示する
    txt累積数.Text = 累積数計算()

End Sub

'************************************************************************
'機  能 :注文リストに１行追加する
'************************************************************************
Private Sub cmd追加2_Click()
    
    Call トランザクションデータの更新

    G_タブNO = 3
    tab情報.Tabs(G_タブNO).Selected = True
    
    If txt顧客ID.Text = "" Then
        Call MsgBox("先ず顧客データを登録して下さい", vbOKOnly, "顧客管理")
        G_タブNO = 1
        tab情報.Tabs(G_タブNO).Selected = True
        Exit Sub
    End If
    
    va注文リスト.MaxRows = va注文リスト.MaxRows + 1
    
    txt受注日.Text = Format(Now, "YYYY/MM/DD")
    cmbステータス.Text = "新規注文"
    cmb商品名.Text = ""
    cmb部門.Text = "ｱｰﾃﾞﾙ"
    cmb注文方法.Text = "クレジット"
    cmb銀行.Text = ""
    txt単価.Value = 0
    txt割引.Value = 0
    txt数量.Value = 1
    txt送料.Value = 0
    txt返金.Value = 0
    txtその他手数料.Value = 0
    txt合計金額.Text = 0
    txt注文ID.Text = "-1"
    txtコモライフ.Text = ""
            
    txt配達日時.Text = ""
    txt出荷日.Text = "____/__/__"
    txt支払番号.Text = ""
    txt問合番号.Text = ""
    txt入金日.Text = "____/__/__"
    txtメール送信.Text = ""
    cmb注文元.Text = ""
    txt注文番号.Text = ""
    txt備考2.Text = ""
    txt出荷予定日.Text = "____/__/__"
    txt決済URL.Text = ""
    
    G_売上_更新日時 = Now
    G_注文元 = ""
    
    cmb宅配業者.Text = "佐川急便"
    G_商品名 = ""


    ' セルのフォーカスを追加した行に設定する
    Call SpreadSetFocus(va注文リスト, va注文リスト.MaxRows, COL_ステータス)

    ' 背景色を変更する
    Call va注文リスト_Click(1, va注文リスト.MaxRows)
        
    ' 売上明細を出力する
    Call 注文_更新
    
End Sub

'************************************************************************
'機  能 :注文リストで選択されている行を削除する。
'************************************************************************
Private Sub cmd削除2_Click()
    
    Dim i As Integer
    Dim 件数 As Integer
    Dim row As Integer
    Dim 注文ID As String
    
    Call トランザクションデータの更新

    ' 確認メッセージを表示する
    If MsgBox("削除してよろしいですか？", vbYesNo, "顧客管理") <> vbYes Then Exit Sub
    
    件数 = 0
    For i = 1 To va注文リスト.MaxRows
        If SpreadGetVal(va注文リスト, i, COL_チェック) = "1" Then
            注文ID = SpreadGetVal(va注文リスト, i, COL_注文ID)
    
            If 注文ID <> "" Then
                Call 売上明細削除(注文ID)
                件数 = 件数 + 1
            End If
        End If
    Next i
    
    ' 注文データを表示し直す
    Call 注文表示(G_顧客リスト_ROW)
    
    If 件数 > 0 Then
        Call MsgBox("注文データを削除しました", vbOKOnly, "顧客管理")
    End If
    
End Sub

'************************************************************************
'機  能 :注文リストの行が変わったら注文を表示し直す
'************************************************************************
Private Sub va注文リスト_Click(ByVal Col As Long, ByVal row As Long)
    
    If row < 1 Then Exit Sub
    
    va注文リスト.ReDraw = False
    
    With va注文リスト
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
    
    G_注文リスト_ROW = row
        
    va注文リスト.ReDraw = True
    
    If SpreadGetVal(va注文リスト, G_注文リスト_ROW, COL_注文方法) = "コンビニ" Then
        txt決済URL.Caption = "決済URL"
    Else
        txt決済URL.Caption = "決済ID"
    End If
    
    Call Tab情報_Click
    
End Sub

'************************************************************************
'機  能 :メール履歴の本文を表示する。
'************************************************************************
Private Sub vaメール履歴_Click(ByVal Col As Long, ByVal row As Long)
    
    Dim メール履歴RS As New ADODB.Recordset
    Dim 注文ID      As String
    Dim 送信日時    As String
    
    If row < 1 Then Exit Sub
    
    vaメール履歴.ReDraw = False
    
    With vaメール履歴
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
    
    vaメール履歴.ReDraw = True
    
    送信日時 = SpreadGetVal(vaメール履歴, row, 1)
    
    注文ID = SpreadGetVal(va注文リスト, G_注文リスト_ROW, COL_注文ID)
    
    If 注文ID = "" Or 注文ID = "注文ID" Then Exit Sub
    
    ' 選択されたメール履歴を取得する
    Call メール履歴検索2(注文ID, 送信日時, メール履歴RS)
    
    If Not メール履歴RS.EOF Then
        txtメール本文.Text = メール履歴RS!メール本文
    End If
    
    メール履歴RS.Close
    
End Sub

'************************************************************************
'機  能　注文詳細を表示する。
'************************************************************************
Private Sub Tab情報_Click()
    
    G_タブNO = Me.tab情報.SelectedItem.Index
    
   
    Select Case G_タブNO
        ' 顧客情報タブ
        Case 1
            frm顧客.Visible = True
            frm注文.Visible = False
            frmメール履歴.Visible = False
            
            txt顧客ID.Visible = True
            txt顧客ID.Enabled = False
            txt楽天メール.Visible = True
            lbアーデルクラブ.Visible = True
            cmbアーデルクラブ.Visible = True
            txt入会日.Visible = True
            txt退会日.Visible = True
            txt誕生日.Visible = True
            cmd転記.Visible = False
            chkメール送信.Visible = True
            lbメール送信.Visible = True
            cmd転居.Visible = True
            chk資料1.Visible = True
            chk資料2.Visible = True
            chk資料3.Visible = True
            chk資料4.Visible = True
            chk資料5.Visible = True
            Call 顧客タブ_表示
        ' 配達先タブ
        Case 2
            If txt顧客ID.Text = "" Then
                Call MsgBox("顧客情報が未入力です", vbOKOnly, "顧客管理")
                G_タブNO = 1
                tab情報.Tabs(G_タブNO).Selected = True
                Exit Sub
            End If
            
            frm顧客.Visible = True
            frm注文.Visible = False
            frmメール履歴.Visible = False
            
            txt顧客ID.Visible = False
            txt楽天メール.Visible = False
            lbアーデルクラブ.Visible = False
            cmbアーデルクラブ.Visible = False
            txt入会日.Visible = False
            txt退会日.Visible = False
            txt誕生日.Visible = False
            cmd転記.Visible = True
            chkメール送信.Visible = False
            lbメール送信.Visible = False
            cmd転居.Visible = False
            chk資料1.Visible = False
            chk資料2.Visible = False
            chk資料3.Visible = False
            chk資料4.Visible = False
            chk資料5.Visible = False
            Call 顧客タブ_表示
            
        ' 注文タブ
        Case 3
            If txt顧客ID.Text = "" Then
                Call MsgBox("顧客情報が未入力です", vbOKOnly, "顧客管理")
                G_タブNO = 1
                tab情報.Tabs(G_タブNO).Selected = True
                Exit Sub
            End If
                    
            frm顧客.Visible = False
            frm注文.Visible = True
            frmメール履歴.Visible = False
            
            Call 注文タブ_表示
            
        ' メール履歴タブ
        Case 4
            If txt顧客ID.Text = "" Then
                Call MsgBox("顧客情報が未入力です", vbOKOnly, "顧客管理")
                G_タブNO = 1
                tab情報.Tabs(G_タブNO).Selected = True
                Exit Sub
            End If
            
            frm顧客.Visible = False
            frm注文.Visible = False
            frmメール履歴.Visible = True
            
            Call メール履歴_表示
            
    End Select
    
    On Error Resume Next
    
    'DoEvents
    
    Select Case G_タブNO
        Case 1
                txt顧客名.SetFocus
        Case 2
                txt顧客名.SetFocus
        Case 3
                txt受注日.SetFocus
        Case 4
                vaメール履歴.SetFocus
    End Select
    
End Sub

'************************************************************************
'機  能 :顧客タブ情報を表示する
'************************************************************************
Private Sub 顧客タブ_表示()

    Dim 顧客ID As String
    Dim 顧客マスタRS As New ADODB.Recordset
  
    Dim 売上明細RS As New ADODB.Recordset
    Dim 注文ID As String
  
    顧客ID = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_顧客ID)
    
    If 顧客ID = "ID" Then Exit Sub
    
    ' 選択された顧客の注文データを取得する
    Select Case G_タブNO
        Case 1
            Call 顧客マスタ1件読込(顧客マスタRS, 顧客ID)
        Case 2
            Call 配送先1件読込(顧客マスタRS, 顧客ID)
        Case 3
            Exit Sub
    End Select
    
    With 顧客マスタRS
        If 顧客マスタRS.EOF Then
            txt顧客ID.Text = 顧客ID
            txt顧客名.Text = ""
            txtフリガナ.Text = ""
            txt郵便番号.Text = ""
            opt男性.Value = True
            opt男性.Value = True
            txt住所_上段.Text = ""
            txt住所_中段.Text = ""
            txt住所_下段.Text = ""
            txt電話番号.Text = ""
            txtメール.Text = ""
            txt楽天メール.Text = ""
            cmbアーデルクラブ.Text = ""
            txt入会日.Text = "____/__/__"
            txt退会日.Text = "____/__/__"
            txt備考.Text = ""
            chkメール送信.Value = 1
            txt誕生日.Text = "____/__/__"
            
            G_顧客マスタ_排他フラグ = True
            chk資料1.Value = 0
            chk資料2.Value = 0
            chk資料3.Value = 0
            chk資料4.Value = 0
            chk資料5.Value = 0
            G_顧客マスタ_排他フラグ = False
            
            If G_タブNO = 1 Then
                G_顧客_更新日時 = Now
            Else
                G_配送_更新日時 = Now
            End If
        Else
            txt顧客ID.Text = !顧客ID
            txt顧客名.Text = !顧客名
            txtフリガナ.Text = !フリガナ
            txt郵便番号.Text = ![〒]
            If !性別 = "1" Then opt男性.Value = True Else opt男性.Value = False
            If !性別 = "2" Then opt女性.Value = True Else opt女性.Value = False
            txt住所_上段.Text = !住所1
            txt住所_中段.Text = !住所2
            txt住所_下段.Text = IIf(IsNull(!住所3), "", !住所3)
            txt電話番号.Text = !電話番号
            txtメール.Text = !メール
            
            If G_タブNO = 1 Then
                txt楽天メール.Text = IIf(IsNull(!楽天メール), "", !楽天メール)
                chkメール送信.Value = !メール送信
                
                If IsNull(!誕生日) Or !誕生日 = "" Then
                    txt誕生日.Text = "____/__/__"
                Else
                    txt誕生日.Text = !誕生日
                End If
            End If
            
            If G_タブNO = 1 Then
                cmbアーデルクラブ.Text = !アーデルクラブ
                If IsNull(!入会日) Or !入会日 = "" Then
                    txt入会日.Text = "____/__/__"
                Else
                    txt入会日.Text = !入会日
                End If
                
                If IsNull(!退会日) Or !退会日 = "" Then
                    txt退会日.Text = "____/__/__"
                Else
                    txt退会日.Text = !退会日
                End If
            End If
            
            txt備考.Text = !備考
            
            If G_タブNO = 1 Then
                G_顧客マスタ_排他フラグ = True
                chk資料1.Value = IIf(IsNull(!資料1), 0, !資料1)
                chk資料2.Value = IIf(IsNull(!資料2), 0, !資料2)
                chk資料3.Value = IIf(IsNull(!資料3), 0, !資料3)
                chk資料4.Value = IIf(IsNull(!資料4), 0, !資料4)
                chk資料5.Value = IIf(IsNull(!資料5), 0, !資料5)
                G_顧客マスタ_排他フラグ = False
            End If
            
            If G_タブNO = 1 Then
                G_顧客_更新日時 = IIf(IsNull(!更新日時), Now, !更新日時)
            Else
                G_配送_更新日時 = IIf(IsNull(!更新日時), Now, !更新日時)
            End If
                        
        End If
    End With
    
    顧客マスタRS.Close
    
#If 0 Then
    注文ID = SpreadGetVal(va注文リスト, G_注文リスト_ROW, COL_注文ID)
    
    ' 選択された顧客の注文データを取得する
    Call 注文検索2(注文ID, 売上明細RS)
    G_売上_更新日時 = IIf(IsNull(売上明細RS!更新日時), Now, 売上明細RS!更新日時)
    売上明細RS.Close
#End If

End Sub

'************************************************************************
'機  能　注文詳細を表示する。
'************************************************************************
Private Sub 注文タブ_表示()
    
    Dim 売上明細RS As New ADODB.Recordset
    Dim 注文ID As String
    
    Dim 顧客ID As String
    Dim 顧客マスタRS As New ADODB.Recordset
      
    注文ID = SpreadGetVal(va注文リスト, G_注文リスト_ROW, COL_注文ID)
    
    ' 選択された顧客の注文データを取得する
    Call 注文検索2(注文ID, 売上明細RS)
        
    With 売上明細RS
        If 売上明細RS.EOF Then
            txt注文ID.Text = ""
            txt受注日.Text = Format(Now, "YYYY/MM/DD")
            cmbステータス.Text = "新規注文"
            cmb商品名.Text = ""
            cmb部門.Text = "ｱｰﾃﾞﾙ"
            cmb注文方法.Text = "クレジット"
            cmb銀行.Text = ""
            txt配達日時.Text = ""
            txt配達日時2.Text = ""
            txt出荷日.Text = "____/__/__"
            cmb宅配業者.Text = "佐川急便"
            txt支払番号.Text = ""
            txt問合番号.Text = ""
            txt仕入金額.Value = 0
            txt入金日.Text = "____/__/__"
            txt単価.Value = 0
            txt割引.Value = 0
            'cmd割引.Caption = "%"
            txt数量.Value = 1
            txt送料.Value = 0
            txt荷造運賃.Value = 0
            txt返金.Value = 0
            txtその他手数料.Value = 0
            txt合計金額.Text = 0
            txtメール送信.Text = ""
            cmb注文元.Text = ""
            txt注文番号.Text = ""
            txt備考2.Text = ""
            txtコモライフ = ""
            txt出荷予定日.Text = "____/__/__"
            txt決済URL.Text = ""

            G_売上_更新日時 = Now
            G_商品名 = ""
            G_注文元 = ""

        Else
            txt注文ID.Text = IIf(Not IsNull(!注文ID), !注文ID, "")
            txt受注日.Text = IIf(Not IsNull(!受注日), IIf(!受注日 <> "", !受注日, "____/__/__"), "____/__/__")
            cmbステータス.Text = IIf(Not IsNull(!ステータス), !ステータス, "")
            cmb商品名.Text = IIf(Not IsNull(!商品名), !商品名, "")
            
            If IsNull(!部門) = True Then
                If アーデル判定(cmb商品名.Text) = 1 Or アーデル判定(cmb商品名.Text) = 9 Then
                    cmb部門.Text = "ｱｰﾃﾞﾙ"
                Else
                    cmb部門.Text = "ｺﾓﾗｲﾌ"
                End If
            Else
                cmb部門.Text = !部門
            End If
            cmb注文方法.Text = IIf(Not IsNull(!注文方法), !注文方法, "")
            cmb銀行.Text = IIf(Not IsNull(!銀行), !銀行, "")
            txt配達日時.Text = IIf(Not IsNull(!配達希望日時), !配達希望日時, "")
            txt配達日時2.Text = IIf(Not IsNull(!配達希望日時2), !配達希望日時2, "")
            txt出荷日.Text = IIf(Not IsNull(!出荷日), IIf(!出荷日 <> "", !出荷日, "____/__/__"), "____/__/__")
            cmb宅配業者 = IIf(Not IsNull(!宅配業者), !宅配業者, "")
            txt支払番号.Text = IIf(Not IsNull(!支払番号), !支払番号, "")
            txt問合番号.Text = IIf(Not IsNull(!問合番号), !問合番号, "")
            txt仕入金額.Value = IIf(Not IsNull(!仕入金額), !仕入金額, 0)
            txt入金日.Text = IIf(Not IsNull(!入金日), IIf(!入金日 <> "", !入金日, "____/__/__"), "____/__/__")
            txt単価.Value = IIf(Not IsNull(!単価), !単価, 0)
            txt割引.Value = IIf(Not IsNull(!割引), !割引, 0)
            'cmd割引.Caption = IIf(Not IsNull(!割引区分), !割引区分, "%")
            txt数量.Value = IIf(Not IsNull(!数量), !数量, 0)
            txt送料.Value = IIf(Not IsNull(!送料), !送料, 0)
            txt荷造運賃.Value = IIf(Not IsNull(!荷造運賃), !荷造運賃, 0)
            txt返金.Value = IIf(Not IsNull(!返金), !返金, 0)
            txtその他手数料.Value = IIf(Not IsNull(!その他手数料), !その他手数料, 0)
            txt合計金額.Text = IIf(Not IsNull(!合計金額), !合計金額, "")
            txtメール送信.Text = IIf(Not IsNull(!メール送信), !メール送信, "")
            cmb注文元.Text = IIf(Not IsNull(!注文元), !注文元, "")
            txt注文番号.Text = IIf(Not IsNull(!Yahoo注文番号), !Yahoo注文番号, "")
            txt備考2.Text = IIf(Not IsNull(!備考1), !備考1, "")
            txtコモライフ = IIf(Not IsNull(!コモライフNO), !コモライフNO, "")
            txt出荷予定日.Text = IIf(Not IsNull(!出荷予定日), IIf(!出荷予定日 <> "", !出荷予定日, "____/__/__"), "____/__/__")
            txt決済URL.Text = IIf(Not IsNull(!決済URL), !決済URL, "")


            G_売上_更新日時 = IIf(IsNull(!更新日時), Now, !更新日時)
        
            G_商品名 = cmb商品名.Text
            G_注文元 = cmb注文元.Text
        
        End If
                
    End With
    
    With 顧客マスタRS
        顧客ID = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_顧客ID)
        
        If 顧客ID = "ID" Then Exit Sub
        If 顧客ID = "" Then Exit Sub
        
        Call 顧客マスタ1件読込(顧客マスタRS, 顧客ID)
        
        If Not 顧客マスタRS.EOF Then
            G_顧客マスタ_排他フラグ = True
            chk資料1_1.Value = IIf(IsNull(!資料1), 0, !資料1)
            chk資料2_1.Value = IIf(IsNull(!資料2), 0, !資料2)
            chk資料3_1.Value = IIf(IsNull(!資料3), 0, !資料3)
            chk資料4_1.Value = IIf(IsNull(!資料4), 0, !資料4)
            chk資料5_1.Value = IIf(IsNull(!資料5), 0, !資料5)
            G_顧客マスタ_排他フラグ = False
        End If
        顧客マスタRS.Close
    End With
    
#If 0 Then

    顧客ID = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_顧客ID)
    
    If 顧客ID = "ID" Then Exit Sub
    
    ' 選択された顧客の注文データを取得する
    Call 顧客マスタ1件読込(顧客マスタRS, 顧客ID)
    G_顧客_更新日時 = IIf(IsNull(顧客マスタRS!更新日時), Now, 顧客マスタRS!更新日時)
    顧客マスタRS.Close
    
    Call 配送先1件読込(顧客マスタRS, 顧客ID)
    G_配送_更新日時 = IIf(IsNull(顧客マスタRS!更新日時), Now, 顧客マスタRS!更新日時)
    顧客マスタRS.Close
    
#End If

End Sub

'************************************************************************
'機  能　メール履歴を表示する。
'************************************************************************
Private Sub メール履歴_表示()
    
    Dim メール履歴RS As New ADODB.Recordset
    Dim 注文ID As String
    
    注文ID = SpreadGetVal(va注文リスト, G_注文リスト_ROW, COL_注文ID)
    
    If 注文ID = "" Or 注文ID = "注文ID" Then
        Exit Sub
    End If
    
    ' 選択された注文のメール履歴を取得する
    Call メール履歴検索(注文ID, メール履歴RS)
    
    With メール履歴RS
        If メール履歴RS.EOF Then
            vaメール履歴.MaxRows = 0
            txtメール本文 = ""
        Else
            vaメール履歴.MaxRows = 0
            Do Until .EOF
                vaメール履歴.MaxRows = vaメール履歴.MaxRows + 1
                Call SpreadSetVal(vaメール履歴, vaメール履歴.MaxRows, 1, Format(!送信日時, "yyyy/mm/dd hh:mm:ss"))
                Call SpreadSetVal(vaメール履歴, vaメール履歴.MaxRows, 2, !件名)
                .MoveNext
            Loop
            .Close
            If vaメール履歴.MaxRows > 0 Then
                Call vaメール履歴_Click(1, 1)
            End If
        End If
    End With
    
End Sub

'************************************************************************
'機  能 :バックグラウンドで顧客タブに情報を設定し直す。
'************************************************************************
Private Sub 顧客タブ_表示2()

    Dim 顧客ID As String
    Dim 顧客マスタRS As New ADODB.Recordset
  
    顧客ID = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_顧客ID)
    
    ' 最初に配送先を読み込む
    Call 配送先1件読込(顧客マスタRS, 顧客ID)
    
    ' 配送先が登録されていなければ、顧客情報を読み込む
    If 顧客マスタRS.EOF Then
        Call 顧客マスタ1件読込(顧客マスタRS, 顧客ID)
    End If
    
    With 顧客マスタRS
        If 顧客マスタRS.EOF Then
            txt顧客ID.Text = 顧客ID
            txt顧客名.Text = ""
            txtフリガナ.Text = ""
            txt郵便番号.Text = ""
            opt男性.Value = True
            opt男性.Value = True
            txt住所_上段.Text = ""
            txt住所_中段.Text = ""
            txt住所_下段.Text = ""
            txt電話番号.Text = ""
            txtメール.Text = ""
            txt楽天メール.Text = ""
            cmbアーデルクラブ.Text = ""
            txt入会日.Text = "____/__/__"
            txt退会日.Text = "____/__/__"
            txt備考.Text = ""
            chkメール送信.Value = 1
            txt誕生日.Text = "____/__/__"
            chk資料1.Value = 0
            chk資料2.Value = 0
            chk資料3.Value = 0
            chk資料4.Value = 0
            chk資料5.Value = 0
        Else
            txt顧客ID.Text = !顧客ID
            txt顧客名.Text = !顧客名
            txtフリガナ.Text = !フリガナ
            txt郵便番号.Text = ![〒]
            If !性別 = "1" Then opt男性.Value = True Else opt男性.Value = False
            If !性別 = "2" Then opt女性.Value = True Else opt女性.Value = False
            txt住所_上段.Text = !住所1
            txt住所_中段.Text = !住所2
            txt住所_下段.Text = IIf(IsNull(!住所3), "", !住所3)
            txt電話番号.Text = !電話番号
            txtメール.Text = !メール
            txt楽天メール.Text = !楽天メール
            
            If G_タブNO = 1 Then
                chkメール送信.Value = !メール送信
                If IsNull(!誕生日) Or !誕生日 = "" Then
                    txt誕生日.Text = "____/__/__"
                Else
                    txt誕生日.Text = !誕生日
                End If
            End If
            
            If G_タブNO = 1 Then
                cmbアーデルクラブ.Text = !アーデルクラブ
                If IsNull(!入会日) Or !入会日 = "" Then
                    txt入会日.Text = "____/__/__"
                Else
                    txt入会日.Text = !入会日
                End If
                
                If IsNull(!退会日) Or !退会日 = "" Then
                    txt退会日.Text = "____/__/__"
                Else
                    txt退会日.Text = !退会日
                End If
            End If
            
            txt備考.Text = !備考
            
            If G_タブNO = 1 Then
                chk資料1.Value = IIf(IsNull(!資料1), 0, !資料1)
                chk資料2.Value = IIf(IsNull(!資料2), 0, !資料2)
                chk資料3.Value = IIf(IsNull(!資料3), 0, !資料3)
                chk資料4.Value = IIf(IsNull(!資料4), 0, !資料4)
                chk資料5.Value = IIf(IsNull(!資料5), 0, !資料5)
            End If
        End If
    End With
    
    顧客マスタRS.Close
    
End Sub

'************************************************************************
'機  能　顧客名入力制御
'************************************************************************
Private Sub txt顧客名_KeyDown(KeyCode As Integer, Shift As Integer)
    
'    Dim 苗字 As String
'    Dim 名前 As String
    
'    Dim 開始位置 As Integer
    
'    開始位置 = InStr(txt顧客名.Text, " ")
    
'    If 開始位置 > 0 Then
'        苗字 = Left(txt顧客名.Text, 開始位置 - 1)
'        名前 = Mid(txt顧客名.Text, 開始位置 + 1)
'        txt顧客名.Text = 苗字 & "　" & 名前
'    End If
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt顧客名_Validate(Cancel As Boolean)
    
    Dim 苗字 As String
    Dim 名前 As String
    
    Dim 開始位置 As Integer
    
    開始位置 = InStr(txt顧客名.Text, " ")
    
    If 開始位置 > 0 Then
        苗字 = Left(txt顧客名.Text, 開始位置 - 1)
        名前 = Mid(txt顧客名.Text, 開始位置 + 1)
        txt顧客名.Text = 苗字 & "　" & 名前
    End If
    
    Cancel = 顧客情報_登録
    
End Sub

Private Sub txt顧客名_GotFocus()

    txt顧客名.BackColor = vbYellow
    
End Sub

Private Sub txt顧客名_LostFocus()
    
    Call 顧客情報_登録
    
    txt顧客名.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　フリガナ入力制御
'************************************************************************
Private Sub txtフリガナ_KeyDown(KeyCode As Integer, Shift As Integer)
    
'    Dim 苗字 As String
'    Dim 名前 As String
    
'    Dim 開始位置 As Integer
    
'    開始位置 = InStr(txtフリガナ.Text, " ")
    
'    If 開始位置 > 0 Then
'        苗字 = Left(txtフリガナ.Text, 開始位置 - 1)
'        名前 = Mid(txtフリガナ.Text, 開始位置 + 1)
'        txtフリガナ.Text = 苗字 & "　" & 名前
'    End If
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txtフリガナ_Validate(Cancel As Boolean)
    
    Dim 苗字 As String
    Dim 名前 As String
    
    Dim 開始位置 As Integer
    
    開始位置 = InStr(txtフリガナ.Text, " ")
    
    If 開始位置 > 0 Then
        苗字 = Left(txtフリガナ.Text, 開始位置 - 1)
        名前 = Mid(txtフリガナ.Text, 開始位置 + 1)
        txtフリガナ.Text = 苗字 & "　" & 名前
    End If

    Cancel = 顧客情報_登録()
    
End Sub

Private Sub txtフリガナ_GotFocus()

    txtフリガナ.BackColor = vbYellow
    
End Sub

Private Sub txtフリガナ_LostFocus()

    txtフリガナ.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　フリガナ入力制御
'************************************************************************
Private Sub txt郵便番号_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt郵便番号_Validate(Cancel As Boolean)

    Cancel = 顧客情報_登録()
    
    Call 郵便番号から住所を変換する
    
End Sub

Private Sub txt郵便番号_GotFocus()

    txt郵便番号.BackColor = vbYellow
    
End Sub

Private Sub txt郵便番号_LostFocus()

    txt郵便番号.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　男性入力制御
'************************************************************************
Private Sub opt男性_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub opt男性_Validate(Cancel As Boolean)

    Cancel = 顧客情報_登録()

End Sub

Private Sub opt男性_GotFocus()

    opt男性.BackColor = vbYellow
    
End Sub

Private Sub opt男性_LostFocus()

    opt男性.BackColor = &H8000000F

End Sub

'************************************************************************
'機  能　女性入力制御
'************************************************************************
Private Sub opt女性_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub opt女性_Validate(Cancel As Boolean)

    Cancel = 顧客情報_登録()

End Sub

Private Sub opt女性_GotFocus()

    opt女性.BackColor = vbYellow
    
End Sub

Private Sub opt女性_LostFocus()

    opt女性.BackColor = &H8000000F

End Sub

'************************************************************************
'機  能　住所_上段入力制御
'************************************************************************
Private Sub txt住所_上段_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt住所_上段_Validate(Cancel As Boolean)

    Cancel = 顧客情報_登録()
    
End Sub

Private Sub txt住所_上段_GotFocus()

    txt住所_上段.BackColor = vbYellow
    
End Sub

Private Sub txt住所_上段_LostFocus()

    txt住所_上段.BackColor = vbWhite

End Sub


'************************************************************************
'機  能　住所_中段入力制御
'************************************************************************
Private Sub txt住所_中段_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt住所_中段_Validate(Cancel As Boolean)

    Cancel = 顧客情報_登録()
    
End Sub

Private Sub txt住所_中段_GotFocus()

    txt住所_中段.BackColor = vbYellow
    
End Sub

Private Sub txt住所_中段_LostFocus()

    txt住所_中段.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　住所_下段入力制御
'************************************************************************
Private Sub txt住所_下段_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt住所_下段_Validate(Cancel As Boolean)

    Cancel = 顧客情報_登録()
    
End Sub

Private Sub txt住所_下段_GotFocus()

    txt住所_下段.BackColor = vbYellow
    
End Sub

Private Sub txt住所_下段_LostFocus()

    txt住所_下段.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　電話番号入力制御
'************************************************************************
Private Sub txt電話番号_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt電話番号_Validate(Cancel As Boolean)

    Cancel = 顧客情報_登録()
    
End Sub

Private Sub txt電話番号_GotFocus()

    txt電話番号.BackColor = vbYellow
    
End Sub

Private Sub txt電話番号_LostFocus()

    txt電話番号.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　メール入力制御
'************************************************************************
Private Sub txtメール_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txtメール_Validate(Cancel As Boolean)
    
    txtメール.Text = Trim(txtメール.Text)
    
    Cancel = 顧客情報_登録()
    
End Sub

Private Sub txtメール_GotFocus()

    txtメール.BackColor = vbYellow
    
End Sub

Private Sub txtメール_LostFocus()

    txtメール.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　楽天メール入力制御
'************************************************************************
Private Sub txt楽天メール_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt楽天メール_Validate(Cancel As Boolean)
    
    txt楽天メール.Text = Trim(txt楽天メール.Text)
    
    Cancel = 顧客情報_登録()
    
End Sub

Private Sub txt楽天メール_GotFocus()

    txt楽天メール.BackColor = vbYellow
    
End Sub

Private Sub txt楽天メール_LostFocus()

    txt楽天メール.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　メール送信入力制御
'************************************************************************
Private Sub chkメール送信_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub chkメール送信_Validate(Cancel As Boolean)

    Cancel = 顧客情報_登録()
    
End Sub

Private Sub chkメール送信_GotFocus()

    chkメール送信.BackColor = vbYellow
    
End Sub

Private Sub chkメール送信_LostFocus()

    chkメール送信.BackColor = &H8000000F

End Sub

'************************************************************************
'機  能　アーデルクラブ入力制御
'************************************************************************
Private Sub cmbアーデルクラブ_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub cmbアーデルクラブ_Validate(Cancel As Boolean)

    If Len(cmbアーデルクラブ.Text) > 10 Then
        Call MsgBox("アーデルクラブが長すぎます。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If

    Cancel = 顧客情報_登録()
    
End Sub

Private Sub cmbアーデルクラブ_GotFocus()

    cmbアーデルクラブ.BackColor = vbYellow
    
End Sub

Private Sub cmbアーデルクラブ_LostFocus()

    cmbアーデルクラブ.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　入会日入力制御
'************************************************************************
Private Sub txt入会日_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt入会日_Validate(Cancel As Boolean)
        
    If txt入会日.Text <> "____/__/__" Then
        If IsDate(txt入会日.Text) = False Then
            Call MsgBox("正しい入会日を入力して下さい。", vbOKOnly, "顧客管理")
            Cancel = True
            Exit Sub
        End If
    End If
    
    Cancel = 顧客情報_登録()

End Sub

Private Sub txt入会日_GotFocus()

    txt入会日.BackColor = vbYellow
    
End Sub

Private Sub txt入会日_LostFocus()

    txt入会日.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　退会日入力制御
'************************************************************************
Private Sub txt退会日_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt退会日_Validate(Cancel As Boolean)

    If txt退会日.Text <> "____/__/__" Then
        If IsDate(txt退会日.Text) = False Then
            Call MsgBox("正しい退会日を入力して下さい。", vbOKOnly, "顧客管理")
            Cancel = True
            Exit Sub
        End If
    End If
    
    Cancel = 顧客情報_登録()

End Sub

Private Sub txt退会日_GotFocus()

    txt退会日.BackColor = vbYellow
    
End Sub

Private Sub txt退会日_LostFocus()

    txt退会日.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　備考入力制御
'************************************************************************
Private Sub txt備考_KeyDown(KeyCode As Integer, Shift As Integer)
    
    ' 備考はマルチライン入力なので、改行キーを押下されてもフィールドを移動しない。
    'Call タブキー送信(KeyCode)

End Sub

Private Sub txt備考_Validate(Cancel As Boolean)

    If Len(txt備考) >= 4096 Then
        Call MsgBox("備考の入力桁数が大きいです。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If

    Cancel = 顧客情報_登録()
    
End Sub

Private Sub txt備考_GotFocus()

    txt備考.BackColor = vbYellow
    
End Sub

Private Sub txt備考_LostFocus()

    txt備考.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　誕生日入力制御
'************************************************************************
Private Sub txt誕生日_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt誕生日_Validate(Cancel As Boolean)

    If txt誕生日.Text <> "____/__/__" Then
        If IsDate(txt誕生日.Text) = False Then
            Call MsgBox("正しい誕生日を入力して下さい。", vbOKOnly, "顧客管理")
            Cancel = True
            Exit Sub
        End If
    End If
    
    Cancel = 顧客情報_登録()

End Sub

Private Sub txt誕生日_GotFocus()

    txt誕生日.BackColor = vbYellow
    
End Sub

Private Sub txt誕生日_LostFocus()

    txt誕生日.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　資料１ボタンクリック
'************************************************************************
Private Sub chk資料1_Click()

    If G_顧客マスタ_排他フラグ = False Then
        Call 顧客情報_登録
    End If
    
End Sub

'************************************************************************
'機  能　資料２ボタンクリック
'************************************************************************
Private Sub chk資料2_Click()

    If G_顧客マスタ_排他フラグ = False Then
        Call 顧客情報_登録
    End If

End Sub

'************************************************************************
'機  能　資料３ボタンクリック
'************************************************************************
Private Sub chk資料3_Click()

    If G_顧客マスタ_排他フラグ = False Then
        Call 顧客情報_登録
    End If

End Sub

'************************************************************************
'機  能　資料４ボタンクリック
'************************************************************************
Private Sub chk資料4_Click()

    If G_顧客マスタ_排他フラグ = False Then
        Call 顧客情報_登録
    End If

End Sub

'************************************************************************
'機  能　資料５ボタンクリック
'************************************************************************
Private Sub chk資料5_Click()

    If G_顧客マスタ_排他フラグ = False Then
        Call 顧客情報_登録
    End If

End Sub

'************************************************************************
'機  能　資料１ボタンクリック
'************************************************************************
Private Sub chk資料1_1_Click()
    
    Dim 顧客ID As String

    If G_顧客マスタ_排他フラグ = False Then
        顧客ID = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_顧客ID)
        Call 資料チェック更新1(顧客ID, chk資料1_1.Value)
    End If
    
End Sub

'************************************************************************
'機  能　資料２ボタンクリック
'************************************************************************
Private Sub chk資料2_1_Click()

    Dim 顧客ID As String

    If G_顧客マスタ_排他フラグ = False Then
        顧客ID = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_顧客ID)
        Call 資料チェック更新2(顧客ID, chk資料2_1.Value)
    End If

End Sub

'************************************************************************
'機  能　資料３ボタンクリック
'************************************************************************
Private Sub chk資料3_1_Click()

    Dim 顧客ID As String

    If G_顧客マスタ_排他フラグ = False Then
        顧客ID = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_顧客ID)
        Call 資料チェック更新3(顧客ID, chk資料3_1.Value)
    End If

End Sub

'************************************************************************
'機  能　資料４ボタンクリック
'************************************************************************
Private Sub chk資料4_1_Click()

    Dim 顧客ID As String

    If G_顧客マスタ_排他フラグ = False Then
        顧客ID = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_顧客ID)
        Call 資料チェック更新4(顧客ID, chk資料4_1.Value)
    End If

End Sub

'************************************************************************
'機  能　資料５ボタンクリック
'************************************************************************
Private Sub chk資料5_1_Click()
    
    Dim 顧客ID As String

    If G_顧客マスタ_排他フラグ = False Then
        顧客ID = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_顧客ID)
        Call 資料チェック更新5(顧客ID, chk資料5_1.Value)
    End If

End Sub

'************************************************************************
'機  能　転記ボタン
'************************************************************************
Private Sub cmd転記_Click()

    Dim 顧客ID As String
    Dim 顧客マスタRS As New ADODB.Recordset
  
    顧客ID = txt顧客ID.Text
    
    Call 顧客マスタ1件読込(顧客マスタRS, 顧客ID)
    
    With 顧客マスタRS
        If 顧客マスタRS.EOF Then
            txt顧客名.Text = ""
            txtフリガナ.Text = ""
            txt郵便番号.Text = ""
            opt男性.Value = True
            opt男性.Value = True
            txt住所_上段.Text = ""
            txt住所_中段.Text = ""
            txt住所_下段.Text = ""
            txt電話番号.Text = ""
            txtメール.Text = ""
        Else
            txt顧客名.Text = !顧客名
            txtフリガナ.Text = !フリガナ
            txt郵便番号.Text = ![〒]
            If !性別 = "1" Then opt男性.Value = True Else opt男性.Value = False
            If !性別 = "2" Then opt女性.Value = True Else opt女性.Value = False
            txt住所_上段.Text = !住所1
            txt住所_中段.Text = !住所2
            txt住所_下段.Text = IIf(IsNull(!住所3), "", !住所3)
            txt電話番号.Text = !電話番号
            txtメール.Text = !メール
            txt顧客名.SetFocus
        End If
    End With
    
    顧客マスタRS.Close
    
End Sub

'************************************************************************
'機  能　受注日入力制御
'************************************************************************
Private Sub txt受注日_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt受注日_Validate(Cancel As Boolean)
    
    If txt受注日.Text = "____/__/__" Then
        Call MsgBox("受注日を入力して下さい。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If
    
    If IsDate(txt受注日.Text) = False Then
        Call MsgBox("正しい受注日を入力して下さい。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If
    
    Cancel = 注文_更新()

End Sub

Private Sub txt受注日_GotFocus()

    txt受注日.BackColor = vbYellow
    
End Sub

Private Sub txt受注日_LostFocus()

    txt受注日.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　ステータス入力制御
'************************************************************************
Private Sub cmbステータス_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub cmbステータス_Validate(Cancel As Boolean)
    
    If Len(cmbステータス.Text) > 20 Then
        Call MsgBox("ステータスの文字列が長すぎます。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If
    
    If cmbステータス.Text = "出荷完了" Then
        If txt出荷日.Text = "____/__/__" Then
            txt出荷日.Text = Format(Date, "yyyy/mm/dd")
        End If
    End If
    
    cmb部門.Text = "その他"
    
    If cmbステータス.Text = "コモライフ" Then
        cmb部門.Text = "ｺﾓﾗｲﾌ"
    End If
    
    If アーデル判定(cmb商品名.Text) = 1 Or アーデル判定(cmb商品名.Text) = 9 Then
        cmb部門.Text = "ｱｰﾃﾞﾙ"
    End If
    
    Cancel = 注文_更新()
    
End Sub

Private Sub cmbステータス_GotFocus()

    cmbステータス.BackColor = vbYellow
    
End Sub

Private Sub cmbステータス_LostFocus()

    cmbステータス.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　商品名入力制御
'************************************************************************
Private Sub cmb商品名_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub cmb商品名_Validate(Cancel As Boolean)
    
    Dim 累積本数 As Long
    Dim 商品マスタRS As New ADODB.Recordset
    
    If Len(cmb商品名.Text) >= 100 Then
        Call MsgBox("商品名が長すぎます。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If
    
    If cmb注文元.Text = "" Then
        Call MsgBox("注文元を選択して下さい。", vbOKOnly, "顧客管理")
        Exit Sub
    End If
    
    If cmb商品名.Text = G_商品名 Then
        Exit Sub
    End If
    
    G_商品名 = cmb商品名.Text
    
    Call 商品マスタ取得(cmb商品名.Text, 商品マスタRS)
    
    With 商品マスタRS
        If Not 商品マスタRS.EOF Then
            'If cmb注文元.Text = "Yahoo" Or cmb注文元.Text = "楽天" Or cmb注文元.Text = "おちゃのこネット" Or cmb注文元.Text = "アマゾン" Or cmb注文元.Text = "たんぽぽモール" Then
                
                累積本数 = 累積数計算2(txt注文ID.Text)
                
#If 0 Then
                If cmb商品名.Text = "アーデル" Then
                    
                    If 累積本数 < 1 Then
                        txt単価.Value = !単価
                        cmd割引.Caption = "\"
                        txt割引.Value = !割引金額
                        txt送料.Value = !送料
                        txt数量.Value = 1
                        txtその他手数料.Value = 0
                        txt返金.Value = 0
                        
                    ElseIf 累積本数 >= 1 And 累積本数 <= 5 Then
                        txt単価.Value = !単価
                        cmd割引.Caption = "%"
                        txt割引.Value = 10
                        txt送料.Value = 0
                        txt数量.Value = 1
                        txtその他手数料.Value = 0
                        txt返金.Value = 0
                    
                    ElseIf 累積本数 >= 6 Then
                         txt単価.Value = !単価
                        cmd割引.Caption = "%"
                        txt割引.Value = 20
                        txt送料.Value = 0
                        txt数量.Value = 1
                        txtその他手数料.Value = 0
                        txt返金.Value = 0
                   End If
                Else
#End If
                    txt単価.Value = !単価
                    'cmd割引.Caption = "\"
                    txt割引.Value = !割引金額
                    txt送料.Value = !送料
                    txt数量.Value = 1
                    txtその他手数料.Value = 0
                    txt返金.Value = 0
                'End If
            'Else
            '    txt単価 = !単価
            '    'cmd割引.Caption = "%"
            '    txt割引.Value = 0
            '    txt送料.Text = !送料
            '    txt数量.Value = 1
            '    txtその他手数料.Value = 0
            '    txt返金.Value = 0
            'End If
        End If
        .Close
    End With
    
    Call 再計算
    
    If cmb商品名.Text = "アーデル" _
        Or cmb商品名.Text = "アーデル2本セット" _
        Or cmb商品名.Text = "アーデル＋シャンプー" _
        Or cmb商品名.Text = "新ブスタ" _
        Or cmb商品名.Text = "新ブスタ＋シャンプー" _
        Or cmb商品名.Text = "ブースター" _
        Or cmb商品名.Text = "ブースター（Ｗ発毛月間）" _
        Or cmb商品名.Text = "ブースター＋シャンプー" _
        Or cmb商品名.Text = "新ハイブリッター" _
        Or cmb商品名.Text = "新ハイブリッター＋シャンプー" _
        Or cmb商品名.Text = "ハイブリッド" _
        Or cmb商品名.Text = "ハイブリッド＋シャンプー" _
        Or cmb商品名.Text = "ナイスレディー" _
        Or cmb商品名.Text = "ナイスレディー＋シャンプー" _
        Or cmb商品名.Text = "シャンプー" _
        Or cmb商品名.Text = "シャンプー（プレゼント）" _
        Or cmb商品名.Text = "シャンプー2本セット" _
        Or cmb商品名.Text = "シャンプー＋トリートメント" _
        Or cmb商品名.Text = "トリートメント" _
        Or cmb商品名.Text = "トリートメント（プレゼント）" _
        Or cmb商品名.Text = "アーデル＆シャンプー試供品" _
        Or cmb商品名.Text = "アーデル試供品" _
        Or cmb商品名.Text = "シャンプー試供品" Then
        cmb宅配業者.Text = "佐川急便"
    ElseIf cmb商品名.Text = "アーデル資料" _
        Or cmb商品名.Text = "ミニまぐ" Then
        cmb宅配業者.Text = "EXPRESS"
    Else
        cmb宅配業者.Text = "クロネコヤマト"
    End If

    Cancel = 注文_更新()
    
End Sub

Private Sub cmb商品名_GotFocus()

    cmb商品名.BackColor = vbYellow
    
End Sub

Private Sub cmb商品名_LostFocus()

    cmb商品名.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　部門制御
'************************************************************************
Private Sub cmb部門_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub cmb部門_Validate(Cancel As Boolean)

    If Len(cmb部門.Text) > 10 Then
        Call MsgBox("部門が長すぎます。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If
    
    Cancel = 注文_更新()
    
End Sub

Private Sub cmb部門_GotFocus()

    cmb部門.BackColor = vbYellow
    
End Sub

Private Sub cmb部門_LostFocus()

    cmb部門.BackColor = vbWhite

End Sub


'************************************************************************
'機  能　注文方法入力制御
'************************************************************************
Private Sub cmb注文方法_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub cmb注文方法_Validate(Cancel As Boolean)

    If Len(cmb注文方法.Text) > 20 Then
        Call MsgBox("注文方法が長すぎます。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If
    
    If cmb注文方法.Text = "コンビニ" Then
        txt決済URL.Caption = "決済URL"
    Else
        txt決済URL.Caption = "決済ID"
    End If
    
    Cancel = 注文_更新()
    
End Sub

Private Sub cmb注文方法_GotFocus()

    cmb注文方法.BackColor = vbYellow
    
End Sub

Private Sub cmb注文方法_LostFocus()

    cmb注文方法.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　銀行入力制御
'************************************************************************
Private Sub cmb銀行_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub cmb銀行_Validate(Cancel As Boolean)

    If Len(cmb銀行.Text) > 10 Then
        Call MsgBox("銀行名が長すぎます。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If
    
    Cancel = 注文_更新()
    
End Sub

Private Sub cmb銀行_GotFocus()

    cmb銀行.BackColor = vbYellow
    
End Sub

Private Sub cmb銀行_LostFocus()

    cmb銀行.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　配達日時入力制御
'************************************************************************
Private Sub txt配達日時_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt配達日時_Validate(Cancel As Boolean)
    
    Cancel = 注文_更新()
    
End Sub

Private Sub txt配達日時_GotFocus()

    txt配達日時.BackColor = vbYellow
    
End Sub

Private Sub txt配達日時_LostFocus()

    txt配達日時.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　配達日時2入力制御
'************************************************************************
Private Sub txt配達日時2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt配達日時2_Validate(Cancel As Boolean)
    
    Cancel = 注文_更新()
    
End Sub

Private Sub txt配達日時2_GotFocus()

    txt配達日時2.BackColor = vbYellow
    
End Sub

Private Sub txt配達日時2_LostFocus()

    txt配達日時2.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　出荷日入力制御
'************************************************************************
Private Sub txt出荷日_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt出荷日_Validate(Cancel As Boolean)

    If txt出荷日.Text <> "____/__/__" Then
        If IsDate(txt出荷日.Text) = False Then
            Call MsgBox("正しい出荷日を入力して下さい。", vbOKOnly, "顧客管理")
            Cancel = True
            Exit Sub
        End If
    End If
    
    Cancel = 注文_更新()

End Sub

Private Sub txt出荷日_GotFocus()

    txt出荷日.BackColor = vbYellow
    
End Sub

Private Sub txt出荷日_LostFocus()

    txt出荷日.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　出荷日設定
'************************************************************************
Private Sub cmd本日1_Click()
    
    txt出荷日.Text = Format(Date, "yyyy/mm/dd")
    
    cmbステータス.Text = "出荷完了"

    Call 注文_更新
    
End Sub

'************************************************************************
'機  能　宅配業者入力制御
'************************************************************************
Private Sub cmb宅配業者_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub cmb宅配業者_Validate(Cancel As Boolean)

    If Len(cmb宅配業者.Text) > 10 Then
        Call MsgBox("宅配業者が長すぎます。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If

    Cancel = 注文_更新()
    
End Sub

Private Sub cmb宅配業者_GotFocus()

    cmb宅配業者.BackColor = vbYellow
    
End Sub

Private Sub cmb宅配業者_LostFocus()

    cmb宅配業者.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　支払番号入力制御
'************************************************************************
Private Sub txt支払番号_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt支払番号_Validate(Cancel As Boolean)

    Cancel = 注文_更新()
    
End Sub

Private Sub txt支払番号_GotFocus()

    txt支払番号.BackColor = vbYellow
    
End Sub

Private Sub txt支払番号_LostFocus()

    txt支払番号.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　問合番号入力制御
'************************************************************************
Private Sub txt問合番号_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt問合番号_Validate(Cancel As Boolean)

    If txt問合番号.Text <> "" Then
        
        ' cmbステータス.Text = "出荷完了"
        
        If txt出荷日.Text = "____/__/__" Then
            txt出荷日.Text = Format(Date, "yyyy/mm/dd")
        End If
    End If

    Cancel = 注文_更新()
    
End Sub

Private Sub txt問合番号_GotFocus()

    txt問合番号.BackColor = vbYellow
    
End Sub

Private Sub txt問合番号_LostFocus()

    txt問合番号.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　仕入金額入力制御
'************************************************************************
Private Sub txt仕入金額_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt仕入金額_Validate(Cancel As Boolean)
    
    Call 再計算
    
    Cancel = 注文_更新()

End Sub

Private Sub txt仕入金額_GotFocus()

    txt仕入金額.BackColor = vbYellow
    
    Call psubIMEOnOff(Me.hwnd, False)
    
End Sub

Private Sub txt仕入金額_LostFocus()
    
    txt仕入金額.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　決済URL入力制御
'************************************************************************
Private Sub txt決済URL_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt決済URL_Validate(Cancel As Boolean)

    Cancel = 注文_更新()
    
End Sub

Private Sub txt決済URL_GotFocus()

    txt決済URL.BackColor = vbYellow
    
End Sub

Private Sub txt決済URL_LostFocus()

    txt決済URL.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　入金日入力制御
'************************************************************************
Private Sub txt入金日_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt入金日_Validate(Cancel As Boolean)
   
    Call 再計算
   
    If txt入金日.Text <> "____/__/__" Then
        If IsDate(txt入金日.Text) = False Then
            Call MsgBox("正しい入金日を入力して下さい。", vbOKOnly, "顧客管理")
            Cancel = True
            Exit Sub
        End If
    End If
    
    Cancel = 注文_更新()

End Sub

Private Sub txt入金日_GotFocus()

    txt入金日.BackColor = vbYellow
    
End Sub

Private Sub txt入金日_LostFocus()
    
    txt入金日.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　入金日設定
'************************************************************************
Private Sub cmd本日2_Click()
    
    txt入金日.Text = Format(Date, "yyyy/mm/dd")
    
    Call 注文_更新
    
End Sub

'************************************************************************
'機  能　単価入力制御
'************************************************************************
Private Sub txt単価_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt単価_Validate(Cancel As Boolean)

    If txt単価.Value < 0 Then
        Call MsgBox("プラス値を入力して下さい。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If
    
    Call 再計算
    
    Cancel = 注文_更新()

End Sub

Private Sub txt単価_GotFocus()

    txt単価.BackColor = vbYellow
    
    Call psubIMEOnOff(Me.hwnd, False)
    
End Sub

Private Sub txt単価_LostFocus()
    
    txt単価.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　割引入力制御
'************************************************************************
Private Sub txt割引_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt割引_Validate(Cancel As Boolean)

    If txt割引.Value > 0 Then
        Call MsgBox("マイナス値を入力して下さい。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If

    Call 再計算

    Cancel = 注文_更新()

End Sub

Private Sub txt割引_GotFocus()

    txt割引.BackColor = vbYellow
    
    Call psubIMEOnOff(Me.hwnd, False)
    
End Sub

Private Sub txt割引_LostFocus()
    
    txt割引.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　割引入力制御
'************************************************************************
Private Sub cmd割引_Click()
    
    txt割引.Text = -770
    
    Call 再計算
    
End Sub

'************************************************************************
'機  能　割引入力制御
'************************************************************************
Private Sub cmd割引2_Click()
    
    txt割引.Text = -4600
    
    Call 再計算
    
End Sub

'************************************************************************
'機  能　割引入力制御
'************************************************************************
Private Sub cmd割引3_Click()
    
    Dim 単価
    
    If IsNumeric(txt単価.Text) Then
        単価 = CLng(txt単価.Text)
    Else
        単価 = 0
    End If
    
    txt割引.Text = CLng(Format(((単価 * 10) / 100), "0")) * -1
    
    Call 再計算
    
End Sub

'************************************************************************
'機  能　割引入力制御
'************************************************************************
Private Sub cmd割引4_Click()
    
    Dim 単価
    
    If IsNumeric(txt単価.Text) Then
        単価 = CLng(txt単価.Text)
    Else
        単価 = 0
    End If
    
    txt割引.Text = CLng(Format(((単価 * 20) / 100), "0")) * -1
    
    Call 再計算
    
End Sub

'************************************************************************
'機  能　割引入力制御
'************************************************************************
Private Sub cmd割引5_Click()
    
    txt割引.Text = 0
    
    Call 再計算
    
End Sub

'************************************************************************
'機  能　数量入力制御
'************************************************************************
Private Sub txt数量_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt数量_Validate(Cancel As Boolean)

    If txt数量.Value < 0 Then
        Call MsgBox("プラス値を入力して下さい。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If

    Call 再計算
    
    Cancel = 注文_更新()

End Sub

Private Sub txt数量_GotFocus()

    txt数量.BackColor = vbYellow
    
    Call psubIMEOnOff(Me.hwnd, False)
    
End Sub

Private Sub txt数量_LostFocus()
    
    txt数量.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　送料入力制御
'************************************************************************
Private Sub txt送料_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt送料_Validate(Cancel As Boolean)
    
    If txt送料.Value < 0 Then
        Call MsgBox("プラス値を入力して下さい。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If
    
    Call 再計算
    
    Cancel = 注文_更新()

End Sub

Private Sub txt送料_GotFocus()

    txt送料.BackColor = vbYellow
    
    Call psubIMEOnOff(Me.hwnd, False)
    
End Sub

Private Sub txt送料_LostFocus()
    
    txt送料.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　荷造運賃入力制御
'************************************************************************
Private Sub txt荷造運賃_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt荷造運賃_Validate(Cancel As Boolean)
    
    Call 再計算
    
    Cancel = 注文_更新()

End Sub

Private Sub txt荷造運賃_GotFocus()

    txt荷造運賃.BackColor = vbYellow
    
    Call psubIMEOnOff(Me.hwnd, False)
    
End Sub

Private Sub txt荷造運賃_LostFocus()
    
    txt荷造運賃.BackColor = vbWhite

End Sub
'************************************************************************
'機  能　返金入力制御
'************************************************************************
Private Sub txt返金_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt返金_Validate(Cancel As Boolean)

    If txt返金.Value > 0 Then
        Call MsgBox("マイナス値を入力して下さい。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If

    Call 再計算
    
    Cancel = 注文_更新()

End Sub

Private Sub txt返金_GotFocus()

    txt返金.BackColor = vbYellow
        
    Call psubIMEOnOff(Me.hwnd, False)

End Sub

Private Sub txt返金_LostFocus()
    
    txt返金.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　その他手数料入力制御
'************************************************************************
Private Sub txtその他手数料_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txtその他手数料_Validate(Cancel As Boolean)

    If txtその他手数料.Value > 0 Then
        Call MsgBox("マイナス値を入力して下さい。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If
    
    Call 再計算
    
    Cancel = 注文_更新()

End Sub

Private Sub txtその他手数料_GotFocus()

    txtその他手数料.BackColor = vbYellow
    
    Call psubIMEOnOff(Me.hwnd, False)
    
End Sub

Private Sub txtその他手数料_LostFocus()

    Call 再計算
    
    Call 注文_更新
    
    txtその他手数料.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　メール送信入力制御
'************************************************************************
Private Sub txtメール送信_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txtメール送信_Validate(Cancel As Boolean)

    Cancel = 注文_更新()
    
End Sub

Private Sub txtメール送信_GotFocus()

    txtメール送信.BackColor = vbYellow
    
End Sub

Private Sub txtメール送信_LostFocus()

    txtメール送信.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　注文元入力制御
'************************************************************************
Private Sub cmb注文元_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub cmb注文元_Validate(Cancel As Boolean)

    Dim Cancel2 As Boolean
    
    If Len(cmb注文元.Text) > 10 Then
        Call MsgBox("注文元が長すぎます。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If
    
    If G_注文元 <> cmb注文元.Text Then
        G_注文元 = cmb注文元.Text
        G_商品名 = ""
    End If
    
    If cmb注文元.Text = "Yahoo" Or cmb注文元.Text = "楽天" Then
    Else
        If txt注文番号.Text = "" Then
            If cmb注文元.Text = "レントラックス" Then
'            If cmb注文元.Text = "おちゃのこネット" Then
                txt注文番号.Text = ""
            ElseIf cmb注文元.Text = "アマゾン" Then
                txt注文番号.Text = ""
            ElseIf cmb注文元.Text = "コマチ" Then
                txt注文番号.Text = "KOMACHI-" & Format(Now, "yyyymmddhhmmss")
            Else
                txt注文番号.Text = "ETC-" & Format(Now, "yyyymmddhhmmss")
            End If
        End If
        
    End If
    
    Call cmb商品名_Validate(Cancel2)
    
    Cancel = 注文_更新()
    
End Sub

Private Sub cmb注文元_GotFocus()

    cmb注文元.BackColor = vbYellow
    
End Sub

Private Sub cmb注文元_LostFocus()

    'cmb注文元.BackColor = vbWhite
    cmb注文元.BackColor = vbRed

End Sub

'************************************************************************
'機  能　注文番号入力制御
'************************************************************************
Private Sub txt注文番号_KeyDown(KeyCode As Integer, Shift As Integer)
        
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt注文番号_Validate(Cancel As Boolean)

    Cancel = 注文_更新()
    
End Sub

Private Sub txt注文番号_GotFocus()

    txt注文番号.BackColor = vbYellow
    
End Sub

Private Sub txt注文番号_LostFocus()

    txt注文番号.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　コモライフNO入力制御
'************************************************************************
Private Sub txtコモライフ_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txtコモライフ_Validate(Cancel As Boolean)

    Cancel = 注文_更新()
    
End Sub

Private Sub txtコモライフ_GotFocus()

    txtコモライフ.BackColor = vbYellow
    
End Sub

Private Sub txtコモライフ_LostFocus()

    txtコモライフ.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　出荷予定日入力制御
'************************************************************************
Private Sub txt出荷予定日_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call タブキー送信(KeyCode)

End Sub

Private Sub txt出荷予定日_Validate(Cancel As Boolean)

    If txt出荷予定日.Text <> "____/__/__" Then
        If IsDate(txt出荷予定日.Text) = False Then
            Call MsgBox("正しい出荷日予定日を入力して下さい。", vbOKOnly, "顧客管理")
            Cancel = True
            Exit Sub
        End If
    End If
    
    Cancel = 注文_更新()
    
End Sub

Private Sub txt出荷予定日_GotFocus()

    txt出荷予定日.BackColor = vbYellow
    
End Sub

Private Sub txt出荷予定日_LostFocus()

    txt出荷予定日.BackColor = vbWhite

End Sub

'************************************************************************
'機  能　備考2入力制御
'************************************************************************
Private Sub txt備考2_KeyDown(KeyCode As Integer, Shift As Integer)
    
    ' 備考はマルチライン入力なので、改行キーを押下されてもフィールドを移動しない。
    'Call タブキー送信(KeyCode)

End Sub

Private Sub txt備考2_Validate(Cancel As Boolean)
    
    If Len(txt備考2) >= 4096 Then
        Call MsgBox("備考の入力桁数が大きいです。", vbOKOnly, "顧客管理")
        Cancel = True
        Exit Sub
    End If

    Cancel = 注文_更新()
    
End Sub

Private Sub txt備考2_GotFocus()

    txt備考2.BackColor = vbYellow
    
End Sub

Private Sub txt備考2_LostFocus()

    txt備考2.BackColor = vbWhite

End Sub

'************************************************************************
'機  能 :タブキーを送信する
'************************************************************************
Public Sub タブキー送信(KeyCode As Integer)
    If (KeyCode = vbKeyReturn) Or (KeyCode = vbKeyDown) Then
        Me.Tag = "Through"
        SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyUp Then
        Me.Tag = "Through"
        SendKeys "+{TAB}"
    End If
End Sub

'************************************************************************
'機  能　注文詳細が変更内用を補正する。
'************************************************************************
Private Sub 再計算()

    Dim 商品名 As String
    Dim アーデルクラブ As String
    Dim 単価 As Long
    Dim 割引 As String
    Dim 数量 As Long
    Dim 金額 As Long
    Dim 送料 As Long
    Dim 返金 As Long
    Dim その他手数料 As Long
    Dim 合計金額 As Long

#If 0 Then
    商品名 = cmb商品名.Text
    アーデルクラブ = cmbアーデルクラブ.Text
    
    If 商品名 = "アーデル" Then txt単価.Value = 15750
    
    If 商品名 = "スーパーアーデル" Then txt単価.Value = 15750
    
    If 商品名 = "アーデル2本セット" Then txt単価.Value = 31500
    
    If 商品名 = "アーデル＋シャンプー" Then txt単価.Value = 17755
    
    If 商品名 = "アーデル＋シャンプー試供品" Then txt単価.Value = 16000
    
    If 商品名 = "アーデル(セール)" Then txt単価.Value = 9400
    
    If 商品名 = "シャンプー" Then txt単価.Value = 2940
    
    If 商品名 = "シャンプー2本セット" Then txt単価.Value = 5880
    
    If 商品名 = "アーデル試供品" Then txt単価.Value = 1000
    
    If 商品名 = "シャンプー試供品" Then txt単価.Value = 525
    
    If 商品名 = "アーデル試供品＋シャンプー試供品" Then txt単価.Value = 1500
    
    If 商品名 = "アーデルサプリ" Then txt単価.Value = 9800
    
    If 商品名 = "ロゲイン２％" Then txt単価.Value = 3500
    
    If 商品名 = "ロゲイン５％" Then txt単価.Value = 3500
    
    If 商品名 = "チーズスイートホーム／眠り" Then txt単価.Value = 3980
    
    If 商品名 = "ブースター" Then txt単価.Value = 10000
    
    If 商品名 = "ブースター（Ｗ発毛月間）" Then txt単価.Value = 0
    
    If 商品名 = "ハイブッド" Then txt単価.Value = 12600
    
    If 商品名 = "あわあわ水素水" Then txt単価.Value = 2980
    
    If 商品名 = "あわあわ水素水2本セット" Then txt単価.Value = 5960

#End If

    If 商品名 = "アーデル資料" Then
        txt単価.Value = 0
        cmbステータス.Text = "資料請求"
        cmb注文方法.Text = "資料請求"
    End If
    
    If 商品名 = "ミニまぐ" Then
        txt単価.Value = 0
        cmbステータス.Text = "資料請求"
        cmb注文方法.Text = "資料請求"
    End If

    単価 = txt単価.Value
    割引 = txt割引.Value
    数量 = txt数量.Value
    送料 = txt送料.Value
    返金 = txt返金.Value
    金額 = CLng(Format(((単価 + 割引) * 数量), "0"))
    
#If 0 Then
    If cmd割引.Caption = "%" Then
        If 割引 > 0 And 割引 < 100 Then
            金額 = CLng(Format(((単価 * (100 - 割引)) / 100 * 数量), "0"))
        Else
            金額 = 単価 * 数量
        End If
    Else
        If 割引 <> 0 Then
            金額 = CLng(Format(((単価 + 割引) * 数量), "0"))
        Else
            金額 = 単価 * 数量
        End If
    End If
#End If

    その他手数料 = txtその他手数料.Value
    合計金額 = 金額 + 送料 + 返金 + その他手数料

    txt合計金額.Text = 合計金額
    
    Call 注文_更新
    
End Sub

'************************************************************************
'機  能 :顧客リストで選択されている行を更新する。
'************************************************************************
Private Function 顧客情報_登録() As Boolean
    
    顧客情報_登録 = True
    
    Select Case G_タブNO
        Case 1
            顧客情報_登録 = 顧客情報_登録_sub()
        Case 2
            顧客情報_登録 = 配送先_登録_sub()
    End Select
    
End Function

'************************************************************************
'機  能 :顧客マスタを登録する。
'************************************************************************
Private Function 顧客情報_登録_sub() As Boolean
    
    Dim 顧客ID As String
    Dim 顧客マスタ As type顧客マスタ
    Dim row As Integer
    Dim 住所 As String
        
    顧客情報_登録_sub = False
    
    On Error GoTo err
    
    If va顧客リスト.MaxRows < 1 Then
        Exit Function
    End If
    
   'If txt顧客名.Text = "" Then
   '    Call MsgBox("顧客名が未入力です", vbOKOnly, "顧客管理")
   '    顧客情報_登録_sub = True
   '    Exit Function
   'End If
    
    
    'MsgBox txt顧客ID.Text, vbOKOnly, "XXXXX"
    'Debug.Print txt顧客ID.Text
    
    With 顧客マスタ
        
        .顧客ID = txt顧客ID.Text
        .顧客名 = txt顧客名.Text
        .フリガナ = txtフリガナ.Text
        .〒 = txt郵便番号.Text
        .住所1 = txt住所_上段.Text
        .住所2 = txt住所_中段.Text
        .住所3 = txt住所_下段.Text
        .電話番号 = txt電話番号.Text
        .メール = txtメール.Text
        .楽天メール = txt楽天メール.Text
        .メール送信 = chkメール送信.Value
        .アーデルクラブ = cmbアーデルクラブ.Text
        
        If txt入会日.Text = "____/__/__" Then
            .入会日 = ""
        Else
            .入会日 = txt入会日.Text
        End If
        
        If txt退会日.Text = "____/__/__" Then
            .退会日 = ""
        Else
            .退会日 = txt退会日.Text
        End If
        
        If txt誕生日.Text = "____/__/__" Then
            .誕生日 = ""
        Else
            .誕生日 = txt誕生日.Text
        End If
        
        .性別 = IIf(opt男性.Value = True, "1", "2")
        .備考 = txt備考.Text
        .資料1 = chk資料1.Value
        .資料2 = chk資料2.Value
        .資料3 = chk資料3.Value
        .資料4 = chk資料4.Value
        .資料5 = chk資料5.Value
        .削除 = "0"
                
        If .顧客名 = "" _
            And .フリガナ = "" _
            And .〒 = "" _
            And .住所1 = "" _
            And .住所2 = "" _
            And .住所3 = "" _
            And .電話番号 = "" _
            And .メール = "" _
            And .楽天メール = "" _
            And .アーデルクラブ = "" _
            And (.入会日 = "____/__/__" Or .入会日 = "") _
            And (.退会日 = "____/__/__" Or .退会日 = "") _
            And (.誕生日 = "____/__/__" Or .誕生日 = "") _
            And .備考 = "" Then
            Exit Function
        End If
        
        If .顧客ID <> "" Then
            ' 顧客IDが採番済みの場合は、顧客データを更新する
            If 顧客マスタ更新(顧客マスタ) = False Then
                If MsgBox("他の端末で更新されているため、更新できません。" + Chr$(13) + Chr$(10) + "リロードしますか？", vbYesNo, "顧客管理") = vbYes Then
                    Call cmd未出荷一覧_Click
                End If
                Exit Function
            End If
        Else
            ' 顧客IDが未採番の場合は、顧客データを新規に登録する
            .顧客ID = 顧客マスタ登録(顧客マスタ)
            txt顧客ID.Text = .顧客ID
        End If
        
        row = G_顧客リスト_ROW
        'Call SpreadSetVal(va顧客リスト, row, COL_チェック, 0)
        Call SpreadSetVal(va顧客リスト, row, COL_顧客ID, .顧客ID)
        Call SpreadSetVal(va顧客リスト, row, COL_顧客名, .顧客名)
        Call SpreadSetVal(va顧客リスト, row, COL_フリガナ, .フリガナ)
        Call SpreadSetVal(va顧客リスト, row, COL_〒, .〒)
        
        住所 = .住所1 + .住所2 + .住所3
        Call SpreadSetVal(va顧客リスト, row, COL_住所1, 住所)
        'Call SpreadSetVal(va顧客リスト, row, COL_住所2, .住所2)
        'Call SpreadSetVal(va顧客リスト, row, COL_住所3, .住所3)
        Call SpreadSetVal(va顧客リスト, row, COL_電話番号, .電話番号)
        Call SpreadSetVal(va顧客リスト, row, COL_メール, .メール)
        Call SpreadSetVal(va顧客リスト, row, COL_アーデルクラブ, .アーデルクラブ)
        Call SpreadSetVal(va顧客リスト, row, COL_入会日, .入会日)
        Call SpreadSetVal(va顧客リスト, row, COL_備考, .備考)
        Call SpreadSetVal(va顧客リスト, row, COL_楽天メール, .楽天メール)
        txt注意喚起.Caption = .備考
            
    End With
              
    Exit Function
    
err:
    Call MsgBox("DB更新エラーにつき再起動して下さい。", vbOKOnly, "顧客管理")
End Function

'************************************************************************
'機  能 :配送先情報を登録する。
'************************************************************************
Private Function 配送先_登録_sub() As Boolean
    
    Dim 顧客ID As String
    Dim 顧客マスタ As type顧客マスタ
    Dim row As Integer
    
    On Error GoTo err
    
    配送先_登録_sub = False
    
    If va顧客リスト.MaxRows < 1 Then
        配送先_登録_sub = False
        Exit Function
    End If
    
    If txt顧客ID.Text = "" Then
        Call MsgBox("顧客情報が未入力です", vbOKOnly, "顧客管理")
        配送先_登録_sub = True
        Exit Function
    End If
    
'   If txt顧客名.Text = "" Then
'       Call MsgBox("顧客名が未入力です", vbOKOnly, "顧客管理")
'       配送先_登録_sub = True
'       Exit Function
'   End If
    
    With 顧客マスタ
        
        .顧客ID = txt顧客ID.Text
        .顧客名 = txt顧客名.Text
        .フリガナ = txtフリガナ.Text
        .〒 = txt郵便番号.Text
        .住所1 = txt住所_上段.Text
        .住所2 = txt住所_中段.Text
        .住所3 = txt住所_下段.Text
        .電話番号 = txt電話番号.Text
        .メール = txtメール.Text
        .性別 = IIf(opt男性.Value = True, "1", "2")
        .備考 = txt備考.Text
        .削除 = "0"
        
        If .顧客ID <> "" Then
            ' 顧客IDが採番済みの場合は、顧客データを更新する
            If 配送先更新(顧客マスタ) = False Then
                If MsgBox("他の端末で更新されているため、更新できません。" + Chr$(13) + Chr$(10) + "リロードしますか？", vbYesNo, "顧客管理") = vbYes Then
                    Call cmd未出荷一覧_Click
                End If
                Exit Function
            End If
        Else
            ' 顧客IDが未採番の場合は、顧客データを新規に登録する
            Call 配送先登録(顧客マスタ)
        End If
        
        row = G_顧客リスト_ROW
        Call SpreadSetVal(va顧客リスト, row, COL_お届け先名, .顧客名)
        Call SpreadSetVal(va顧客リスト, row, COL_お届け先メール, .メール)
    
    End With
    
    Exit Function
err:
    Call MsgBox("DB更新エラーにつき再起動して下さい。", vbOKOnly, "顧客管理")

End Function

'************************************************************************
'機  能　注文を更新する。
'************************************************************************
Private Function 注文_更新() As Boolean
    
    Dim i               As Long
    Dim row             As Long
    Dim 売上明細RS      As New ADODB.Recordset
    Dim 注文ID          As String
    Dim 顧客ID          As String
    Dim 顧客名          As String
    Dim 累積本数        As Integer
    Dim 配達希望日時    As String
    Dim メルマガ送信予定日  As Date
    On Error GoTo err
    
    Dim 売上明細 As type売上明細
    
    'MsgBox txt注文ID.Text, vbOKOnly, "XXXXX"
    'Debug.Print txt注文ID.Text
    
    注文_更新 = True
    
    If va顧客リスト.MaxRows < 1 Then
        注文_更新 = False
        Exit Function
    End If
    
    If va注文リスト.MaxRows < 1 Then
        注文_更新 = False
        Call cmd追加2_Click
        Exit Function
    End If
    
    注文ID = txt注文ID.Text
    
    If 注文ID = "" Then
        顧客ID = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_顧客ID)
        顧客名 = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_顧客名)
        '顧客ID = txt顧客ID.Text
        '顧客名 = txt顧客名.Text
    
        If 顧客ID = "" Then
            Call MsgBox("顧客情報が未入力です", vbOKOnly, "顧客管理")
            注文_更新 = False
            Exit Function
        End If
    End If
    
    If cmbステータス.Text = "出荷完了" Then
    
        If txt受注日.Text >= "2012/02/15" Then
        
            If txt出荷日.Text = "____/__/__" Then
                Call MsgBox("出荷日が未入力です", vbOKOnly, "顧客管理")
                注文_更新 = False
                Exit Function
            End If
        
            If txt注文番号.Text = "" Then
                Call MsgBox("注文番号が未入力です", vbOKOnly, "顧客管理")
                注文_更新 = False
                Exit Function
            End If
            
            Select Case cmb注文元.Text
                Case "楽天"
                    If Mid(txt注文番号.Text, 7, 1) = "-" And Mid(txt注文番号.Text, 16, 1) = "-" Then
                    Else
                        Call MsgBox("注文番号の形式が誤りです", vbOKOnly, "顧客管理")
                        注文_更新 = False
                        Exit Function
                    End If
                Case "Yahoo"
                    If Mid(txt注文番号.Text, 1, 6) = "adele-" Or IsNumeric(txt注文番号.Text) = True Then
                    Else
                        Call MsgBox("注文番号の形式が誤りです", vbOKOnly, "顧客管理")
                        注文_更新 = False
                        Exit Function
                    End If
                Case "レントラックス"
                    If Mid(txt注文番号.Text, 1, 1) = "R" Then
                    Else
                        Call MsgBox("注文番号の形式が誤りです", vbOKOnly, "顧客管理")
                        注文_更新 = False
                        Exit Function
                    End If
'                Case "おちゃのこネット"
'                    If Mid(txt注文番号.Text, 1, 5) = "OCNK-" Then
'                    Else
'                        Call MsgBox("注文番号の形式が誤りです", vbOKOnly, "顧客管理")
'                        注文_更新 = False
'                        Exit Function
'                    End If
                Case "アマゾン"
                    If Mid(txt注文番号.Text, 4, 1) = "-" And Mid(txt注文番号.Text, 12, 1) = "-" Then
                    Else
                        Call MsgBox("注文番号の形式が誤りです", vbOKOnly, "顧客管理")
                        注文_更新 = False
                        Exit Function
                    End If
                Case "コマチ"
                    If Mid(txt注文番号.Text, 1, 8) = "KOMACHI-" Then
                    Else
                        Call MsgBox("注文番号の形式が誤りです", vbOKOnly, "顧客管理")
                        注文_更新 = False
                        Exit Function
                    End If
                Case Else
                    If Mid(txt注文番号.Text, 1, 4) = "ETC-" Then
                    Else
                        Call MsgBox("注文番号の形式が誤りです", vbOKOnly, "顧客管理")
                        注文_更新 = False
                        Exit Function
                    End If
            End Select
            
            If CLng(txt合計金額.Text) < 0 Then
                Call MsgBox("合計金額がマイナスにならないように入力して下さい", vbOKOnly, "顧客管理")
                注文_更新 = False
                Exit Function
            End If
            
            If cmb部門.Text = "ｱｰﾃﾞﾙ" Then
                If アーデル判定(cmb商品名.Text) = 1 Or アーデル判定(cmb商品名.Text) = 9 Then
                Else
                    Call MsgBox("部門が誤っています", vbOKOnly, "顧客管理")
                    注文_更新 = False
                    Exit Function
                End If
            End If
            
            If cmb部門.Text = "その他" Then
                If アーデル判定(cmb商品名.Text) = 1 Or アーデル判定(cmb商品名.Text) = 9 Then
                    Call MsgBox("部門が誤っています", vbOKOnly, "顧客管理")
                    注文_更新 = False
                    Exit Function
                End If
            End If
            
            If cmb部門.Text = "ｺﾓﾗｲﾌ" And CLng(txt合計金額.Text) <> 0 Then
'               If txt仕入金額.Value = 0 Or txt荷造運賃.Value = 0 Then
                If txt仕入金額.Value = 0 Then
                    Call MsgBox("仕入金額／荷造運賃を入力して下さい", vbOKOnly, "顧客管理")
                    注文_更新 = False
                    Exit Function
                End If
            End If
            
            If cmb注文方法.Text = "銀行振込" Then
                If cmb銀行.Text = "" Then
                    Call MsgBox("銀行振込の場合、銀行名を入力して下さい", vbOKOnly, "顧客管理")
                    注文_更新 = False
                    Exit Function
                End If
            End If
            
            If cmb注文方法.Text = "商品代引" Then
                If アーデル判定(cmb商品名.Text) = 1 Then
                    If cmb宅配業者.Text = "佐川急便" Or cmb宅配業者.Text = "ゆうパック" Then
                    Else
                        Call MsgBox("商品代引きの場合、佐川急便 or ゆうパックを入力して下さい", vbOKOnly, "顧客管理")
                        注文_更新 = False
                        Exit Function
                    End If
                End If
            End If
        End If
    End If

    
#If 0 Then
    If cmb注文元.Text = "" Then
        Call MsgBox("注文元を選択して下さい。", vbOKOnly, "顧客管理")
        注文_更新 = False
        Exit Function
    End If
    
    If cmb商品名.Text = "" Then
        Call MsgBox("商品を選択して下さい。", vbOKOnly, "顧客管理")
        注文_更新 = False
        Exit Function
    End If
#End If
        
    With 売上明細
        
        If txt受注日.Text = "____/__/__" Then
            .受注日 = ""
        Else
            .受注日 = txt受注日.Text
        End If
        .ステータス = cmbステータス.Text
        .商品名 = cmb商品名.Text
        .部門 = cmb部門.Text
        .注文方法 = cmb注文方法.Text
        .銀行 = cmb銀行.Text
        .配達希望日時 = txt配達日時.Text
        .配達希望日時2 = txt配達日時2.Text
        .仕入金額 = txt仕入金額.Value
        .単価 = txt単価.Value
        .割引 = txt割引.Value
        .割引区分 = "\"
        .数量 = txt数量.Value
        .金額 = (txt単価.Value + txt割引.Value) * txt数量.Value
        .消費税 = 0
        .送料 = txt送料.Value
        .荷造運賃 = txt荷造運賃.Value
        .返金 = txt返金.Value
        .その他手数料 = txtその他手数料.Value
        .合計金額 = CLng(txt合計金額.Text)
        
        If txt入金日.Text = "____/__/__" Then
            .入金日 = ""
        Else
            .入金日 = txt入金日.Text
        End If
        
        If txt出荷日.Text = "____/__/__" Then
            .出荷日 = ""
        Else
            .出荷日 = txt出荷日.Text
        End If
        
        .着荷日 = ""
        .宅配業者 = cmb宅配業者.Text
        .注文元 = cmb注文元.Text
        .Yahoo注文番号 = Trim(txt注文番号.Text)
        .参照元 = ""
        .キーワード = ""
        .入力ポイント = ""
        .商品コード = ""
        .ロイヤリティー = 0
        .送付資料 = ""
        .返品対象 = ""
        .支払番号 = txt支払番号.Text
        .問合番号 = txt問合番号.Text
        .備考1 = txt備考2.Text
        .備考2 = ""
        .備考3 = ""
        .コモライフNO = Trim(txtコモライフ.Text)
        
        If txt出荷予定日.Text = "____/__/__" Then
            .出荷予定日 = ""
        Else
            .出荷予定日 = txt出荷予定日.Text
        End If

        .決済URL = txt決済URL.Text

        If 注文ID <> "" Then .注文ID = CLng(注文ID) Else .注文ID = -1
        .顧客ID = 顧客ID
        .顧客名 = 顧客名
        .メール送信 = txtメール送信.Text
        .売上抽出 = "0"
        .削除 = "0"

        If .注文ID <> -1 Then
            ' 注文IDが採番済みの場合は、注文データを更新する
            If 売上明細更新(売上明細) = False Then
                If MsgBox("他の端末で更新されているため、更新できません。" + Chr$(13) + Chr$(10) + "リロードしますか？", vbYesNo, "顧客管理") = vbYes Then
                    Call cmd未出荷一覧_Click
                End If
                Exit Function
            End If
        Else
            ' 注文IDが未採番の場合は、注文データを新規に登録する
            注文ID = 売上明細登録(売上明細)
        End If
        
        txt注文ID.Text = CStr(注文ID)
        
        If va注文リスト.MaxRows < 1 Then
            va注文リスト.MaxRows = 1
            G_注文リスト_ROW = 1
        End If
        
        row = G_注文リスト_ROW
        Call SpreadSetVal(va注文リスト, row, COL_受注日, .受注日)
        Call SpreadSetVal(va注文リスト, row, COL_ステータス, .ステータス)
        Call SpreadSetVal(va注文リスト, row, COL_商品名, .商品名)
        Call SpreadSetVal(va注文リスト, row, COL_注文方法, .注文方法)
        配達希望日時 = .配達希望日時 + " " + .配達希望日時2
        Call SpreadSetVal(va注文リスト, row, COL_配達希望日時, 配達希望日時)
        Call SpreadSetVal(va注文リスト, row, COL_単価, .単価)
        Call SpreadSetVal(va注文リスト, row, COL_割引, .割引)
        Call SpreadSetVal(va注文リスト, row, COL_数量, .数量)
        Call SpreadSetVal(va注文リスト, row, COL_金額, .金額)
        Call SpreadSetVal(va注文リスト, row, COL_送料, .送料)
        Call SpreadSetVal(va注文リスト, row, COL_返金, .返金)
        Call SpreadSetVal(va注文リスト, row, COL_その他手数料, .その他手数料)
        Call SpreadSetVal(va注文リスト, row, COL_合計金額, .合計金額)
        Call SpreadSetVal(va注文リスト, row, COL_入金日, .入金日)
        Call SpreadSetVal(va注文リスト, row, COL_出荷日, .出荷日)
        Call SpreadSetVal(va注文リスト, row, COL_着荷日, .着荷日)
        Call SpreadSetVal(va注文リスト, row, COL_宅配業者, .宅配業者)
        Call SpreadSetVal(va注文リスト, row, COL_注文元, .注文元)
        Call SpreadSetVal(va注文リスト, row, COL_Yahoo注文番号, .Yahoo注文番号)
        Call SpreadSetVal(va注文リスト, row, COL_参照元, .参照元)
        Call SpreadSetVal(va注文リスト, row, COL_キーワード, .キーワード)
        Call SpreadSetVal(va注文リスト, row, COL_入力ポイント, .入力ポイント)
        Call SpreadSetVal(va注文リスト, row, COL_商品コード, .商品コード)
        Call SpreadSetVal(va注文リスト, row, COL_ロイヤリティー, .ロイヤリティー)
        Call SpreadSetVal(va注文リスト, row, COL_送付資料, .送付資料)
        Call SpreadSetVal(va注文リスト, row, COL_返品対象, .返品対象)
        Call SpreadSetVal(va注文リスト, row, COL_支払番号, .支払番号)
        Call SpreadSetVal(va注文リスト, row, COL_問合番号, .問合番号)
        Call SpreadSetVal(va注文リスト, row, COL_備考1, .備考1)
        Call SpreadSetVal(va注文リスト, row, COL_備考2, .備考2)
        Call SpreadSetVal(va注文リスト, row, COL_備考3, .備考3)
        Call SpreadSetVal(va注文リスト, row, COL_注文ID, 注文ID)
        Call SpreadSetVal(va注文リスト, row, COL_メール送信, .メール送信)
        Call SpreadSetVal(va注文リスト, row, COL_割引区分, "円")
        Call SpreadSetVal(va注文リスト, row, COL_出荷予定日, .出荷予定日)
        Call SpreadSetVal(va注文リスト, row, COL_決済URL, .決済URL)
        
    End With
    
    注文_更新 = False
    
    txt累積数.Text = 累積数計算()
    
    If cmbステータス.Text = "出荷完了" Then
        顧客ID = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_顧客ID)
        'メルマガ送信予定日 = Format(DateAdd("d", 30, txt出荷日.Text), "yyyy/mm/dd")
        'Call メルマガ発行NO更新(顧客ID, 0, "'" + CStr(メルマガ送信予定日) + "'")
    End If
    
    Exit Function
    
err:
    Call MsgBox("DB更新エラーにつき再起動して下さい。", vbOKOnly, "顧客管理")
    
End Function

'************************************************************************
'機  能 :アーデルの累積購入数を取得する
'************************************************************************
Private Function 累積数計算() As Long
    
    Dim i           As Long
    Dim 累積本数    As Long
    Dim ステータス  As String
    Dim 商品名      As String
    Dim 数量        As String
    
    累積本数 = 0
    累積数計算 = 0
    
    With va注文リスト
        For i = 1 To .MaxRows
            ステータス = SpreadGetVal(va注文リスト, i, COL_ステータス)
            商品名 = SpreadGetVal(va注文リスト, i, COL_商品名)
            数量 = SpreadGetVal(va注文リスト, i, COL_数量)
            
            If ステータス <> "キャンセル" And ステータス <> "保留" And ステータス <> "資料請求" Then
                If 商品名 = "アーデル" Or 商品名 = "スーパーアーデル" Or 商品名 = "アーデル(セール)" Or 商品名 = "アーデル＋シャンプー" Or 商品名 = "アーデル＋シャンプー試供品" Then
                    累積本数 = 累積本数 + IIf(IsNumeric(数量), CInt(数量), 0)
                End If
                
                If 商品名 = "アーデル2本セット" Then
                    累積本数 = 累積本数 + IIf(IsNumeric(数量), CInt(数量), 0) * 2
                End If
            End If
        Next i
    End With
    
    累積数計算 = 累積本数

End Function

'************************************************************************
'機  能 :アーデルの累積購入数を取得する
'************************************************************************
Private Function 累積数計算2(ByVal ID As Long) As Long
    
    Dim i           As Long
    Dim 累積本数    As Long
    Dim ステータス  As String
    Dim 商品名      As String
    Dim 数量        As String
    Dim 注文ID      As Long
    
    累積本数 = 0
    累積数計算2 = 0
    
    With va注文リスト
        For i = 1 To .MaxRows
            注文ID = SpreadGetVal2(va注文リスト, i, COL_注文ID)
            ステータス = SpreadGetVal(va注文リスト, i, COL_ステータス)
            商品名 = SpreadGetVal(va注文リスト, i, COL_商品名)
            数量 = SpreadGetVal(va注文リスト, i, COL_数量)
            
            If 注文ID <> ID Then
                If ステータス <> "キャンセル" And ステータス <> "保留" And ステータス <> "資料請求" Then
                If 商品名 = "アーデル" Or 商品名 = "スーパーアーデル" Or 商品名 = "アーデル(セール)" Or 商品名 = "アーデル＋シャンプー" Or 商品名 = "アーデル＋シャンプー試供品" Then
                        累積本数 = 累積本数 + IIf(IsNumeric(数量), CInt(数量), 0)
                    End If
                    
                    If 商品名 = "アーデル2本セット" Then
                        累積本数 = 累積本数 + IIf(IsNumeric(数量), CInt(数量), 0) * 2
                    End If
                End If
            End If
        Next i
    End With
    
    累積数計算2 = 累積本数

End Function

'************************************************************************
'機  能 :スプレッドシートにデータを設定する
'************************************************************************
Public Sub SpreadSetVal(ByVal Spread As vaSpread, ByVal lngRow As Long, ByVal lngCol As Long, ByVal strText As String)
    With Spread
        .row = lngRow
        .Col = lngCol
        .Text = strText
    End With
End Sub

'************************************************************************
'機  能 :スプレッドシートからデータを取得する
'************************************************************************
Public Function SpreadGetVal(ByVal Spread As vaSpread, ByVal lngRow As Long, ByVal lngCol As Long) As String
    With Spread
        .row = lngRow
        .Col = lngCol
        SpreadGetVal = Trim(.Text)
    End With
End Function

'************************************************************************
'機  能 :スプレッドシートからデータを取得する
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
'機  能 :スプレッドシートのセル位置を設定する
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
'機  能 :チェックされている件数を取得する
'************************************************************************
Function チェック件数取得() As Integer
    
    Dim i As Integer
    Dim cnt As Integer
    
    cnt = 0
    
    For i = 1 To va注文リスト.MaxRows
        If SpreadGetVal(va注文リスト, i, COL_チェック) = "1" Then
            cnt = cnt + 1
        End If
    Next i
    
    チェック件数取得 = cnt
    
End Function

'************************************************************************
'機  能 :CSV出力処理を行う
'************************************************************************
Private Sub cmdCSV出力_Click()

    Dim 顧客ID          As String
    Dim 顧客名          As String
    Dim 〒              As String
    Dim 住所1           As String
    Dim 住所2           As String
    Dim 住所3           As String
    Dim 電話番号        As String
    Dim intFileNo       As Integer

    Dim CSV抽出RS As New ADODB.Recordset
    Dim 配送先RS As New ADODB.Recordset
    
    If MsgBox("住所録ＣＳＶを出力してもよろしいですか？", vbYesNo, "顧客管理") = vbNo Then
        Exit Sub
    End If

    MousePointer = vbHourglass
    
    intFileNo = FreeFile()
    Open "C:\顧客管理\佐川急便.csv" For Output As #intFileNo

    Call CSV未出力顧客マスタ読込(CSV抽出RS)

    If CSV抽出RS.EOF Then
        CSV抽出RS.Close
        Close #intFileNo
        MousePointer = vbNormal
        Call MsgBox("新規住所録データは存在しません", vbOKOnly, "顧客管理")
        Exit Sub
    End If
    
    With CSV抽出RS
        Do Until .EOF
        
        顧客名 = !顧客名
        〒 = ![〒]
        住所1 = !住所1
        住所2 = !住所2
        住所3 = IIf(IsNull(!住所3), "", !住所3)
        電話番号 = !電話番号
        
        Call 配送先1件読込(配送先RS, !顧客ID)
        
        If Not 配送先RS.EOF Then
            If 配送先RS!顧客名 <> "" Then
                顧客名 = 配送先RS!顧客名
                〒 = 配送先RS![〒]
                住所1 = 配送先RS!住所1
                住所2 = 配送先RS!住所2
                住所3 = 配送先RS!住所3
                電話番号 = 配送先RS!電話番号
            End If
        End If
        
        配送先RS.Close
        
        Print #intFileNo, CStr(CLng(!顧客ID)) & "," _
                            & 住所1 & "," _
                            & 住所2 & "," _
                            & 住所3 & "," _
                            & 顧客名 & "," _
                            & "," _
                            & 電話番号 & "," _
                            & 〒 & ",,,,,,,,,,,,,,,000,,,00,,,,,,10,00,,,,,,"
                            

        Call CSV出力フラグ更新(!顧客ID)
        CSV抽出RS.MoveNext
        Loop
        .Close
    End With
    
    Close #intFileNo
    MousePointer = vbNormal
    Call MsgBox("「C:\顧客管理\佐川急便.csv」に、住所録データを出力しました", vbOKOnly, "顧客管理")
    
End Sub

'************************************************************************
'機  能 :売上データをＣＳＶ出力する
'************************************************************************
Private Sub cmd売上_Click()
    
    Dim 会計            As type会計
    Dim 抽出日          As String
    Dim ADF018          As New ADF018

    Dim 売上抽出RS As New ADODB.Recordset
    
    If MsgBox("売上ＣＳＶを出力してもよろしいですか？", vbYesNo, "顧客管理") = vbNo Then
        Exit Sub
    End If
    
    If MsgBox("本当に作成してもよろしいですか？", vbYesNo, "顧客管理") = vbNo Then
        Exit Sub
    End If
    
    'Call ADF018.Show(1)

    '抽出日 = ADF018.抽出日取得()

    MousePointer = vbHourglass
    
    Call 売上データ削除
    
    'Call 未出力売上データ読込(売上抽出RS, 抽出日)
    Call 未出力売上データ読込(売上抽出RS)

    If 売上抽出RS.EOF Then
        売上抽出RS.Close
        MousePointer = vbNormal
        Call MsgBox("新規売上ＣＳＶデータは存在しません", vbOKOnly, "顧客管理")
        Exit Sub
    End If
    
    
    会計.識別フラグ = "11"
    会計.伝票NO = 0                         '伝票番号取得()                    ' """"""
'    会計.決算 = """"""
    会計.取引日時 = ""
    
    会計.タイプ = "3"
    会計.生成元 = "振伝"
    
    ' 売上仮処理
    Call 売上処理(会計, 売上抽出RS, False)
    
    売上抽出RS.Close
    
    ' 借方金額と貸方金額をチェックする
    If 売上データチェック() = False Then
        MousePointer = vbNormal
        Call MsgBox("借方金額と貸方金額が合いません", vbOKOnly, "顧客管理")
        Exit Sub
    End If
    
    Call 売上データ削除
    
    'Call 未出力売上データ読込(売上抽出RS, 抽出日)
    Call 未出力売上データ読込(売上抽出RS)
    
    
    ' 売上本処理
    Call 売上処理(会計, 売上抽出RS, True)
    
    売上抽出RS.Close
    
    Call 売上データ削除2
    
    Call 売上データコピー
    
    Call 識別フラグ設定
    
    If G_店舗名 = "トリニティー楽天市場店" Then
        Call 楽天_注文ステータス変更
    Else
        Call Yahoo_注文ステータス変更
    End If
    
    Call 会計CSV出力

    MousePointer = vbNormal

    If G_店舗名 = "トリニティー楽天市場店" Then
        Call MsgBox("「C:\顧客管理\楽天_売上.csv」に、データを出力しました", vbOKOnly, "顧客管理")
    Else
        Call MsgBox("「C:\顧客管理\Yahoo_売上.csv」に、データを出力しました", vbOKOnly, "顧客管理")
    End If
    
End Sub

'************************************************************************
'機  能 :売上処理
'************************************************************************
Private Sub 売上処理(ByRef 会計 As type会計, ByVal 売上抽出RS As ADODB.Recordset, ByVal 本番区分 As Boolean)
    
    Dim 区分            As Integer
    
    With 売上抽出RS
    
        Do Until .EOF
            ' コモライフの場合、１件づつ注文をばらかせる。
            ' そうしないと、複数注文が発生した場合、仕入れと、荷造運賃が最初に寄ってしまう
            '
            If !部門 <> "ｺﾓﾗｲﾌ" Then
                会計.注文番号 = !Yahoo注文番号
            Else
                会計.注文番号 = !Yahoo注文番号 & "#" & !注文ID
            End If
            
            会計.注文ID = !注文ID
            会計.取引日時 = !出荷日
            
            If !部門 = "ｱｰﾃﾞﾙ" Then
                区分 = 1
            ElseIf !部門 = "ｺﾓﾗｲﾌ" Then
                区分 = 2
            ElseIf !部門 = "その他" Then
                区分 = 3
            Else
                If アーデル判定(!商品名) = 1 Then
                    区分 = 1
                Else
                    区分 = 2
                End If
            End If
            
            '
            ' アーデル売上出力
            '
            If 区分 = 1 Then
                If !合計金額 > 0 Then
                    If !注文方法 = "銀行振込" Then
                        Call 現金_通常出力(会計, 売上抽出RS)
                    ElseIf !注文方法 = "楽天バンク決済" Then
                        Call 楽天バンク_通常出力(会計, 売上抽出RS)
                    Else
                        Call 売掛金_通常出力(会計, 売上抽出RS)
                    End If
                Else
                    If !合計金額 = 0 And !その他手数料 < 0 Then
                        Call 売掛金_ポイント出力(会計, 売上抽出RS)
                    End If
                End If
                
            '
            ' コモライフ売上出力
            '
            ElseIf 区分 = 2 Then
                If !合計金額 > 0 Then
                    If !注文方法 = "銀行振込" Then
                        Call 現金_コモライフ出力(会計, 売上抽出RS)
                    ElseIf !注文方法 = "楽天バンク決済" Then
                        Call 楽天バンク_コモライフ出力(会計, 売上抽出RS)
                    Else
                        Call 売掛金_コモライフ出力(会計, 売上抽出RS)
                    End If
                Else
                    If !合計金額 = 0 And !その他手数料 < 0 Then
                        Call 売掛金_ポイント_コモライフ出力(会計, 売上抽出RS)
                    End If
                End If
                
                If !金額 > 0 Then
                    Call 買掛金_出力(会計, 売上抽出RS)
                    
                    Call 荷造運賃_出力(会計, 売上抽出RS)
                End If
            '
            ' その他出力
            '
            Else
                If !合計金額 > 0 Then
                    If !注文方法 = "銀行振込" Then
                        Call 現金_その他出力(会計, 売上抽出RS)
                    ElseIf !注文方法 = "楽天バンク決済" Then
                        Call 楽天バンク_その他出力(会計, 売上抽出RS)
                    Else
                        Call 売掛金_その他出力(会計, 売上抽出RS)
                    End If
                Else
                    If !合計金額 = 0 And !その他手数料 < 0 Then
                        Call 売掛金_ポイント_その他出力(会計, 売上抽出RS)
                    End If
                End If
                
            End If
            
            If 本番区分 = True Then
                Call 売上出力フラグ更新(!注文ID)
            End If
            
            .MoveNext
        Loop
    End With

End Sub

'************************************************************************
'機  能 :売掛金出力（通常売上）
'************************************************************************
Private Sub 売掛金_通常出力(ByRef k As type会計, ByVal u As ADODB.Recordset)
    
    k.借方勘定科目 = "売掛金"
'    If u!注文元 = "Yahoo" Or u!注文元 = "楽天" Or u!注文元 = "自社サイト" Or u!注文元 = "おちゃのこネット" Or u!注文元 = "コマチ" Then
    If u!注文元 = "Yahoo" Or u!注文元 = "楽天" Or u!注文元 = "自社サイト" Or u!注文元 = "レントラックス" Or u!注文元 = "コマチ" Then
        k.借方補助科目 = 借方補助科目取得1(u!注文方法, u!宅配業者)
    Else
        k.借方補助科目 = u!注文元
    End If
    k.借方部門 = "全社"
    k.借方税区分 = "対象外"
    k.借方金額 = u!合計金額 - u!送料
    k.借方税金額 = 0
    
    k.貸方勘定科目 = "売上高"
    If u!注文元 = "Yahoo" Or u!注文元 = "楽天" Or u!注文元 = "自社サイト" Then
        k.貸方補助科目 = 貸方補助科目取得1(u!商品名)
    Else
        k.貸方補助科目 = u!注文元
    End If
    
    k.貸方部門 = "全社"
    k.貸方税区分 = G_売内
    k.貸方金額 = u!合計金額 + u!その他手数料 * -1 - u!送料
    k.貸方税金額 = 消費税計算(k.貸方金額)
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)
    
    If u!送料 > 0 Then
        Call 荷造運賃_出力2(k, u)
    End If
    
    If u!返金 < 0 Then
        k.借方勘定科目 = "売掛金"
        k.借方補助科目 = 借方補助科目取得1(u!注文方法, u!宅配業者)
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!返金 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = "現金"
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = u!返金 * -1
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
    
        Call 会計出力(k)
    
    End If
    
    If u!その他手数料 < 0 Then
        k.借方勘定科目 = "売掛金"
        k.借方補助科目 = "ポイント"
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!その他手数料 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = ""
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = 0
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
        
        Call 会計出力(k)
        
    End If

End Sub

'************************************************************************
'機  能 :現金出力
'************************************************************************
Private Sub 現金_通常出力(ByRef k As type会計, ByVal u As ADODB.Recordset)
                    
    k.借方勘定科目 = "普通預金"
    k.借方補助科目 = u!銀行
    k.借方部門 = "全社"
    k.借方税区分 = "対象外"
    k.借方金額 = u!合計金額 - u!送料
    k.借方税金額 = 0
    
    k.貸方勘定科目 = "売上高"
    
    If u!注文元 = "Yahoo" Or u!注文元 = "楽天" Or u!注文元 = "自社サイト" Then
        k.貸方補助科目 = 貸方補助科目取得1(u!商品名)
    Else
        k.貸方補助科目 = u!注文元
    End If
    
    k.貸方部門 = "全社"
    k.貸方税区分 = G_売内
    k.貸方金額 = u!合計金額 + u!その他手数料 * -1 - u!送料
    k.貸方税金額 = 消費税計算(k.貸方金額)
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)
    
    If u!送料 > 0 Then
        Call 荷造運賃_出力2(k, u)
    End If

    If u!返金 < 0 Then
        k.借方勘定科目 = "普通預金"
        k.借方補助科目 = u!銀行
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!返金 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = "現金"
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = u!返金 * -1
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
    
        Call 会計出力(k)
    
    End If
    
    If u!その他手数料 < 0 Then
        k.借方勘定科目 = "売掛金"
        k.借方補助科目 = "ポイント"
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!その他手数料 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = ""
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = 0
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
        
        Call 会計出力(k)
        
    End If

End Sub


'************************************************************************
'機  能 :楽天バンク出力
'************************************************************************
Private Sub 楽天バンク_通常出力(ByRef k As type会計, ByVal u As ADODB.Recordset)
                    
    k.借方勘定科目 = "普通預金"
    k.借方補助科目 = "楽天銀行"
    k.借方部門 = "全社"
    k.借方税区分 = "対象外"
    k.借方金額 = u!合計金額 - u!送料 - G_振り込手数料
    k.借方税金額 = 0
    
    k.貸方勘定科目 = "売上高"
    
    If u!注文元 = "Yahoo" Or u!注文元 = "楽天" Or u!注文元 = "自社サイト" Then
        k.貸方補助科目 = 貸方補助科目取得1(u!商品名)
    Else
        k.貸方補助科目 = u!注文元
    End If
    
    k.貸方部門 = "全社"
    k.貸方税区分 = G_売内
    k.貸方金額 = u!合計金額 + u!その他手数料 * -1 - u!送料
    k.貸方税金額 = 消費税計算(k.貸方金額)
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)
        
    If u!送料 > 0 Then
        Call 荷造運賃_出力2(k, u)
    End If

    If u!返金 < 0 Then
        k.借方勘定科目 = "普通預金"
        k.借方補助科目 = "楽天銀行"
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!返金 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = "現金"
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = u!返金 * -1
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
    
        Call 会計出力(k)
    
    End If
    
    If u!その他手数料 < 0 Then
        k.借方勘定科目 = "売掛金"
        k.借方補助科目 = "ポイント"
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!その他手数料 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = ""
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = 0
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
        
        Call 会計出力(k)
        
    End If

    k.借方勘定科目 = "支払手数料"
    k.借方補助科目 = "振り込手数料"
    k.借方部門 = "全社"
    k.借方税区分 = G_仕内
    k.借方金額 = G_振り込手数料
    k.借方税金額 = 消費税計算(G_振り込手数料)
    
    k.貸方勘定科目 = ""
    k.貸方補助科目 = ""
    k.貸方部門 = "全社"
    k.貸方税区分 = "対象外"
    k.貸方金額 = 0
    k.貸方税金額 = 0
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)

End Sub

'************************************************************************
'機  能 :売掛金出力（ポイント）
'************************************************************************
Private Sub 売掛金_ポイント出力(k As type会計, ByVal u As ADODB.Recordset)
                    
    k.借方勘定科目 = "売掛金"
    k.借方補助科目 = "ポイント"
    k.借方部門 = "全社"
    k.借方税区分 = "対象外"
    k.借方金額 = u!その他手数料 * -1 - u!送料
    k.借方税金額 = 0
    
    k.貸方勘定科目 = "売上高"
    
    If u!注文元 = "Yahoo" Or u!注文元 = "楽天" Or u!注文元 = "自社サイト" Then
        k.貸方補助科目 = 貸方補助科目取得1(u!商品名)
    Else
        k.貸方補助科目 = u!注文元
    End If
    
    k.貸方部門 = "全社"
    k.貸方税区分 = G_売内
    k.貸方金額 = u!その他手数料 * -1 - u!送料
    k.貸方税金額 = 消費税計算(k.貸方金額)
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)
        
    If u!送料 > 0 Then
        Call 荷造運賃_出力2(k, u)
    End If

    If u!返金 < 0 Then
        k.借方勘定科目 = "売掛金"
        k.借方補助科目 = "ポイント"
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!返金 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = "現金"
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = u!返金 * -1
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
    
        Call 会計出力(k)
    
    End If

End Sub

'************************************************************************
'機  能 :売掛金出力（コモライフ売上）
'************************************************************************
Private Sub 売掛金_コモライフ出力(ByRef k As type会計, ByVal u As ADODB.Recordset)
    
    k.借方勘定科目 = "売掛金"
    If u!注文元 = "Yahoo" Or u!注文元 = "楽天" Or u!注文元 = "自社サイト" Then
        k.借方補助科目 = 借方補助科目取得2(u!注文方法, u!宅配業者)
    Else
        k.借方補助科目 = u!注文元
    End If
    
    k.借方部門 = "全社"
    k.借方税区分 = "対象外"
    k.借方金額 = u!合計金額 '+ u!その他手数料
    k.借方税金額 = 0
    
    k.貸方勘定科目 = "売上高"
    k.貸方補助科目 = "こもらいふ"
    k.貸方部門 = "全社"
    k.貸方税区分 = G_売内
    k.貸方金額 = u!合計金額 - u!送料 + u!その他手数料 * -1
    k.貸方税金額 = 消費税計算(k.貸方金額)
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)
    
    If u!返金 < 0 Then
        k.借方勘定科目 = "売掛金"
        k.借方補助科目 = 借方補助科目取得2(u!注文方法, u!宅配業者)
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!返金 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = "現金"
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = u!返金 * -1
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
    
        Call 会計出力(k)
    
    End If
    
    If u!その他手数料 < 0 Then
        k.借方勘定科目 = "売掛金"
        k.借方補助科目 = "ポイント"
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!その他手数料 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = ""
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = 0
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
        
        Call 会計出力(k)
    End If

End Sub

'************************************************************************
'機  能 :売掛金出力（コモライフポイント）
'************************************************************************
Private Sub 売掛金_ポイント_コモライフ出力(ByRef k As type会計, ByVal u As ADODB.Recordset)
                    
    k.借方勘定科目 = "売掛金"
    k.借方補助科目 = "ポイント"
    k.借方部門 = "全社"
    k.借方税区分 = "対象外"
    k.借方金額 = u!その他手数料 * -1
    k.借方税金額 = 0
    
    k.貸方勘定科目 = "売上高"
    k.貸方補助科目 = "こもらいふ"
    k.貸方部門 = "全社"
    k.貸方税区分 = G_売内
    k.貸方金額 = u!その他手数料 * -1 - u!送料
    k.貸方税金額 = 消費税計算(k.貸方金額)
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)
    
    If u!返金 < 0 Then
        k.借方勘定科目 = "売掛金"
        k.借方補助科目 = "ポイント"
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!返金 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = "現金"
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = u!返金 * -1
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
    
        Call 会計出力(k)
    
    End If
    
End Sub

'************************************************************************
'機  能 :現金出力（コモライフ売上）
'************************************************************************
Private Sub 現金_コモライフ出力(ByRef k As type会計, ByVal u As ADODB.Recordset)
                    
    k.借方勘定科目 = "普通預金"
    k.借方補助科目 = u!銀行
    k.借方部門 = "全社"
    k.借方税区分 = "対象外"
    k.借方金額 = u!合計金額 '+ u!その他手数料
    k.借方税金額 = 0
    
    k.貸方勘定科目 = "売上高"
    k.貸方補助科目 = "こもらいふ"
    k.貸方部門 = "全社"
    k.貸方税区分 = G_売内
    k.貸方金額 = u!合計金額 - u!送料 + u!その他手数料 * -1
    k.貸方税金額 = 消費税計算(k.貸方金額)
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)
    
    If u!返金 < 0 Then
        k.借方勘定科目 = "普通預金"
        k.借方補助科目 = u!銀行
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!返金 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = "現金"
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = u!返金 * -1
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
    
        Call 会計出力(k)
    
    End If
    
    If u!その他手数料 < 0 Then
        k.借方勘定科目 = "売掛金"
        k.借方補助科目 = "ポイント"
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!その他手数料 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = ""
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = 0
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
        
        Call 会計出力(k)
    End If

End Sub

'************************************************************************
'機  能 :楽天バンク出力（コモライフ売上）
'************************************************************************
Private Sub 楽天バンク_コモライフ出力(ByRef k As type会計, ByVal u As ADODB.Recordset)

    k.借方勘定科目 = "普通預金"
    k.借方補助科目 = "楽天銀行"
    k.借方部門 = "全社"
    k.借方税区分 = "対象外"
    k.借方金額 = u!合計金額 - G_振り込手数料
    k.借方税金額 = 0
    
    k.貸方勘定科目 = "売上高"
    k.貸方補助科目 = "こもらいふ"
    k.貸方部門 = "全社"
    k.貸方税区分 = G_売内
    k.貸方金額 = u!合計金額 - u!送料 + u!その他手数料 * -1
    k.貸方税金額 = 消費税計算(k.貸方金額)
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)
    
    If u!返金 < 0 Then
        k.借方勘定科目 = "普通預金"
        k.借方補助科目 = u!銀行
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!返金 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = "現金"
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = u!返金 * -1
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
    
        Call 会計出力(k)
    
    End If
    
    If u!その他手数料 < 0 Then
        k.借方勘定科目 = "売掛金"
        k.借方補助科目 = "ポイント"
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!その他手数料 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = ""
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = 0
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
        
        Call 会計出力(k)
    End If

    k.借方勘定科目 = "支払手数料"
    k.借方補助科目 = "振り込手数料"
    k.借方部門 = "全社"
    k.借方税区分 = G_仕内
    k.借方金額 = G_振り込手数料
    k.借方税金額 = 消費税計算(G_振り込手数料)
    
    k.貸方勘定科目 = ""
    k.貸方補助科目 = ""
    k.貸方部門 = "全社"
    k.貸方税区分 = "対象外"
    k.貸方金額 = 0
    k.貸方税金額 = 0
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)
End Sub
'************************************************************************
'機  能 :売掛金出力（その他売上）
'************************************************************************
Private Sub 売掛金_その他出力(ByRef k As type会計, ByVal u As ADODB.Recordset)
                    
    k.借方勘定科目 = "売掛金"
    k.借方補助科目 = 借方補助科目取得1(u!注文方法, u!宅配業者)
    k.借方部門 = "全社"
    k.借方税区分 = "対象外"
    k.借方金額 = u!合計金額 - u!送料
    k.借方税金額 = 0
    
    k.貸方勘定科目 = "売上高"
    If k.借方補助科目 = "ヤフオク" Then
        k.貸方補助科目 = "ヤフオク"
    Else
        k.貸方補助科目 = "その他"
    End If
    k.貸方部門 = "全社"
    k.貸方税区分 = G_売内
    k.貸方金額 = u!合計金額 + u!その他手数料 * -1 - u!送料
    k.貸方税金額 = 消費税計算(k.貸方金額)
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)
    
    If u!送料 > 0 Then
        Call 荷造運賃_出力2(k, u)
    End If
    
    If u!返金 < 0 Then
        k.借方勘定科目 = "売掛金"
        k.借方補助科目 = 借方補助科目取得1(u!注文方法, u!宅配業者)
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!返金 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = "現金"
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = u!返金 * -1
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
    
        Call 会計出力(k)
    
    End If
    
    If u!その他手数料 < 0 Then
        k.借方勘定科目 = "売掛金"
        k.借方補助科目 = "ポイント"
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!その他手数料 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = ""
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = 0
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
        
        Call 会計出力(k)
        
    End If

End Sub

'************************************************************************
'機  能 :現金出力（その他）
'************************************************************************
Private Sub 現金_その他出力(ByRef k As type会計, ByVal u As ADODB.Recordset)
                    
    k.借方勘定科目 = "普通預金"
    k.借方補助科目 = u!銀行
    k.借方部門 = "全社"
    k.借方税区分 = "対象外"
    k.借方金額 = u!合計金額 - u!送料
    k.借方税金額 = 0
    
    k.貸方勘定科目 = "売上高"
    k.貸方補助科目 = "その他"
    k.貸方部門 = "全社"
    k.貸方税区分 = G_売内
    k.貸方金額 = u!合計金額 + u!その他手数料 * -1 - u!送料
    k.貸方税金額 = 消費税計算(k.貸方金額)
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)
    
    If u!送料 > 0 Then
        Call 荷造運賃_出力2(k, u)
    End If

    If u!返金 < 0 Then
        k.借方勘定科目 = "普通預金"
        k.借方補助科目 = u!銀行
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!返金 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = "現金"
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = u!返金 * -1
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
    
        Call 会計出力(k)
    
    End If
    
    If u!その他手数料 < 0 Then
        k.借方勘定科目 = "売掛金"
        k.借方補助科目 = "ポイント"
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!その他手数料 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = ""
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = 0
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
        
        Call 会計出力(k)
        
    End If

End Sub


'************************************************************************
'機  能 :楽天バンク出力（その他）
'************************************************************************
Private Sub 楽天バンク_その他出力(ByRef k As type会計, ByVal u As ADODB.Recordset)
                    
    k.借方勘定科目 = "普通預金"
    k.借方補助科目 = "楽天銀行"
    k.借方部門 = "全社"
    k.借方税区分 = "対象外"
    k.借方金額 = u!合計金額 - u!送料 - G_振り込手数料
    k.借方税金額 = 0
    
    k.貸方勘定科目 = "売上高"
    k.貸方補助科目 = "その他"
    k.貸方部門 = "全社"
    k.貸方税区分 = G_売内
    k.貸方金額 = u!合計金額 + u!その他手数料 * -1 - u!送料
    k.貸方税金額 = 消費税計算(k.貸方金額)
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)
        
    If u!送料 > 0 Then
        Call 荷造運賃_出力2(k, u)
    End If

    If u!返金 < 0 Then
        k.借方勘定科目 = "普通預金"
        k.借方補助科目 = "楽天銀行"
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!返金 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = "現金"
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = u!返金 * -1
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
    
        Call 会計出力(k)
    
    End If
    
    If u!その他手数料 < 0 Then
        k.借方勘定科目 = "売掛金"
        k.借方補助科目 = "ポイント"
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!その他手数料 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = ""
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = 0
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
        
        Call 会計出力(k)
        
    End If

    k.借方勘定科目 = "支払手数料"
    k.借方補助科目 = "振り込手数料"
    k.借方部門 = "全社"
    k.借方税区分 = G_仕内
    k.借方金額 = G_振り込手数料
    k.借方税金額 = 消費税計算(G_振り込手数料)
    
    k.貸方勘定科目 = ""
    k.貸方補助科目 = ""
    k.貸方部門 = "全社"
    k.貸方税区分 = "対象外"
    k.貸方金額 = 0
    k.貸方税金額 = 0
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)

End Sub

'************************************************************************
'機  能 :売掛金出力（その他ポイント）
'************************************************************************
Private Sub 売掛金_ポイント_その他出力(k As type会計, ByVal u As ADODB.Recordset)
                    
    k.借方勘定科目 = "売掛金"
    k.借方補助科目 = "ポイント"
    k.借方部門 = "全社"
    k.借方税区分 = "対象外"
    k.借方金額 = u!その他手数料 * -1 - u!送料
    k.借方税金額 = 0
    
    k.貸方勘定科目 = "売上高"
    k.貸方補助科目 = "その他"
    k.貸方部門 = "全社"
    k.貸方税区分 = G_売内
    k.貸方金額 = u!その他手数料 * -1 - u!送料
    k.貸方税金額 = 消費税計算(k.貸方金額)
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)
        
    If u!送料 > 0 Then
        Call 荷造運賃_出力2(k, u)
    End If

    If u!返金 < 0 Then
        k.借方勘定科目 = "売掛金"
        k.借方補助科目 = "ポイント"
        k.借方部門 = "全社"
        k.借方税区分 = "対象外"
        k.借方金額 = u!返金 * -1
        k.借方税金額 = 0
        
        k.貸方勘定科目 = "現金"
        k.貸方補助科目 = ""
        k.貸方部門 = "全社"
        k.貸方税区分 = "対象外"
        k.貸方金額 = u!返金 * -1
        k.貸方税金額 = 0
        
        k.摘要 = u!顧客名
    
        Call 会計出力(k)
    
    End If

End Sub

'************************************************************************
'機  能 :売掛金出力（インフォトップ）
'************************************************************************
Private Sub 売掛金_インフォトップ出力(ByRef k As type会計, ByVal u As ADODB.Recordset)
                    
    k.借方勘定科目 = "売掛金"
    k.借方補助科目 = "インフォトップ"
    k.借方部門 = "全社"
    k.借方税区分 = "対象外"
    k.借方金額 = u!合計金額 + u!送料 * -1
    k.借方税金額 = 0
    
    k.貸方勘定科目 = "売上高"
    k.貸方補助科目 = 貸方補助科目取得1(u!商品名)
    k.貸方部門 = "全社"
    k.貸方税区分 = G_売内
    k.貸方金額 = u!合計金額 + u!その他手数料 * -1 + u!送料 * -1
    k.貸方税金額 = 消費税計算(k.貸方金額)
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)
    
    k.借方勘定科目 = "売掛金"
    k.借方補助科目 = "インフォトップ"
    k.借方部門 = "全社"
    k.借方税区分 = "対象外"
    k.借方金額 = u!送料
    k.借方税金額 = 0
    
    k.貸方勘定科目 = "荷造運賃発送費"
    k.貸方補助科目 = ""
    k.貸方部門 = "全社"
    k.貸方税区分 = G_仕内
    k.貸方金額 = u!送料
    k.貸方税金額 = 消費税計算(u!送料)
    
    k.摘要 = u!顧客名

    Call 会計出力(k)

End Sub

'************************************************************************
'機  能 :売掛金出力
'************************************************************************
Private Sub 売掛金_レントラックス出力(ByRef k As type会計, ByVal u As ADODB.Recordset)
                    
    k.借方勘定科目 = "売掛金"
    k.借方補助科目 = "レントラックス"
'    k.借方補助科目 = "おちゃのこネット"
    k.借方部門 = "全社"
    k.借方税区分 = "対象外"
    k.借方金額 = u!合計金額 + u!送料 * -1
    k.借方税金額 = 0
    
    k.貸方勘定科目 = "売上高"
    k.貸方補助科目 = 貸方補助科目取得1(u!商品名)
    k.貸方部門 = "全社"
    k.貸方税区分 = G_売内
    k.貸方金額 = u!合計金額 + u!その他手数料 * -1 + u!送料 * -1
    k.貸方税金額 = 消費税計算(k.貸方金額)
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)
    
    k.借方勘定科目 = "売掛金"
    k.借方補助科目 = "レントラックス"
'    k.借方補助科目 = "おちゃのこネット"
    k.借方部門 = "全社"
    k.借方税区分 = "対象外"
    k.借方金額 = u!送料
    k.借方税金額 = 0
    
    k.貸方勘定科目 = "荷造運賃発送費"
    k.貸方補助科目 = ""
    k.貸方部門 = "全社"
    k.貸方税区分 = G_仕内
    k.貸方金額 = u!送料
    k.貸方税金額 = 消費税計算(u!送料)
    
    k.摘要 = u!顧客名

    Call 会計出力(k)

End Sub


'************************************************************************
'機  能 :買掛金出力
'************************************************************************
Private Sub 買掛金_出力(ByRef k As type会計, ByVal u As ADODB.Recordset)
                    
    k.借方勘定科目 = "仕入高"
    If G_店舗名 = "トリニティー楽天市場店" Then
        k.借方補助科目 = "こもらいふ"
    Else
        k.借方補助科目 = "コモライフ"
    End If
    k.借方部門 = "全社"
    k.借方税区分 = G_仕内
    k.借方金額 = u!仕入金額
    k.借方税金額 = 消費税計算(k.借方金額)
    
    k.貸方勘定科目 = "買掛金"
    k.貸方補助科目 = "こもらいふ"
    k.貸方部門 = "全社"
    k.貸方税区分 = "対象外"
    k.貸方金額 = u!仕入金額 + u!荷造運賃
    k.貸方税金額 = 0
    
    k.摘要 = u!顧客名
    
    Call 会計出力(k)

End Sub

'************************************************************************
'機  能 :荷造運賃出力
'************************************************************************
Private Sub 荷造運賃_出力(ByRef k As type会計, ByVal u As ADODB.Recordset)
                    
    k.借方勘定科目 = "荷造運賃発送費"
    k.借方補助科目 = ""
    k.借方部門 = "全社"
    k.借方税区分 = G_仕内
    k.借方金額 = u!荷造運賃
    k.借方税金額 = 消費税計算(k.借方金額)
    
    k.貸方勘定科目 = "荷造運賃発送費"
    k.貸方補助科目 = ""
    k.貸方部門 = "全社"
    k.貸方税区分 = G_仕内
    k.貸方金額 = u!送料
    k.貸方税金額 = 消費税計算(k.貸方金額)
    
    k.摘要 = ""
    
    Call 会計出力(k)

End Sub

'************************************************************************
'機  能 :荷造運賃出力
'************************************************************************
Private Sub 荷造運賃_出力2(ByRef k As type会計, ByVal u As ADODB.Recordset)
                    
'    k.借方勘定科目 = "荷造運賃発送費"
'    k.借方補助科目 = ""
    k.借方部門 = "全社"
    k.借方税区分 = "対象外"
    k.借方金額 = u!送料
    k.借方税金額 = 0
    
    k.貸方勘定科目 = "荷造運賃発送費"
    k.貸方補助科目 = ""
    k.貸方部門 = "全社"
    k.貸方税区分 = G_仕内
    k.貸方金額 = u!送料
    k.貸方税金額 = 消費税計算(k.貸方金額)
    
    k.摘要 = ""
    
    Call 会計出力(k)

End Sub


'************************************************************************
'機  能 :消費税計算
'************************************************************************
Private Function 消費税計算(ByVal 金額 As Long)
    
    Dim 金額2       As Double
    Dim 商品代金    As Long
    Dim 消費税      As Long
   
    金額2 = 金額
    商品代金 = CLng(Format(CStr((金額2 / (G_消費税 + 1))), "0000000000"))
    
    消費税計算 = 金額 - 商品代金
    
End Function

'************************************************************************
'機  能 :借方補助科目取得
'************************************************************************
Private Function 借方補助科目取得1(ByVal 注文方法 As String, ByVal 宅配業者 As String) As String
    
    借方補助科目取得1 = "その他"
    
    ' 注文方法が「クレジット」の場合
    If 注文方法 = "クレジット" Then
        借方補助科目取得1 = "クレジット"
        Exit Function
    End If
    
    ' 注文方法が「東京クレジット」の場合
    If 注文方法 = "東京クレジット" Then
        借方補助科目取得1 = "東京クレジット"
        Exit Function
    End If
                
    ' 注文方法が「商品代引き」の場合
    If 注文方法 = "商品代引" Then
        If 宅配業者 = "佐川急便" Then
            借方補助科目取得1 = "佐川急便"
        Else
            借方補助科目取得1 = "ゆうパック"
        End If
        Exit Function
    End If
    
    ' 注文方法が「ポイント」の場合
    If 注文方法 = "ポイント" Then
        借方補助科目取得1 = "ポイント"
        Exit Function
    End If
    
    ' 注文方法が「後払い」の場合
    If 注文方法 = "後払い" Then
        借方補助科目取得1 = "後払い"
        Exit Function
    End If
    
    ' 注文方法が「ペイジー」の場合
    If 注文方法 = "ペイジー" Then
        借方補助科目取得1 = "ペイジー"
        Exit Function
    End If
    
    ' 注文方法が「コンビニ」の場合
    If 注文方法 = "コンビニ" Then
        借方補助科目取得1 = "コンビニ"
        Exit Function
    End If
    
    ' 注文方法が「携帯決済」の場合
    If 注文方法 = "携帯決済" Then
        借方補助科目取得1 = "携帯決済"
        Exit Function
    End If
    
    ' 注文方法が「ヤフオク」の場合
    If 注文方法 = "ヤフオク" Then
        借方補助科目取得1 = "ヤフオク"
        Exit Function
    End If
    
    ' 注文方法が「電子マネー」の場合
    If 注文方法 = "電子マネー" Then
        借方補助科目取得1 = "電子マネー"
        Exit Function
    End If
    
End Function

'************************************************************************
'機  能 :借方補助科目取得
'************************************************************************
Private Function 借方補助科目取得2(ByVal 注文方法 As String, ByVal 宅配業者 As String) As String
    
    借方補助科目取得2 = "その他"
    
    ' 注文方法が「クレジット」の場合
    If 注文方法 = "クレジット" Then
        借方補助科目取得2 = "クレジット"
        Exit Function
    End If
                
    ' 注文方法が「東京クレジット」の場合
    If 注文方法 = "東京クレジット" Then
        借方補助科目取得2 = "東京クレジット"
        Exit Function
    End If
                
    ' 注文方法が「商品代引き」の場合
    If 注文方法 = "商品代引" Then
        借方補助科目取得2 = "こもらいふ"
        Exit Function
    End If
    
    ' 注文方法が「ポイント」の場合
    If 注文方法 = "ポイント" Then
        借方補助科目取得2 = "ポイント"
        Exit Function
    End If
    
    ' 注文方法が「後払い」の場合
    If 注文方法 = "後払い" Then
        借方補助科目取得2 = "後払い"
        Exit Function
    End If
    
    ' 注文方法が「ペイジー」の場合
    If 注文方法 = "ペイジー" Then
        借方補助科目取得2 = "ペイジー"
        Exit Function
    End If
    
    ' 注文方法が「コンビニ」の場合
    If 注文方法 = "コンビニ" Then
        借方補助科目取得2 = "コンビニ"
        Exit Function
    End If
    
    ' 注文方法が「携帯決済」の場合
    If 注文方法 = "携帯決済" Then
        借方補助科目取得2 = "携帯決済"
        Exit Function
    End If
    
    ' 注文方法が「ヤフオク」の場合
    If 注文方法 = "ヤフオク" Then
        借方補助科目取得2 = "ヤフオク"
        Exit Function
    End If
    
    ' 注文方法が「電子マネー」の場合
    If 注文方法 = "電子マネー" Then
        借方補助科目取得2 = "電子マネー"
        Exit Function
    End If
    
End Function

'************************************************************************
'機  能 :借方補助科目取得
'************************************************************************
Private Function 貸方補助科目取得1(ByVal 商品名 As String) As String

    貸方補助科目取得1 = 商品名
    
    If 商品名 = "アーデル＋シャンプー" Then
        If G_店舗名 = "トリニティー楽天市場店" Then
            貸方補助科目取得1 = "セット物"
        Else
            貸方補助科目取得1 = "セット"
        End If
    End If
    
    If 商品名 = "アーデル2本セット" Then
        貸方補助科目取得1 = "アーデル"
    End If
    
    If 商品名 = "シャンプー2本セット" Then
        貸方補助科目取得1 = "シャンプー"
    End If
    
    If 商品名 = "アーデル＆シャンプー試供品" Then
        貸方補助科目取得1 = "試供品"
    End If
    
    If 商品名 = "モイストリッチ クレンジング" Or _
       商品名 = "モイストリッチ ウォッシング" Or _
       商品名 = "モイストリッチ ローション" Or _
       商品名 = "モイストリッチ ジェル" Or _
       商品名 = "モイストリッチ ロイヤルエッセンス" Or _
       商品名 = "モイストリッチ 基礎化粧品セット" Then
       貸方補助科目取得1 = "ﾓｲｽﾄﾘｯﾁ"
    End If

End Function

'************************************************************************
'機  能 :アーデル製品判定
'************************************************************************
Private Function アーデル判定(ByVal 商品名 As String) As Integer
    
    アーデル判定 = 2
    
    If 商品名 = "アーデル" Or 商品名 = "アーデル2本セット" Or _
       商品名 = "アーデル＋シャンプー" Or _
       商品名 = "新ブスタ" Or _
       商品名 = "新ブスタ＋シャンプー" Or _
       商品名 = "ブースター" Or _
       商品名 = "ブースター（Ｗ発毛月間）" Or _
       商品名 = "ブースター＋シャンプー" Or _
       商品名 = "新ハイブリッター" Or _
       商品名 = "新ハイブリッター＋シャンプー" Or _
       商品名 = "ハイブリッド" Or _
       商品名 = "ハイブリッド＋シャンプー" Or _
       商品名 = "ナイスレディー" Or _
       商品名 = "ナイスレディー＋シャンプー" Or _
       商品名 = "ハイブリッド（プレゼント）" Or _
       商品名 = "シャンプー" Or _
       商品名 = "シャンプー2本セット" Or _
       商品名 = "シャンプー（プレゼント）" Or _
       商品名 = "シャンプー＋トリートメント" Or _
       商品名 = "トリートメント" Or _
       商品名 = "トリートメント（プレゼント）" Or _
       商品名 = "アーデル＆シャンプー試供品" Or _
       商品名 = "アーデル試供品" Or _
       商品名 = "シャンプー試供品" Then
       
       アーデル判定 = 1
       
    End If
    
    If 商品名 = "アーデル活用・マニュアル（プレゼント）" Or _
       商品名 = "毎日の積み重ねが大切です・マニュアル（プレゼント）" Or _
       商品名 = "ドクターアーデル・育毛ＤＶＤ（プレゼント）" Or _
       商品名 = "育毛と運動・マニュアル（プレゼント）" Or _
       商品名 = "育毛・発毛マニュアル（プレゼント）" Then
       
       アーデル判定 = 9
       
    End If
    
End Function

'************************************************************************
'機  能 :売上データのCSVを出力する
'************************************************************************
Private Sub 会計CSV出力()
    
    Dim intFileNo       As Integer
    Dim 売上データRS As New ADODB.Recordset
    
    intFileNo = FreeFile()
    
    Call 売上データ読込(売上データRS)
    
    If G_店舗名 = "トリニティー楽天市場店" Then
        Open "C:\顧客管理\楽天_売上.csv" For Output As #intFileNo
    Else
        Open "C:\顧客管理\Yahoo_売上.csv" For Output As #intFileNo
    End If
    
    With 売上データRS
        Do Until .EOF
    
'        Print #intFileNo, !識別フラグ & "," & !伝票NO & "," & !決算 & "," & !取引日時 & "," & !借方勘定科目 & "," & !借方補助科目 & "," & !借方部門 & "," & _
'                            !借方税区分 & "," & !借方金額 & "," & !借方税金額 & "," & !貸方勘定科目 & "," & !貸方補助科目 & "," & _
'                            !貸方部門 & "," & !貸方税区分 & "," & !貸方金額 & "," & !貸方税金額 & "," & !摘要 & "," & _
'                            !番号 & "," & !期日 & "," & !タイプ & "," & !生成元 & "," & !仕分メモ & "," & !付箋1 & "," & !付箋2 & "," & !調整
        
        Print #intFileNo, !識別フラグ1 & !識別フラグ2 & !識別フラグ3 & "," & !伝票NO & "," & !取引日時 & "," & _
                            !借方勘定科目 & "," & !借方補助科目 & "," & !借方部門 & "," & !借方税区分 & "," & !借方金額 & "," & _
                            !貸方勘定科目 & "," & !貸方補助科目 & "," & !貸方部門 & "," & !貸方税区分 & "," & !貸方金額 & "," & _
                            !摘要 & "," & !タイプ & "," & !生成元 & "," & "0" & "," & "0" & "," & !借方税金額 & "," & !貸方税金額 & "," & "no" & "," & "no" & "," & "no" & "," & """"""
        .MoveNext
        Loop
        
        .Close
    End With
    
    Close #intFileNo
    
End Sub


'************************************************************************
'機  能 :Yahoo注文ステータス変更
'************************************************************************
Private Sub 楽天_注文ステータス変更()
    
    Dim intFileNo1      As Integer
    Dim intFileNo2      As Integer
    Dim 売上データRS    As New ADODB.Recordset
    Dim 注文データRS    As New ADODB.Recordset
    Dim 注文番号        As String
    Dim 注文番号w       As String
    Dim 位置            As Integer
    Dim 配送日          As String
    Dim フラグ1         As Boolean
    Dim フラグ2         As Boolean
    Dim 注文ID          As String
    
    フラグ1 = False
    フラグ2 = False
    
    ' FileSystemObject (FSO) の新しいインスタンスを生成する
    Dim cFso As FileSystemObject
    Set cFso = New FileSystemObject

    ' ファイルを削除する
    On Error Resume Next
    Call cFso.DeleteFile("C:\顧客管理\rakuten_status_001.csv")
    On Error Resume Next
    Call cFso.DeleteFile("C:\顧客管理\rakuten_status_002.csv")

    ' 不要になった時点で参照を解放する (Terminate イベントを早めに起こす)
    Set cFso = Nothing
    
    Call 売上データ読込(売上データRS)
    
    intFileNo1 = FreeFile()
    Open "C:\顧客管理\rakuten_status_001.csv" For Output As #intFileNo1
    
    intFileNo2 = FreeFile()
    Open "C:\顧客管理\rakuten_status_002.csv" For Output As #intFileNo2
    
    Print #intFileNo1, """受注番号""" + "," + """受注ステータス""" + "," + """配送日""" + "," + """お荷物伝票番号"""
    
    Print #intFileNo2, """共同購入受注番号""" + "," + """受注ステータス""" + "," + """配送日""" + "," + """お荷物伝票番号"""
    
    注文番号w = ""
    
    With 売上データRS
        Do Until .EOF
            If !借方勘定科目 = "売掛金" Or !借方勘定科目 = "普通預金" Then
                
                位置 = InStr(!注文番号, "#")
                
                If 位置 > 0 Then
                    注文番号 = Left(!注文番号, 位置 - 1)
                Else
                    注文番号 = Trim(!注文番号)
                End If
                
                注文ID = !注文ID
                
                If 注文番号 <> 注文番号w Then
                    Call 注文番号検索(注文ID, 注文データRS)
                    
                    配送日 = """" + Mid(注文データRS!出荷日, 1, 4) + "-" + Mid(注文データRS!出荷日, 6, 2) + "-" + Mid(注文データRS!出荷日, 9, 2) + """"
                    
                    If Not 注文データRS.EOF Then
                        If 注文データRS!注文元 = "楽天" Then
                            
                            If InStr(注文番号, "-g") > 0 Then
                                ' 共同購入
                                Print #intFileNo2, """" + 注文番号 + """" + "," + """処理済""" + "," + 配送日 + "," + """" + Trim(注文データRS!問合番号) + """"
                                フラグ2 = True
                            Else
                                ' 通常購入
                                Print #intFileNo1, """" + 注文番号 + """" + "," + """処理済""" + "," + 配送日 + "," + """" + Trim(注文データRS!問合番号) + """"
                                フラグ1 = True
                            End If
                        End If
                    End If
                    
                    注文データRS.Close
                    注文番号w = 注文番号
                
                End If
            End If
            .MoveNext
        Loop
        
        .Close
    End With
    
    Close #intFileNo1
    Close #intFileNo2
    
    ' FileSystemObject (FSO) の新しいインスタンスを生成する
    Set cFso = New FileSystemObject
    
    ' ファイルを削除する
    If フラグ1 = False Then
        On Error Resume Next
        Call cFso.DeleteFile("C:\顧客管理\rakuten_status_001.csv")
    End If
    
    If フラグ2 = False Then
        On Error Resume Next
        Call cFso.DeleteFile("C:\顧客管理\rakuten_status_002.csv")
    End If
    
    Set cFso = Nothing
    
End Sub

'************************************************************************
'機  能 :Yahoo注文ステータス変更
'************************************************************************
Private Sub Yahoo_注文ステータス変更()
    
    Dim intFileNo       As Integer
    Dim 売上データRS    As New ADODB.Recordset
    Dim 注文データRS    As New ADODB.Recordset
    Dim 注文番号        As String
    Dim 注文番号w       As String
    Dim 位置            As Integer
    Dim 注文ID          As String
    
    ' FileSystemObject (FSO) の新しいインスタンスを生成する
    Dim cFso As FileSystemObject
    Set cFso = New FileSystemObject

    ' ファイルを削除する
    On Error Resume Next
    Call cFso.DeleteFile("C:\顧客管理\Yahoo_status.csv")

    ' 不要になった時点で参照を解放する (Terminate イベントを早めに起こす)
    Set cFso = Nothing

    intFileNo = FreeFile()
    
    Call 売上データ読込(売上データRS)
    
    Open "C:\顧客管理\Yahoo_status.csv" For Output As #intFileNo
    
    Print #intFileNo, """OrderID""" + "," + """Status""" + "," + """Quantity1""" + "," + """Shipping1""" + "," + """Paymentcharge1""" + "," + """Gift Wrap1""" + "," + """Discount1"""
    
    注文番号w = ""
    
    With 売上データRS
        Do Until .EOF
            If !借方勘定科目 = "売掛金" Or !借方勘定科目 = "普通預金" Then
                
                位置 = InStr(!注文番号, "#")
                
                If 位置 > 0 Then
                    注文番号 = Left(!注文番号, 位置 - 1)
                Else
                    注文番号 = Trim(!注文番号)
                End If
                  
                注文ID = !注文ID
                
                If 注文番号 <> 注文番号w Then
                    Call 注文番号検索(注文ID, 注文データRS)
                    
                    If Not 注文データRS.EOF Then
                        If 注文データRS!注文元 = "Yahoo" Then
                            Print #intFileNo, 注文番号 + "," + """完了""" + "," + "" + "," + "" + "," + "" + "," + "" + "," + ""
                        End If
                    End If
                    
                    注文データRS.Close
                    注文番号w = 注文番号
                
                End If
            End If
            .MoveNext
        Loop
        
        .Close
    End With
    
    Close #intFileNo
    
End Sub


'************************************************************************
'機  能 :オートシップ画面を表示する
'************************************************************************
Private Sub cmdオートシップ_Click()

    Dim ADF016      As New ADF016

    Call ADF016.Show(1)
    
End Sub

'************************************************************************
'機  能 :一括完了する
'************************************************************************
Private Sub cmd一括完了_Click()

    Dim ADF019      As New ADF019

    Call ADF019.Show(1)

    'Call cmd検索_Click
    
    Call cmd未出荷一覧_Click
    
End Sub

'************************************************************************
'機  能 :郵便番号から住所を表示する
'************************************************************************
Private Sub 郵便番号から住所を変換する()
    
    Dim 郵便番号        As String
    Dim 郵便番号辞書RS  As New ADODB.Recordset
    Dim ADF014          As New ADF014
    Dim 件数            As Integer
    Dim 住所_上段       As String
    Dim 住所_中段       As String
    
    If txt住所_上段 = "" Then
        
        郵便番号 = txt郵便番号.Text
        
        If Len(郵便番号) = 7 Then
            郵便番号 = Mid(郵便番号, 1, 3) & "-" & Mid(郵便番号, 4, 4)
            txt郵便番号.Text = 郵便番号
        End If
        
        If Len(郵便番号) = 8 Then
            件数 = 住所件数検索(郵便番号)
            
            If 件数 > 1 Then
                Call ADF014.SET_郵便番号(郵便番号)
                Call ADF014.Show(1)
                Call ADF014.GET_住所(住所_上段, 住所_中段)
                txt住所_上段.Text = 住所_上段
                txt住所_中段.Text = 住所_中段
            Else
                Call 住所検索(郵便番号, 郵便番号辞書RS)
                If Not 郵便番号辞書RS.EOF Then
                    txt住所_上段.Text = 郵便番号辞書RS!都道府県名 + 郵便番号辞書RS!市区町村名
                    txt住所_中段.Text = 郵便番号辞書RS!町域名
                End If
            
                郵便番号辞書RS.Close
            End If
        End If
    End If
    
End Sub

'************************************************************************
'機  能 :顧客情報クリアする。
'************************************************************************
Private Sub 顧客情報クリア()
        
    txt顧客ID.Text = ""
    txt顧客名.Text = ""
    txtフリガナ.Text = ""
    txt郵便番号.Text = ""
    txt住所_上段.Text = ""
    txt住所_中段.Text = ""
    txt住所_下段.Text = ""
    txt電話番号.Text = ""
    txtメール.Text = ""
    txt楽天メール.Text = ""
    cmbアーデルクラブ.ListIndex = 0
    txt入会日.Text = "____/__/__"
    txt退会日.Text = "____/__/__"
    opt男性.Value = True
    opt女性.Value = False
    txt備考.Text = ""
    chkメール送信 = 1
    txt誕生日.Text = "____/__/__"

End Sub

'************************************************************************
'機  能 :注文情報クリアする。
'************************************************************************
Private Sub 注文情報クリア()
        
    txt受注日.Text = "____/__/__"
    txt注文ID.Text = ""
    txt注文番号.Text = ""
    cmbステータス.ListIndex = 0
    cmb商品名.ListIndex = 0
    cmb注文方法.ListIndex = 0
    txt配達日時.Text = ""
    txt出荷日.Text = "____/__/__"
    cmb宅配業者.ListIndex = 0
    txt支払番号.Text = ""
    txt問合番号.Text = ""
    txt単価.Value = 0
    txt割引.Value = 0
    txt数量.Value = 0
    txt送料.Value = 0
    txt返金.Value = 0
    txtその他手数料.Value = 0
    txt合計金額.Text = 0
    txtメール送信.Text = ""
    cmb注文元.Text = ""
    txt備考2.Text = ""
    txtコモライフ.Text = ""
    txt出荷予定日.Text = "____/__/__"
    txt決済URL.Text = ""

End Sub

'************************************************************************
'機  能 :顧客検索
'************************************************************************
Private Sub cmd検索_Click()

    Dim 検索値 As String
    Dim 顧客マスタRS As New ADODB.Recordset

    If cmb検索条件.Text = "" Then Exit Sub
    
    MousePointer = vbHourglass
    
    If cmb検索条件.Text = "出荷日" Then
        ' 入力された出荷日を元に、顧客検索を行なう
        Call 顧客検索2(txt検索条件.Text, 顧客マスタRS)
    ElseIf cmb検索条件.Text = "決済ID" Then
    
        ' 入力された決済ID検索を行う
        Call 顧客検索3(txt検索条件.Text, 顧客マスタRS)
    ElseIf cmb検索条件.Text = "コモライフNO" Then
    
        ' 入力されたコモライフNO検索を行う
        Call 顧客検索4(txt検索条件.Text, 顧客マスタRS)
    ElseIf cmb検索条件.Text = "注文番号" Then
    
        ' 入力された注文番号検索を行う
        Call 顧客検索5(txt検索条件.Text, 顧客マスタRS)
        
    ElseIf cmb検索条件.Text = "〒" Then
    
        ' 入力された〒検索を行う
        検索値 = txt検索条件.Text
        Call 顧客検索(検索値, 顧客マスタRS, "[〒]")
        
    ElseIf cmb検索条件.Text = "問合番号" Then
        If InStr(txt検索条件.Text, "-") <= 0 Then
            検索値 = 問合番号編集(txt検索条件.Text)
        Else
            検索値 = txt検索条件.Text
        End If
            
        ' 入力された顧客名を元にワイルドカード検索を行う
        Call 顧客検索6(検索値, 顧客マスタRS)
    Else
        検索値 = "%" & txt検索条件.Text & "%"
        
        ' 入力された顧客名を元にワイルドカード検索を行う
        Call 顧客検索(検索値, 顧客マスタRS, cmb検索条件.Text)
    End If
    
    ' 顧客リストを表示する
    Call 顧客リスト表示(顧客マスタRS)
    
    MousePointer = vbNormal

End Sub

'************************************************************************
'機  能 :問合番号を編集する
'************************************************************************
Private Function 問合番号編集(ByVal 問合番号 As String) As String
    
    Dim 問合番号1 As String
    Dim 問合番号2 As String
    Dim 問合番号3 As String
    
    問合番号1 = Mid(問合番号, 1, 4)
    問合番号2 = Mid(問合番号, 5, 4)
    問合番号3 = Mid(問合番号, 9, 4)
    
    問合番号編集 = 問合番号1 & "-" & 問合番号2 & "-" & 問合番号3
    
End Function

'************************************************************************
'機  能 :未出荷を検索する
'************************************************************************
Private Sub cmd未出荷一覧_Click()

    Dim 顧客マスタRS As New ADODB.Recordset
    
    MousePointer = vbHourglass
        
    Call トランザクションデータの更新
    
    ' 未出荷一覧を取得する
    Call 未出荷検索(顧客マスタRS)

    ' 顧客リストを表示する
    Call 顧客リスト表示(顧客マスタRS)
        
    MousePointer = vbNormal

End Sub

'************************************************************************
'機  能 :未入金を検索する
'************************************************************************
Private Sub cmd未入金_Click()

    Dim 顧客マスタRS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call トランザクションデータの更新
    
    ' 未出荷一覧を取得する
    Call 未入金検索(顧客マスタRS)

    ' 顧客リストを表示する
    Call 顧客リスト表示(顧客マスタRS)
    
    MousePointer = vbNormal

End Sub

'************************************************************************
'機  能 :出荷予定一覧を出力する
'************************************************************************
Private Sub cmd出荷予定一覧_Click()

    Dim 出荷予定一覧RS As New ADODB.Recordset
    Dim ADF012 As New ADF012
    
    Call トランザクションデータの更新
    
    MousePointer = vbHourglass
    
    ' 未出荷一覧を取得する
    Call 出荷予定一覧(出荷予定一覧RS)
    
    MousePointer = vbNormal
    
   
    ' 確認メッセージを表示する
    'If MsgBox("納品書を印刷してよろしいですか？", vbYesNo, "顧客管理") <> vbYes Then Exit Sub
    
    If 出荷予定一覧RS.EOF Then
        Call MsgBox("出荷予定がありません", vbOKOnly, "顧客管理")
        Exit Sub
    End If
    
    Set G_出荷予定リスト = Nothing
    Set G_出荷予定リスト = New 出荷予定リスト
    Call G_出荷予定リスト.Database.SetDataSource(出荷予定一覧RS)
    Call ADF012.初期設定("出荷予定リスト")
    Call ADF012.Show(vbModal)
    出荷予定一覧RS.Close

End Sub

'************************************************************************
'機  能 :アーデル購入者
'************************************************************************
Private Sub cmdアーデル購入者_Click()

    Dim 顧客マスタRS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call トランザクションデータの更新
    
    ' アーデル購入者を取得する
    Call アーデル購入者検索(顧客マスタRS)

    ' 顧客リストを表示する
    Call 顧客リスト表示(顧客マスタRS)
    
    MousePointer = vbNormal

End Sub

'************************************************************************
'機  能 :アーデルクラブ会員
'************************************************************************
Private Sub cmdクラブ検索_Click()

    Dim 顧客マスタRS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call トランザクションデータの更新
    
    ' アーデルクラブ会員を取得する
    Call アーデルクラブ会員検索(顧客マスタRS)

    ' 顧客リストを表示する
    Call 顧客リスト表示(顧客マスタRS)
    
    MousePointer = vbNormal

End Sub

'************************************************************************
'機  能 :アーデルクラブ未加入検索
'************************************************************************
Private Sub cmdクラブ未加入_Click()

    Dim i As Integer
    Dim 顧客マスタRS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call トランザクションデータの更新
    
    ' アーデルクラブメール未送信顧客を取得する
    Call メール検索3(顧客マスタRS)

    ' 顧客リストを表示する
    Call 顧客リスト表示(顧客マスタRS)

    MousePointer = vbNormal

End Sub


'************************************************************************
'機  能 :アーデルモモ検索
'************************************************************************
Private Sub cmdモモ検索_Click()

    Dim 顧客マスタRS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call トランザクションデータの更新
    
    ' アーデルモモ検索を行う
    Call モモ検索(顧客マスタRS)

    ' 顧客リストを表示する
    Call 顧客リスト表示(顧客マスタRS)

    MousePointer = vbNormal

End Sub

'************************************************************************
'機  能 :コモライフの仕入金額、荷造運賃の計算を行う
'************************************************************************
Private Sub cmd計算_Click()
    
    Dim 仕入金額    As Long
    Dim 荷造運賃    As Long
    
    Dim ADF017      As New ADF017

    Call ADF017.Show(1)

    Call ADF017.仕入金額_荷造運賃取得(仕入金額, 荷造運賃)
    
    txt仕入金額.Value = 仕入金額
    txt荷造運賃.Value = 荷造運賃
    
    Call 注文_更新
    
End Sub

'************************************************************************
'機  能 :新規注文の検索を行う
'************************************************************************
Private Sub cmd新規注文_Click()

    Dim 顧客マスタRS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call トランザクションデータの更新

    ' 新規注文の検索を行う
    Call ステータス検索("新規注文", 顧客マスタRS)

    ' 顧客リストを表示する
    Call 顧客リスト表示(顧客マスタRS)

    MousePointer = vbNormal

End Sub

'************************************************************************
'機  能 :入金待ちの検索を行う
'************************************************************************
Private Sub cmd入金待ち_Click()

    Dim 顧客マスタRS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call トランザクションデータの更新
    
    ' 入金待ちの検索を行う
    Call ステータス検索("入金待ち", 顧客マスタRS)

    ' 顧客リストを表示する
    Call 顧客リスト表示(顧客マスタRS)

    MousePointer = vbNormal

End Sub

'************************************************************************
'機  能 :出荷処理中の検索を行う
'************************************************************************
Private Sub cmd出荷処理中_Click()

    Dim 顧客マスタRS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call トランザクションデータの更新
    
    ' 出荷処理中の検索を行う
    Call ステータス検索("出荷処理", 顧客マスタRS)

    ' 顧客リストを表示する
    Call 顧客リスト表示(顧客マスタRS)

    MousePointer = vbNormal

End Sub

'************************************************************************
'機  能 :出荷済みの検索を行う
'************************************************************************
Private Sub cmd出荷済み_Click()

    Dim 顧客マスタRS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call トランザクションデータの更新
    
    ' 出荷済みの検索を行う
    Call ステータス検索("出荷完了", 顧客マスタRS)

    ' 顧客リストを表示する
    Call 顧客リスト表示(顧客マスタRS)

    MousePointer = vbNormal

End Sub

'************************************************************************
'機  能 :コモライフ検索を行う
'************************************************************************
Private Sub cmdコモライフ_Click()

    Dim 顧客マスタRS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call トランザクションデータの更新
    
    ' 保留中の検索を行う
    Call ステータス検索("コモライフ", 顧客マスタRS)

    ' 顧客リストを表示する
    Call 顧客リスト表示(顧客マスタRS)

    MousePointer = vbNormal


End Sub

'************************************************************************
'機  能 :保留中の検索を行う
'************************************************************************
Private Sub cmd保留中検索_Click()

    Dim 顧客マスタRS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call トランザクションデータの更新
    
    ' 保留中の検索を行う
    Call ステータス検索("保留", 顧客マスタRS)

    ' 顧客リストを表示する
    Call 顧客リスト表示(顧客マスタRS)

    MousePointer = vbNormal

End Sub

'************************************************************************
'機  能 :キャンセルの検索を行う
'************************************************************************
Private Sub cmdキャンセル検索_Click()

    Dim 顧客マスタRS As New ADODB.Recordset
    
    MousePointer = vbHourglass
    
    Call トランザクションデータの更新
    
    ' キャンセルの検索を行う
    Call ステータス検索("キャンセル", 顧客マスタRS)

    ' 顧客リストを表示する
    Call 顧客リスト表示(顧客マスタRS)

    MousePointer = vbNormal

End Sub

'************************************************************************
'機  能 :全顧客選択
'************************************************************************
Private Sub cmd全選択_Click()

    Dim i As Integer
    
    MousePointer = vbHourglass
    
    Call トランザクションデータの更新
    
    For i = 1 To va顧客リスト.MaxRows
        Call SpreadSetVal(va顧客リスト, i, COL_チェック, "1")
    Next i
    
    MousePointer = vbNormal

End Sub

'************************************************************************
'機  能 :全顧客選択解除
'************************************************************************
Private Sub cmd全解除_Click()

    Dim i As Integer
    
    MousePointer = vbHourglass
    
    Call トランザクションデータの更新
    
    For i = 1 To va顧客リスト.MaxRows
        Call SpreadSetVal(va顧客リスト, i, COL_チェック, "0")
    Next i
    
    MousePointer = vbNormal

End Sub

'************************************************************************
'機  能 :納品書を印刷する。
'************************************************************************
Private Sub cmd納品書_Click()
    
    Dim i As Integer
    Dim 注文元 As String
        
    Call トランザクションデータの更新

    If チェック件数取得() <= 0 Then
        Call MsgBox("納品書を印刷する明細にチェックを付けて下さい", vbOKOnly, "顧客管理")
        Exit Sub
    End If
    
    For i = 1 To va注文リスト.MaxRows
        If SpreadGetVal(va注文リスト, i, COL_チェック) = "1" Then
            注文元 = SpreadGetVal(va注文リスト, i, COL_注文元)
            
            If 注文元 <> "野口さん" Then
                Call cmd納品書_sub1
                Exit For
            Else
                Call cmd納品書_sub2
                Exit For
            End If
        End If
    Next i
    
    'Call Sleep(3000)

End Sub

'************************************************************************
'機  能 :納品書を印刷する（キャットハンド用）
'************************************************************************
Private Sub cmd納品書_sub1()

    Dim i As Integer
    Dim 顧客ID As String
    Dim 注文ID As String
    Dim ADF012 As New ADF012
    Dim 納品書RS As New ADODB.Recordset
   
    ' 確認メッセージを表示する
    'If MsgBox("納品書を印刷してよろしいですか？", vbYesNo, "顧客管理") <> vbYes Then Exit Sub
    
    顧客ID = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_顧客ID)
    注文ID = ""
    
    For i = 1 To va注文リスト.MaxRows
        If SpreadGetVal(va注文リスト, i, COL_チェック) = "1" Then
            注文ID = 注文ID & SpreadGetVal2(va注文リスト, i, COL_注文ID) & ","
        End If
    Next i
        
    If 注文ID <> "" Then
        注文ID = Left(注文ID, Len(注文ID) - 1)
        Call 納品データ取得(顧客ID, 注文ID, 納品書RS)
        If Not 納品書RS.EOF Then
            Set G_納品書 = Nothing
            Set G_納品書 = New 納品書
            Call G_納品書.Database.SetDataSource(納品書RS)
            Call ADF012.初期設定("納品書")
            Call ADF012.Show(vbModal)
        End If
        納品書RS.Close
    End If

End Sub
'************************************************************************
'機  能 :納品書を印刷する（アーデルモ用）
'************************************************************************
Private Sub cmd納品書_sub2()

    Dim i As Integer
    Dim 顧客ID As String
    Dim 注文ID As String
    Dim ADF012 As New ADF012
    Dim 納品書RS As New ADODB.Recordset
   
    ' 確認メッセージを表示する
    'If MsgBox("納品書を印刷してよろしいですか？", vbYesNo, "顧客管理") <> vbYes Then Exit Sub
    
    顧客ID = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_顧客ID)
    注文ID = ""
    
    For i = 1 To va注文リスト.MaxRows
        If SpreadGetVal(va注文リスト, i, COL_チェック) = "1" Then
            注文ID = 注文ID & SpreadGetVal2(va注文リスト, i, COL_注文ID) & ","
        End If
    Next i
        
    If 注文ID <> "" Then
        注文ID = Left(注文ID, Len(注文ID) - 1)
        Call 納品データ取得(顧客ID, 注文ID, 納品書RS)
        If Not 納品書RS.EOF Then
            Set G_納品書2 = Nothing
            Set G_納品書2 = New 納品書2
            Call G_納品書2.Database.SetDataSource(納品書RS)
            Call ADF012.初期設定("納品書2")
            Call ADF012.Show(vbModal)
        End If
        納品書RS.Close
    End If
    
End Sub

'************************************************************************
'機  能 お礼状を印刷する。
'************************************************************************
Private Sub cmd礼状_Click()
    Dim i As Integer
    Dim 注文元 As String
    
    Call トランザクションデータの更新

    If チェック件数取得() <= 0 Then
        Call MsgBox("お礼状を印刷する明細にチェックを付けて下さい", vbOKOnly, "顧客管理")
        Exit Sub
    End If
    
    For i = 1 To va注文リスト.MaxRows
        If SpreadGetVal(va注文リスト, i, COL_チェック) = "1" Then
            注文元 = SpreadGetVal(va注文リスト, i, COL_注文元)
            
            If 注文元 <> "野口さん" Then
                Call cmd礼状_sub1(i)
            Else
                Call cmd礼状_sub2(i)
            End If
        End If
    Next i

End Sub

'************************************************************************
'機  能 お礼状を印刷する（キャットハンド用）
'************************************************************************
Private Sub cmd礼状_sub1(ByVal i As Integer)

    Dim 顧客名 As String
    Dim 商品名 As String
    Dim ADF012 As New ADF012
    Dim お礼状RS As New ADODB.Recordset
    Dim 注意事項RS As New ADODB.Recordset
    Dim ミニまぐRS As New ADODB.Recordset
    
    If SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_お届け先名) <> "" Then
        顧客名 = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_お届け先名)
    Else
        顧客名 = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_顧客名)
    End If
    
    商品名 = SpreadGetVal(va注文リスト, i, COL_商品名)
        
    ' 確認メッセージを表示する
    'If MsgBox("お礼状を印刷してよろしいですか？", vbYesNo, "顧客管理") <> vbYes Then Exit Sub
    
    If Left(商品名, 4) = "アーデル" Then
        
        Call お礼状データ取得(顧客名, 商品名, お礼状RS)
        
        If Not お礼状RS.EOF Then
            Set G_お礼状 = Nothing
            Set G_お礼状 = New お礼状
            Call G_お礼状.Database.SetDataSource(お礼状RS)
            Call ADF012.初期設定("お礼状")
            Call ADF012.Show(vbModal)
        End If
        
        お礼状RS.Close
        
        If 商品名 <> "アーデル資料" Then
            
            Set G_注意事項 = Nothing
            Set G_注意事項 = New 注意事項
            Call ADF012.初期設定("注意事項")
            Call ADF012.Show(vbModal)
            
        End If
        
        If InStr(1, 商品名, "シャンプー") > 0 Then
    
            Call お礼状データ取得(顧客名, 商品名, お礼状RS)
            
            If Not お礼状RS.EOF Then
                Set G_お礼状3 = Nothing
                Set G_お礼状3 = New お礼状3
                Call G_お礼状3.Database.SetDataSource(お礼状RS)
                Call ADF012.初期設定("お礼状3")
                Call ADF012.Show(vbModal)
            End If
            
            お礼状RS.Close
        End If
    
    ElseIf Left(商品名, 13) = "シャンプー＋トリートメント" Then
    
        Call お礼状データ取得(顧客名, 商品名, お礼状RS)
        
        If Not お礼状RS.EOF Then
            Set G_お礼状3 = Nothing
            Set G_お礼状3 = New お礼状3
            Call G_お礼状3.Database.SetDataSource(お礼状RS)
            Call ADF012.初期設定("お礼状3")
            Call ADF012.Show(vbModal)
        End If
        
        お礼状RS.Close
        
        Call お礼状データ取得(顧客名, 商品名, お礼状RS)
        
        If Not お礼状RS.EOF Then
            Set G_お礼状7 = Nothing
            Set G_お礼状7 = New お礼状7
            Call G_お礼状7.Database.SetDataSource(お礼状RS)
            Call ADF012.初期設定("お礼状7")
            Call ADF012.Show(vbModal)
        End If
        
        お礼状RS.Close
        
    ElseIf Left(商品名, 5) = "シャンプー" Then
    
        Call お礼状データ取得(顧客名, 商品名, お礼状RS)
        
        If Not お礼状RS.EOF Then
            Set G_お礼状3 = Nothing
            Set G_お礼状3 = New お礼状3
            Call G_お礼状3.Database.SetDataSource(お礼状RS)
            Call ADF012.初期設定("お礼状3")
            Call ADF012.Show(vbModal)
        End If
        
        お礼状RS.Close
        
    ElseIf Left(商品名, 7) = "トリートメント" Then
    
        Call お礼状データ取得(顧客名, 商品名, お礼状RS)
        
        If Not お礼状RS.EOF Then
            Set G_お礼状7 = Nothing
            Set G_お礼状7 = New お礼状7
            Call G_お礼状7.Database.SetDataSource(お礼状RS)
            Call ADF012.初期設定("お礼状7")
            Call ADF012.Show(vbModal)
        End If
        
        お礼状RS.Close
        
    ElseIf Left(商品名, 5) = "ブースター" Then
    
        Call お礼状データ取得(顧客名, 商品名, お礼状RS)
        
        If 商品名 = "ブースター（Ｗ発毛月間）" Then
            If Not お礼状RS.EOF Then
                Set G_お礼状11 = Nothing
                Set G_お礼状11 = New お礼状11
                Call G_お礼状11.Database.SetDataSource(お礼状RS)
                Call ADF012.初期設定("お礼状11")
                Call ADF012.Show(vbModal)
            End If
        Else
            If Not お礼状RS.EOF Then
                Set G_お礼状4 = Nothing
                Set G_お礼状4 = New お礼状4
                Call G_お礼状4.Database.SetDataSource(お礼状RS)
                Call ADF012.初期設定("お礼状4")
                Call ADF012.Show(vbModal)
            End If
        End If
        
        お礼状RS.Close
    
        
        If InStr(1, 商品名, "シャンプー") > 0 Then
    
            Call お礼状データ取得(顧客名, 商品名, お礼状RS)
            
            If Not お礼状RS.EOF Then
                Set G_お礼状3 = Nothing
                Set G_お礼状3 = New お礼状3
                Call G_お礼状3.Database.SetDataSource(お礼状RS)
                Call ADF012.初期設定("お礼状3")
                Call ADF012.Show(vbModal)
            End If
            
            お礼状RS.Close
        End If
    
    ElseIf Left(商品名, 6) = "ハイブリッド" Then
    
        Call お礼状データ取得(顧客名, 商品名, お礼状RS)
        
        If Not お礼状RS.EOF Then
            Set G_お礼状5 = Nothing
            Set G_お礼状5 = New お礼状5
            Call G_お礼状5.Database.SetDataSource(お礼状RS)
            Call ADF012.初期設定("お礼状5")
            Call ADF012.Show(vbModal)
        End If
        
        お礼状RS.Close
    
        
        If InStr(1, 商品名, "シャンプー") > 0 Then
    
            Call お礼状データ取得(顧客名, 商品名, お礼状RS)
            
            If Not お礼状RS.EOF Then
                Set G_お礼状3 = Nothing
                Set G_お礼状3 = New お礼状3
                Call G_お礼状3.Database.SetDataSource(お礼状RS)
                Call ADF012.初期設定("お礼状3")
                Call ADF012.Show(vbModal)
            End If
            
            お礼状RS.Close
        End If
    
    ElseIf Left(商品名, 7) = "ナイスレディー" Then
    
        Call お礼状データ取得(顧客名, 商品名, お礼状RS)
        
        If Not お礼状RS.EOF Then
            Set G_お礼状6 = Nothing
            Set G_お礼状6 = New お礼状6
            Call G_お礼状6.Database.SetDataSource(お礼状RS)
            Call ADF012.初期設定("お礼状6")
            Call ADF012.Show(vbModal)
        End If
        
        お礼状RS.Close
    
        
        If InStr(1, 商品名, "シャンプー") > 0 Then
    
            Call お礼状データ取得(顧客名, 商品名, お礼状RS)
            
            If Not お礼状RS.EOF Then
                Set G_お礼状3 = Nothing
                Set G_お礼状3 = New お礼状3
                Call G_お礼状3.Database.SetDataSource(お礼状RS)
                Call ADF012.初期設定("お礼状3")
                Call ADF012.Show(vbModal)
            End If
            
            お礼状RS.Close
        End If
    
    ElseIf Left(商品名, 4) = "新ブスタ" Then
    
        Call お礼状データ取得(顧客名, 商品名, お礼状RS)
        
        If Not お礼状RS.EOF Then
            Set G_お礼状8 = Nothing
            Set G_お礼状8 = New お礼状8
            Call G_お礼状8.Database.SetDataSource(お礼状RS)
            Call ADF012.初期設定("お礼状8")
            Call ADF012.Show(vbModal)
        End If
        
        お礼状RS.Close
    
        
        If InStr(1, 商品名, "シャンプー") > 0 Then
    
            Call お礼状データ取得(顧客名, 商品名, お礼状RS)
            
            If Not お礼状RS.EOF Then
                Set G_お礼状3 = Nothing
                Set G_お礼状3 = New お礼状3
                Call G_お礼状3.Database.SetDataSource(お礼状RS)
                Call ADF012.初期設定("お礼状3")
                Call ADF012.Show(vbModal)
            End If
            
            お礼状RS.Close
        End If
    
    ElseIf Left(商品名, 8) = "新ハイブリッター" Then
    
        Call お礼状データ取得(顧客名, 商品名, お礼状RS)
        
        If Not お礼状RS.EOF Then
            Set G_お礼状10 = Nothing
            Set G_お礼状10 = New お礼状10
            Call G_お礼状10.Database.SetDataSource(お礼状RS)
            Call ADF012.初期設定("お礼状10")
            Call ADF012.Show(vbModal)
        End If
        
        お礼状RS.Close
    
        
        If InStr(1, 商品名, "シャンプー") > 0 Then
    
            Call お礼状データ取得(顧客名, 商品名, お礼状RS)
            
            If Not お礼状RS.EOF Then
                Set G_お礼状3 = Nothing
                Set G_お礼状3 = New お礼状3
                Call G_お礼状3.Database.SetDataSource(お礼状RS)
                Call ADF012.初期設定("お礼状3")
                Call ADF012.Show(vbModal)
            End If
            
            お礼状RS.Close
        End If
    
    
    ElseIf 商品名 = "ミニまぐ" Then
    
        Call ミニまぐデータ取得(顧客名, ミニまぐRS)
        
        If Not ミニまぐRS.EOF Then
            Set G_ミニまぐ = Nothing
            Set G_ミニまぐ = New ミニまぐ
            Call G_ミニまぐ.Database.SetDataSource(ミニまぐRS)
            Call ADF012.初期設定("ミニまぐ")
            Call ADF012.Show(vbModal)
        End If
        
        ミニまぐRS.Close
    Else
        Call MsgBox("指定した商品のお礼状はサポートされていません。", vbOK, "顧客管理")

    End If
    
End Sub

'************************************************************************
'機  能 お礼状を印刷する（アーデルモモ用）
'************************************************************************
Private Sub cmd礼状_sub2(ByVal i As Integer)

    Dim 顧客名 As String
    Dim 商品名 As String
    Dim ADF012 As New ADF012
    Dim お礼状RS As New ADODB.Recordset
    Dim 注意事項RS As New ADODB.Recordset

    If SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_お届け先名) <> "" Then
        顧客名 = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_お届け先名)
    Else
        顧客名 = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_顧客名)
    End If
    
    商品名 = SpreadGetVal(va注文リスト, i, COL_商品名)
        
    ' 確認メッセージを表示する
    'If MsgBox("お礼状を印刷してよろしいですか？", vbYesNo, "顧客管理") <> vbYes Then Exit Sub
    
    If 商品名 = "アーデル" Or 商品名 = "アーデル2本セット" Or 商品名 = "アーデル試供品" Or 商品名 = "アーデル(セール)" Or 商品名 = "アーデル＆シャンプー試供品" Or 商品名 = "アーデル資料" Then
        
        Call お礼状データ取得(顧客名, 商品名, お礼状RS)
        
        If Not お礼状RS.EOF Then
            Set G_お礼状2 = Nothing
            Set G_お礼状2 = New お礼状2
            Call G_お礼状2.Database.SetDataSource(お礼状RS)
            Call ADF012.初期設定("お礼状2")
            Call ADF012.Show(vbModal)
        End If
        
        お礼状RS.Close
        
        If 商品名 <> "アーデル資料" Then
            
            Set G_注意事項 = Nothing
            Set G_注意事項 = New 注意事項
            Call ADF012.初期設定("注意事項")
            Call ADF012.Show(vbModal)
            
        End If
    ElseIf 商品名 = "アーデルシャンプー" Or 商品名 = "アーデルシャンプー2本セット" Or 商品名 = "アーデルシャンプー試供品" Or 商品名 = "アーデル＆シャンプー試供品" Then
    
        Call お礼状データ取得(顧客名, 商品名, お礼状RS)
        
        If Not お礼状RS.EOF Then
            Set G_お礼状3 = Nothing
            Set G_お礼状3 = New お礼状3
            Call G_お礼状3.Database.SetDataSource(お礼状RS)
            Call ADF012.初期設定("お礼状3")
            Call ADF012.Show(vbModal)
        End If
        
        お礼状RS.Close
    
    End If
    
End Sub

'************************************************************************
'機  能 メールを行う。
'************************************************************************
Private Sub cmdメール_Click()
    
    Dim 顧客名      As String
    Dim メールID    As String
    Dim ADF015      As New ADF015
    Dim i           As Integer
        
    Call トランザクションデータの更新

    If チェック件数取得() < 1 Then
        Call MsgBox("メールする明細に１件以上チェックを付けて下さい", vbOKOnly, "顧客管理")
        Exit Sub
    End If
    
    For i = 1 To va注文リスト.MaxRows
        If SpreadGetVal(va注文リスト, i, COL_チェック) = "1" Then
        
            G_注文ROW = i
            
            ' 注文が未登録の場合エラーメッセージを表示する
            If SpreadGetVal(va注文リスト, G_注文ROW, COL_注文ID) = "-1" Then
                Call MsgBox("先ず注文を登録して下さい。", vbOKOnly, "顧客管理")
                Exit Sub
            End If
            
            If SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_お届け先名) <> "" Then
                顧客名 = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_お届け先名)
            Else
                顧客名 = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_顧客名)
            End If
            
            If SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_楽天メール) <> "" Then
                メールID = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_楽天メール)
            Else
                メールID = SpreadGetVal(va顧客リスト, G_顧客リスト_ROW, COL_メール)
            End If
        
            If 顧客名 = "" Then
                Call MsgBox("顧客名を入力して下さい。", vbOKOnly, "顧客管理")
                Exit Sub
            End If
        
            If メールID = "" Then
                Call MsgBox("メールIDを入力して下さい。", vbOKOnly, "顧客管理")
                Exit Sub
            End If
            
        End If
    Next i
            
    Call ADF015.Show(1)
    
End Sub

'************************************************************************
'機  能 個別メールを行う。
'************************************************************************
Private Sub cmd個別メール_Click()
    
    Dim ADF020      As New ADF020

    Call ADF020.Show(1)

End Sub

'************************************************************************
'機  能 アーデルクラブメールを行う。
'************************************************************************
Private Sub cmdアーデルクラブ_Click()
    
    Dim cnt             As Integer
    Dim 顧客ID          As String
    Dim 顧客名          As String
    Dim メールID        As String
    Dim row             As Integer
    Dim メール本文RS    As New ADODB.Recordset
    Dim メール内容      As String
    Dim サーバ          As String
    Dim 宛先            As String
    Dim 送信元          As String
    Dim 件名            As String
    Dim ret             As String
    
    If MsgBox("アーデルクラブメールを送信してもよろしいですか？", vbYesNo, "顧客管理") = vbNo Then
        Exit Sub
    End If

    MousePointer = vbHourglass
        
    Call トランザクションデータの更新

    cnt = 0
    For row = 1 To va顧客リスト.MaxRows
    
        If SpreadGetVal(va顧客リスト, row, COL_チェック) = "1" Then
            
            顧客名 = SpreadGetVal(va顧客リスト, row, COL_顧客名)
            If SpreadGetVal(va顧客リスト, row, COL_楽天メール) <> "" Then
                メールID = SpreadGetVal(va顧客リスト, row, COL_楽天メール)
            Else
                メールID = SpreadGetVal(va顧客リスト, row, COL_メール)
            End If
        
            If 顧客名 = "" Then
                MousePointer = vbNormal
                Call MsgBox("顧客名を入力して下さい。", vbOKOnly, "顧客管理")
                Exit Sub
            End If
        
            'If メールID = "" Then
            '    MousePointer = vbNormal
            '    Call MsgBox("メールIDを入力して下さい。", vbOKOnly, "顧客管理")
            '    Exit Sub
            'End If
            
            cnt = cnt + 1
        End If
    Next
        
    If cnt <= 0 Then
        MousePointer = vbNormal
        Call MsgBox("メールをする顧客に１件以上チェックを付けて下さい", vbOKOnly, "顧客管理")
        Exit Sub
    End If
    
    Call メール本文検索(8, メール本文RS)
    
    For row = 1 To va顧客リスト.MaxRows
    
        If SpreadGetVal(va顧客リスト, row, COL_チェック) = "1" Then
            
            顧客ID = SpreadGetVal(va顧客リスト, row, COL_顧客ID)
            顧客名 = SpreadGetVal(va顧客リスト, row, COL_顧客名)
            If SpreadGetVal(va顧客リスト, row, COL_楽天メール) <> "" Then
                メールID = SpreadGetVal(va顧客リスト, row, COL_楽天メール)
            Else
                メールID = SpreadGetVal(va顧客リスト, row, COL_メール)
            End If
            
            If メールID <> "" Then
                宛先 = メールID ' + Chr(9) + "info@cathand.jp"    ' 宛先
                件名 = メール本文RS!件名                        ' 件名
                メール内容 = ""
                メール内容 = メール内容 + 顧客名 + "様" + Chr$(13) + Chr$(10)
                メール内容 = メール内容 + Chr$(13) + Chr$(10)
                メール内容 = メール内容 + メール本文RS!文章1 + Chr$(13) + Chr$(10)
                'メール内容 = メール内容 + "※メールが不要な場合、「お名前」を明記の上、「メール不要」として返信下さい。" + Chr$(13) + Chr$(10)
            
                ' メール送信
                ret = SendMail(G_サーバ, 宛先, G_送信元, 件名, メール内容, "")
                                
                If Len(ret) <> 0 Then
                   'Call MsgBox("メール送信エラー：" & ret, vbOKOnly, "顧客管理")
                End If
                
                Sleep (1000 * 3)
                
                'If メール内容 <> "" Then
                '    Shell "..\bin\sendmail " + "|" + メール内容 + "|"
                '    Call メール送信者登録3(顧客ID)
                'End If
            End If
        End If
    Next
    
    Call MsgBox("メールを送信しました", vbOKOnly, "顧客管理")
    
    If メール本文RS.State <> adStateClosed Then
        メール本文RS.Close
    End If
    
    MousePointer = vbNormal
    
End Sub


'************************************************************************
'機  能 メルマガを発行する。
'************************************************************************
Private Sub cmdメルマガ発行_Click()
    
    Dim 顧客名          As String
    Dim メール内容      As String
    Dim 宛先            As String
    Dim 件名            As String
    Dim メルマガ送信予定日  As String
    Dim ret             As String
    Dim メルマガRS      As New ADODB.Recordset
    Dim 顧客マスタRS    As New ADODB.Recordset
    
    If MsgBox("メルマガを送信してもよろしいですか？", vbYesNo, "顧客管理") = vbNo Then
        Exit Sub
    End If
    
    MousePointer = vbHourglass
    
    Call トランザクションデータの更新
    
    ' 顧客マスタを全件リードする
    Call 顧客マスタ読込2(顧客マスタRS)
    
    With 顧客マスタRS
        Do Until .EOF
            If (!メール <> "" Or !楽天メール <> "") And !メルマガNO >= 0 Then
                
#If 0 Then
                 If Format(!メルマガ送信予定日, "yyyy/mm/dd") <= Format(Now, "yyyy/mm/dd") Or IsNull(!メルマガ送信予定日) Then
                'If Format(!メルマガ送信予定日, "yyyy/mm/dd") <= Format(Now, "yyyy/mm/dd") Then
                    
                    Call メルマガ本文検索(IIf(!メルマガNO <= 0, 1, !メルマガNO), メルマガRS)
                    'Call メルマガ本文検索(0, メルマガRS)
                    
                    If Not メルマガRS.EOF Then
                        
                        If !楽天メール <> "" Then
                            宛先 = !楽天メール ' + Chr(9) + "info@cathand.jp"    ' 宛先
                        Else
                            宛先 = !メール ' + Chr(9) + "info@cathand.jp"    ' 宛先
                        End If
                        '件名 = !顧客名 + "様 " + "メールマガジン [" & IIf(!メルマガNO <= 0, 1, !メルマガNO) & "] 第"            ' 件名
                        '件名 = !顧客名 + "様 " + "いつもご利用ありがとうございます"            ' 件名
                        件名 = !顧客名 + "様 " + "育毛・発毛攻略法＆無料プレゼント"            ' 件名
                        
                        メール内容 = ""
                        メール内容 = メール内容 + !顧客名 + "様" + Chr$(13) + Chr$(10)
                        メール内容 = メール内容 + Chr$(13) + Chr$(10)
                        メール内容 = メール内容 + メルマガRS!メルマガ + Chr$(13) + Chr$(10)
                        
                        ' メール送信
                        ret = SendMail(G_サーバ, 宛先, G_送信元, 件名, メール内容, "")
                    
                        If Len(ret) <> 0 Then
                           'Call MsgBox("メール送信エラー：" & ret, vbOKOnly, "顧客管理")
                        End If
                        
                        メルマガ送信予定日 = Format(DateAdd("d", 7, Now), "yyyy/mm/dd")
                        Call メルマガ発行NO更新(!顧客ID, IIf(!メルマガNO <= 0, 2, !メルマガNO + 1), "'" + メルマガ送信予定日 + "'")
                        'Call メルマガ発行NO更新(!顧客ID, -1, "NULL")
                    
                        Sleep (1000 * 3)
                    
                    End If
                    
                    メルマガRS.Close
                    
                End If
#Else
                    
                Call メルマガ本文検索(0, メルマガRS)
                
                If Not メルマガRS.EOF Then
                    
                    If !楽天メール <> "" Then
                        宛先 = !楽天メール ' + Chr(9) + "info@cathand.jp"    ' 宛先
                    Else
                        宛先 = !メール ' + Chr(9) + "info@cathand.jp"    ' 宛先
                    End If
                    件名 = !顧客名 + "様 " + "育毛剤アーデル！プレゼント応募付き！"            ' 件名
                    
                    メール内容 = ""
                    メール内容 = メール内容 + !顧客名 + "様" + Chr$(13) + Chr$(10)
                    メール内容 = メール内容 + Chr$(13) + Chr$(10)
                    メール内容 = メール内容 + メルマガRS!メルマガ + Chr$(13) + Chr$(10)
                    メール内容 = Replace(メール内容, "##########", !顧客ID)
                    
                    ' メール送信
                    ret = SendMail(G_サーバ, 宛先, G_送信元, 件名, メール内容, "")
                
                    If Len(ret) <> 0 Then
                       'Call MsgBox("メール送信エラー：" & ret, vbOKOnly, "顧客管理")
                    End If
                    
                    Sleep (1000 * 3)
                
                End If
                
                メルマガRS.Close

#End If

            End If
            
            .MoveNext
        Loop
        
        .Close
    End With
    
    Call MsgBox("メールを送信しました", vbOKOnly, "顧客管理")
        
    MousePointer = vbNormal
    
End Sub

'************************************************************************
'機  能：テンプレート検索
'************************************************************************
Private Sub cmdテンプレート_Click()
    
    Dim テンプレート As String
    Dim 表示    As String
    
    テンプレート = cmbテンプレート.Text
    
    If テンプレート = "アーデル新規" Then
        表示 = "アーデル" + Chr$(13) + Chr$(10)
        表示 = 表示 + "１０％割引用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "返金用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル資料" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "アーデル＆シャンプーセット新規" Then
        表示 = "アーデル＆シャンプーセット" + Chr$(13) + Chr$(10)
        表示 = 表示 + "１０％割引用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "返金用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "チラシ" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル資料" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "新ブスタ新規" Then
        表示 = "新ブスタ" + Chr$(13) + Chr$(10)
        表示 = 表示 + "返金用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル資料" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "新ブスタ＆シャンプーセット新規" Then
        表示 = "新ブスタ" + Chr$(13) + Chr$(10)
        表示 = 表示 + "シャンプー" + Chr$(13) + Chr$(10)
        表示 = 表示 + "返金用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル資料" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "ブースター新規" Then
        表示 = "ブースター" + Chr$(13) + Chr$(10)
        表示 = 表示 + "返金用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル資料" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "ブースター＆シャンプーセット新規" Then
        表示 = "ブースター" + Chr$(13) + Chr$(10)
        表示 = 表示 + "シャンプー" + Chr$(13) + Chr$(10)
        表示 = 表示 + "返金用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル資料" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "新ハイブリッター新規" Then
        表示 = "新ハイブリッター" + Chr$(13) + Chr$(10)
        表示 = 表示 + "返金用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル資料" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "新ハイブリッター＆シャンプーセット新規" Then
        表示 = "新ハイブリッター" + Chr$(13) + Chr$(10)
        表示 = 表示 + "シャンプー" + Chr$(13) + Chr$(10)
        表示 = 表示 + "返金用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル資料" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "ハイブリッド新規" Then
        表示 = "ハイブリッド" + Chr$(13) + Chr$(10)
        表示 = 表示 + "返金用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル資料" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "ハイブリッド＆シャンプーセット新規" Then
        表示 = "ハイブリッド" + Chr$(13) + Chr$(10)
        表示 = 表示 + "シャンプー" + Chr$(13) + Chr$(10)
        表示 = 表示 + "返金用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル資料" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "ナイスレディー新規" Then
        表示 = "ナイスレディー" + Chr$(13) + Chr$(10)
        表示 = 表示 + "返金用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル資料" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "ナイスレディー＆シャンプーセット新規" Then
        表示 = "ナイスレディー" + Chr$(13) + Chr$(10)
        表示 = 表示 + "シャンプー" + Chr$(13) + Chr$(10)
        表示 = 表示 + "返金用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル資料" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "シャンプー２本セット新規" Then
        表示 = "シャンプー２本" + Chr$(13) + Chr$(10)
        表示 = 表示 + "返金用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル資料" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "シャンプー新規" Then
        表示 = "シャンプー" + Chr$(13) + Chr$(10)
        表示 = 表示 + "返金用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル資料" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "シャンプー＆トリートメント新規" Then
        表示 = "シャンプー＆トリートメント" + Chr$(13) + Chr$(10)
        表示 = 表示 + "返金用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル資料" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "トリートメント新規" Then
        表示 = "トリートメント" + Chr$(13) + Chr$(10)
        表示 = 表示 + "返金用紙" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル資料" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "試供品新規" Then
        表示 = "試供品セット" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル資料" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "アーデル活用" Then
        表示 = "" 'Chr$(13) + Chr$(10)
        表示 = 表示 + "チラシ" + Chr$(13) + Chr$(10)
        表示 = 表示 + "アーデル活用・マニュアル" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "毎日の積み重ねが大切です" Then
        表示 = "" 'Chr$(13) + Chr$(10)
        表示 = 表示 + "チラシ" + Chr$(13) + Chr$(10)
        表示 = 表示 + "毎日の積み重ねが大切です・マニュアル" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "育毛ＤＶＤ" Then
        表示 = "" 'Chr$(13) + Chr$(10)
        表示 = 表示 + "チラシ" + Chr$(13) + Chr$(10)
        表示 = 表示 + "ドクターアーデル・育毛ＤＶＤ" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "育毛と運動" Then
        表示 = "" 'Chr$(13) + Chr$(10)
        表示 = 表示 + "チラシ" + Chr$(13) + Chr$(10)
        表示 = 表示 + "育毛と運動・マニュアル" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If
    
    If テンプレート = "育毛・発毛" Then
        表示 = "" 'Chr$(13) + Chr$(10)
        表示 = 表示 + "チラシ" + Chr$(13) + Chr$(10)
        表示 = 表示 + "育毛・発毛マニュアル" + Chr$(13) + Chr$(10)
        txt備考2.Text = txt備考2.Text + 表示
    End If

    Call 注文_更新

End Sub

'************************************************************************
'機  能：転居ボタン
'************************************************************************
Private Sub cmd転居_Click()
    
    Call 転居更新(txt顧客ID.Text)
    
End Sub

'************************************************************************
'機  能：トランザクションデータの更新
'************************************************************************
Private Sub トランザクションデータの更新()
    
#If 0 Then
    Select Case G_タブNO
        Case 1
            Call 顧客情報_登録
        Case 2
            Call 顧客情報_登録
        Case 3
            Call 注文_更新
    End Select
#End If
End Sub

'************************************************************************
'機  能：入力チェック
'************************************************************************
Private Function 入力チェック(ByVal 入力値 As String) As Boolean

    Dim 位置 As Integer
    
    位置 = InStr(入力値, "'")
    
    If 位置 > 0 Then
        入力チェック = True
    Else
        入力チェック = False
    End If
    
End Function

'************************************************************************
'機  能：IME オン/オフ 切り替え
'************************************************************************
Private Sub psubIMEOnOff(ByVal hwnd As Long, ByVal booOnOff As Boolean)
    Dim himc As Long    'IMEハンドル
    'IMEハンドル取得
    himc = ImmGetContext(hwnd)
    'IME切り替え
    Call ImmSetOpenStatus(himc, booOnOff)
    'IMEハンドル解放
    Call ImmReleaseContext(hwnd, himc)
End Sub

'************************************************************************
'機  能：IMEモードの切り替え
'************************************************************************
Private Sub ImeMode(Index As Integer)
    Dim himc As Long            'IMEハンドル
    Dim lngConversion As Long   '入力モード
    Dim lngSentence As Long     'モード数
    'IMEハンドル取得
    himc = ImmGetContext(Me.hwnd)
    'IMEステータス取得
    If Not ImmGetOpenStatus(himc) Then
        'IME切り替え
        Call ImmSetOpenStatus(himc, 1)
    End If
    'IME入力モード取得
    Call ImmGetConversionStatus(himc, lngConversion, lngSentence)
    'IME入力モード設定
    Select Case Index
    Case 0  '全角ひらがな
        lngConversion = MY_IME_CHMODE_ZEN_HIRA
    Case 1  '全角カタカナ
        lngConversion = MY_IME_CHMODE_ZEN_KATA
    Case 2  '全角英数
        lngConversion = MY_IME_CHMODE_ZEN_EISU
    Case 3  '半角カタカナ
        lngConversion = MY_IME_CHMODE_HAN_KATA
    Case 4  '半角英数
        lngConversion = MY_IME_CHMODE_HAN_EISU
    End Select
    Call ImmSetConversionStatus(himc, lngConversion, lngSentence)
    'IMEハンドル解放
    Call ImmReleaseContext(Me.hwnd, himc)
End Sub

'************************************************************************
'機  能 :注文掲示板更新
'************************************************************************
Private Sub cmd更新_Click()
    
    Dim row             As Long
    Dim 注文元          As String
    Dim 注文件数        As String
    Dim クレジット      As String
    Dim 東京クレジット  As String
    Dim 商品代引        As String
    Dim コンビニ        As String
    Dim 銀行振込        As String
    Dim 楽天バンク決済  As String
    Dim ペイジー        As String
    Dim 後払            As String
    Dim ポイント        As String
    Dim 携帯決済        As String
    Dim 電子マネー      As String
    Dim ヤフオク        As String
    
    If G_店舗名 = "トリニティー楽天市場店" Then
        注文元 = "楽天"
    Else
        注文元 = "Yahoo"
    End If
    
    row = 1
    Call 注文件数取得(注文元, 注文件数, クレジット, 東京クレジット, 商品代引, コンビニ, 銀行振込, 楽天バンク決済, ペイジー, 後払, ポイント, 携帯決済, 電子マネー, ヤフオク)
    
    Call SpreadSetVal(va注文掲示板, row, 2, 注文件数)
    Call SpreadSetVal(va注文掲示板, row, 3, クレジット)
    Call SpreadSetVal(va注文掲示板, row, 4, 東京クレジット)
    Call SpreadSetVal(va注文掲示板, row, 5, 商品代引)
    Call SpreadSetVal(va注文掲示板, row, 6, コンビニ)
    Call SpreadSetVal(va注文掲示板, row, 7, 銀行振込)
    Call SpreadSetVal(va注文掲示板, row, 8, 楽天バンク決済)
    Call SpreadSetVal(va注文掲示板, row, 9, ペイジー)
    Call SpreadSetVal(va注文掲示板, row, 10, 後払)
    Call SpreadSetVal(va注文掲示板, row, 11, ポイント)
    Call SpreadSetVal(va注文掲示板, row, 12, 携帯決済)
    Call SpreadSetVal(va注文掲示板, row, 13, 電子マネー)
    Call SpreadSetVal(va注文掲示板, row, 14, ヤフオク)
    
    ' 自社
    row = row + 1
    Call 注文件数取得("自社サイト", 注文件数, クレジット, 東京クレジット, 商品代引, コンビニ, 銀行振込, 楽天バンク決済, ペイジー, 後払, ポイント, 携帯決済, 電子マネー, ヤフオク)
    
    Call SpreadSetVal(va注文掲示板, row, 2, 注文件数)
    Call SpreadSetVal(va注文掲示板, row, 3, クレジット)
    Call SpreadSetVal(va注文掲示板, row, 4, 東京クレジット)
    Call SpreadSetVal(va注文掲示板, row, 5, 商品代引)
    Call SpreadSetVal(va注文掲示板, row, 6, コンビニ)
    Call SpreadSetVal(va注文掲示板, row, 7, 銀行振込)
    Call SpreadSetVal(va注文掲示板, row, 8, 楽天バンク決済)
    Call SpreadSetVal(va注文掲示板, row, 9, ペイジー)
    Call SpreadSetVal(va注文掲示板, row, 10, 後払)
    Call SpreadSetVal(va注文掲示板, row, 11, ポイント)
    Call SpreadSetVal(va注文掲示板, row, 12, 携帯決済)
    Call SpreadSetVal(va注文掲示板, row, 13, 電子マネー)
    Call SpreadSetVal(va注文掲示板, row, 14, ヤフオク)
        
    ' アマゾン
    row = row + 1
    Call 注文件数取得("アマゾン", 注文件数, クレジット, 東京クレジット, 商品代引, コンビニ, 銀行振込, 楽天バンク決済, ペイジー, 後払, ポイント, 携帯決済, 電子マネー, ヤフオク)
    
    Call SpreadSetVal(va注文掲示板, row, 2, 注文件数)
    Call SpreadSetVal(va注文掲示板, row, 3, クレジット)
    Call SpreadSetVal(va注文掲示板, row, 4, 東京クレジット)
    Call SpreadSetVal(va注文掲示板, row, 5, 商品代引)
    Call SpreadSetVal(va注文掲示板, row, 6, コンビニ)
    Call SpreadSetVal(va注文掲示板, row, 7, 銀行振込)
    Call SpreadSetVal(va注文掲示板, row, 8, 楽天バンク決済)
    Call SpreadSetVal(va注文掲示板, row, 9, ペイジー)
    Call SpreadSetVal(va注文掲示板, row, 10, 後払)
    Call SpreadSetVal(va注文掲示板, row, 11, ポイント)
    Call SpreadSetVal(va注文掲示板, row, 12, 携帯決済)
    Call SpreadSetVal(va注文掲示板, row, 13, 電子マネー)
    Call SpreadSetVal(va注文掲示板, row, 14, ヤフオク)
    
    ' レントラックス
    row = row + 1
    Call 注文件数取得("レントラックス", 注文件数, クレジット, 東京クレジット, 商品代引, コンビニ, 銀行振込, 楽天バンク決済, ペイジー, 後払, ポイント, 携帯決済, 電子マネー, ヤフオク)
'    Call 注文件数取得("おちゃのこネット", 注文件数, クレジット, 商品代引, コンビニ, 銀行振込, 楽天バンク決済, ペイジー, 後払, ポイント, 携帯決済)
    
    Call SpreadSetVal(va注文掲示板, row, 2, 注文件数)
    Call SpreadSetVal(va注文掲示板, row, 3, クレジット)
    Call SpreadSetVal(va注文掲示板, row, 4, 東京クレジット)
    Call SpreadSetVal(va注文掲示板, row, 5, 商品代引)
    Call SpreadSetVal(va注文掲示板, row, 6, コンビニ)
    Call SpreadSetVal(va注文掲示板, row, 7, 銀行振込)
    Call SpreadSetVal(va注文掲示板, row, 8, 楽天バンク決済)
    Call SpreadSetVal(va注文掲示板, row, 9, ペイジー)
    Call SpreadSetVal(va注文掲示板, row, 10, 後払)
    Call SpreadSetVal(va注文掲示板, row, 11, ポイント)
    Call SpreadSetVal(va注文掲示板, row, 12, 携帯決済)
    Call SpreadSetVal(va注文掲示板, row, 13, 電子マネー)
    Call SpreadSetVal(va注文掲示板, row, 14, ヤフオク)
    
    ' ヤフオク
    row = row + 1
    Call 注文件数取得("ヤフオク", 注文件数, クレジット, 東京クレジット, 商品代引, コンビニ, 銀行振込, 楽天バンク決済, ペイジー, 後払, ポイント, 携帯決済, 電子マネー, ヤフオク)
    
    Call SpreadSetVal(va注文掲示板, row, 2, 注文件数)
    Call SpreadSetVal(va注文掲示板, row, 3, クレジット)
    Call SpreadSetVal(va注文掲示板, row, 4, 東京クレジット)
    Call SpreadSetVal(va注文掲示板, row, 5, 商品代引)
    Call SpreadSetVal(va注文掲示板, row, 6, コンビニ)
    Call SpreadSetVal(va注文掲示板, row, 7, 銀行振込)
    Call SpreadSetVal(va注文掲示板, row, 8, 楽天バンク決済)
    Call SpreadSetVal(va注文掲示板, row, 9, ペイジー)
    Call SpreadSetVal(va注文掲示板, row, 10, 後払)
    Call SpreadSetVal(va注文掲示板, row, 11, ポイント)
    Call SpreadSetVal(va注文掲示板, row, 12, 携帯決済)
    Call SpreadSetVal(va注文掲示板, row, 13, 電子マネー)
    Call SpreadSetVal(va注文掲示板, row, 14, ヤフオク)

End Sub

'************************************************************************
'機  能 :郵便番号検索
'************************************************************************
Private Sub cmd郵便番号_Click()
    
    Dim ADF025      As New ADF025
    Dim i           As Integer
    Dim 郵便番号    As String
    
    Call ADF025.Show(1)
    
    郵便番号 = ADF025.get郵便番号()
    If 郵便番号 <> "" Then
        txt郵便番号.Text = 郵便番号
    End If
    
End Sub

'************************************************************************
'機  能 :問合番号セット
'************************************************************************
Private Sub cmd問合番号_Click()

    Dim cCsvReader  As CsvReader
    Set cCsvReader = New CsvReader
    Dim 顧客ID      As String
    Dim 出荷日      As String
    Dim 問合番号    As String
    Dim 削除区分    As String
    
    MousePointer = vbHourglass
    
    ' 指定した CSV ファイルを開く
    If cCsvReader.OpenStream("c:\顧客管理\配送履歴.csv") = False Then
        MousePointer = vbNormal
        Call MsgBox("配送履歴がありません。", vbOK, "顧客管理")
        Exit Sub
    End If
    
    ' 最初の行をヘッダとして読み込む
    Call cCsvReader.ReadHeader

    ' CSV ファイルの中身をすべて取得する
    Dim cTable As Collection
    Set cTable = cCsvReader.ReadToEnd()

    ' すべての中身 (Table) から 行 (Row) を列挙して取り出す
    Dim cRow As Collection
    
    For Each cRow In cTable
        ' 行からカラム名を使って各 Item を出力する
        On Error GoTo skip
        顧客ID = cRow("住所録コード")
        If 顧客ID = "" Then
            Exit For
        End If
        顧客ID = Format(CLng(顧客ID), "00000")
        出荷日 = cRow("出荷日時")
        問合番号 = 問合番号編集(cRow("お問合せ送り状№"))
        削除区分 = cRow("削除区分")
        If 削除区分 = "0" Then
            Call 問い合わせ番号更新(顧客ID, 問合番号)
        End If
    Next
    
skip:
   
    Call MsgBox("問合番号を読み込みました。", vbOKOnly, "顧客管理")
    
    MousePointer = vbNormal

End Sub

'************************************************************************
'機  能 :楽天→Yahoo移行
'************************************************************************
Private Sub cmd移行_Click()
    
    Dim 顧客ID      As String
    Dim 顧客名      As String
    Dim 顧客マスタRS As New ADODB.Recordset
    Dim ADF027      As New ADF027

    Call ADF027.Show(1)
    
    Call ADF027.顧客ID取得(顧客ID, 顧客名)
    
    cmb検索条件.Text = "顧客名"
    txt検索条件.Text = 顧客名
    
    MousePointer = vbHourglass
    
    Call 顧客検索7(顧客ID, 顧客マスタRS)
    
    Call 顧客リスト表示(顧客マスタRS)
    
    MousePointer = vbNormal
    
End Sub

