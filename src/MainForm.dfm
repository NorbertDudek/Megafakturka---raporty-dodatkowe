object frmMain: TfrmMain
  Left = 0
  Top = 0
  Caption = 'Raport zakup'#243'w i sprzeda'#380'y'
  ClientHeight = 640
  ClientWidth = 1090
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  TextHeight = 15
  object pnlTop: TPanel
    Left = 0
    Top = 0
    Width = 1090
    Height = 138
    Align = alTop
    BevelOuter = bvNone
    Padding.Left = 10
    Padding.Top = 8
    Padding.Right = 10
    Padding.Bottom = 8
    TabOrder = 0
    object lblFile: TLabel
      Left = 14
      Top = 14
      Width = 78
      Height = 15
      Caption = 'Plik bazy MDB:'
    end
    object lblFrom: TLabel
      Left = 14
      Top = 50
      Width = 19
      Height = 15
      Caption = 'Od:'
    end
    object lblTo: TLabel
      Left = 200
      Top = 50
      Width = 18
      Height = 15
      Caption = 'Do:'
    end
    object lblReportType: TLabel
      Left = 14
      Top = 87
      Width = 63
      Height = 15
      Caption = 'Typ raportu:'
    end
    object edtFile: TEdit
      Left = 90
      Top = 11
      Width = 740
      Height = 23
      TabOrder = 0
    end
    object btnBrowse: TButton
      Left = 836
      Top = 10
      Width = 70
      Height = 25
      Caption = 'Przegl'#261'daj...'
      TabOrder = 1
      OnClick = btnBrowseClick
    end
    object dtpFrom: TDateTimePicker
      Left = 44
      Top = 47
      Width = 140
      Height = 23
      Date = 45413.000000000000000000
      Time = 45413.000000000000000000
      TabOrder = 2
    end
    object dtpTo: TDateTimePicker
      Left = 224
      Top = 47
      Width = 140
      Height = 23
      Date = 45413.000000000000000000
      Time = 45413.000000000000000000
      TabOrder = 3
    end
    object btnRange: TButton
      Left = 374
      Top = 46
      Width = 130
      Height = 25
      Caption = 'Szybki wyb'#243'r '#9660
      TabOrder = 4
      OnClick = btnRangeClick
    end
    object btnGenerate: TButton
      Left = 524
      Top = 46
      Width = 130
      Height = 25
      Caption = 'Generuj raport'
      Default = True
      TabOrder = 5
      OnClick = btnGenerateClick
    end
    object btnSave: TButton
      Left = 660
      Top = 46
      Width = 110
      Height = 25
      Caption = 'Zapisz HTML...'
      TabOrder = 6
      OnClick = btnSaveClick
    end
    object btnOpenBrowser: TButton
      Left = 776
      Top = 46
      Width = 130
      Height = 25
      Caption = 'Otw'#243'rz w przegl'#261'darce'
      TabOrder = 7
      OnClick = btnOpenBrowserClick
    end
    object btnSavePdf: TButton
      Left = 916
      Top = 46
      Width = 150
      Height = 25
      Caption = 'Zapisz jako PDF (A4)'
      TabOrder = 10
      OnClick = btnSavePdfClick
    end
    object cmbReportType: TComboBox
      Left = 110
      Top = 83
      Width = 250
      Height = 23
      Style = csDropDownList
      TabOrder = 8
      Items.Strings = (
        'Zakupy (FZ / KFZ)'
        'Sprzeda'#380' (FV / KFV)')
    end
    object pnlStatus: TPanel
      Left = 10
      Top = 113
      Width = 1070
      Height = 17
      Align = alBottom
      Alignment = taLeftJustify
      BevelOuter = bvNone
      Caption = '   Gotowy.'
      ParentBackground = False
      TabOrder = 9
    end
  end
  object memHTML: TMemo
    Left = 0
    Top = 138
    Width = 1090
    Height = 502
    Align = alClient
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Consolas'
    Font.Style = []
    ParentFont = False
    ScrollBars = ssBoth
    TabOrder = 1
    WordWrap = False
  end
  object dlgOpen: TOpenDialog
    DefaultExt = 'mdb'
    Filter = 
      'Bazy Access (*.mdb;*.accdb)|*.mdb;*.accdb|Wszystkie pliki (*.*)|' +
      '*.*'
    Title = 'Wybierz plik bazy danych'
    Left = 920
    Top = 8
  end
  object dlgSave: TSaveDialog
    DefaultExt = 'html'
    Filter = 'Plik HTML (*.html)|*.html|Wszystkie pliki (*.*)|*.*'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofPathMustExist, ofEnableSizing]
    Title = 'Zapisz raport'
    Left = 920
    Top = 40
  end
  object popRange: TPopupMenu
    Left = 480
    Top = 8
    object miToday: TMenuItem
      Tag = 1
      Caption = 'Dzie'#324' dzisiejszy'
      OnClick = RangeMenuClick
    end
    object miYesterday: TMenuItem
      Tag = 2
      Caption = 'Dzie'#324' wczorajszy'
      OnClick = RangeMenuClick
    end
    object miSep1: TMenuItem
      Caption = '-'
    end
    object miCurMonth: TMenuItem
      Tag = 6
      Caption = 'Bie'#380#261'cy miesi'#261'c'
      OnClick = RangeMenuClick
    end
    object miPrevMonth: TMenuItem
      Tag = 3
      Caption = 'Poprzedni miesi'#261'c'
      OnClick = RangeMenuClick
    end
    object miCurQuarter: TMenuItem
      Tag = 7
      Caption = 'Bie'#380#261'cy kwarta'#322
      OnClick = RangeMenuClick
    end
    object miPrevQuarter: TMenuItem
      Tag = 4
      Caption = 'Poprzedni kwarta'#322
      OnClick = RangeMenuClick
    end
    object miCurYear: TMenuItem
      Tag = 8
      Caption = 'Bie'#380#261'cy rok'
      OnClick = RangeMenuClick
    end
    object miPrevYear: TMenuItem
      Tag = 5
      Caption = 'Poprzedni rok'
      OnClick = RangeMenuClick
    end
    object miSep2: TMenuItem
      Caption = '-'
    end
    object miLast30: TMenuItem
      Tag = 9
      Caption = 'Ostatnie 30 dni'
      OnClick = RangeMenuClick
    end
    object miLast90: TMenuItem
      Tag = 10
      Caption = 'Ostatnie 90 dni'
      OnClick = RangeMenuClick
    end
    object miLast120: TMenuItem
      Tag = 11
      Caption = 'Ostatnie 120 dni'
      OnClick = RangeMenuClick
    end
  end
  object FDConn: TFDConnection
    LoginPrompt = False
    Left = 920
    Top = 72
  end
end
