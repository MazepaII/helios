object Form2: TForm2
  Left = 481
  Top = 173
  AutoScroll = False
  Caption = #1054#1090#1095#1077#1090
  ClientHeight = 512
  ClientWidth = 659
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Calibri'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  OnCreate = FormCreate
  OnResize = FormResize
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Memo1: TMemo
    Left = 0
    Top = 0
    Width = 554
    Height = 543
    BevelInner = bvNone
    BiDiMode = bdLeftToRight
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Courier New'
    Font.Style = []
    ParentBiDiMode = False
    ParentFont = False
    ScrollBars = ssBoth
    TabOrder = 0
  end
  object MainMenu1: TMainMenu
    Left = 24
    Top = 16
    object A1: TMenuItem
      Caption = #1060#1072#1081#1083
      object N3: TMenuItem
        Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100
        OnClick = N3Click
      end
      object N4: TMenuItem
        Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100' '#1082#1072#1082'...'
        OnClick = N4Click
      end
      object N5: TMenuItem
        Caption = #1055#1077#1095#1072#1090#1100'...'
        OnClick = N5Click
      end
      object N1: TMenuItem
        Caption = #1042#1099#1093#1086#1076
        OnClick = N1Click
      end
    end
    object N2: TMenuItem
      Caption = #1057#1087#1088#1072#1074#1082#1072
      object N6: TMenuItem
        Caption = #1054' '#1087#1088#1086#1075#1088#1072#1084#1084#1077
        OnClick = N6Click
      end
    end
  end
  object SaveDialog1: TSaveDialog
    Filter = #1058#1077#1082#1089#1090#1086#1074#1099#1081' '#1076#1086#1082#1091#1084#1077#1090'(*.txt)|*.txt|'#1042#1089#1077' '#1092#1072#1081#1083#1099' (*.*)|*.*'
    Left = 64
    Top = 16
  end
  object PrintDialog1: TPrintDialog
    Left = 104
    Top = 16
  end
end
