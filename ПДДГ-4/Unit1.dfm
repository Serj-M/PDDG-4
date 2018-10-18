object Form1: TForm1
  Left = 190
  Top = 105
  Width = 533
  Height = 487
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 80
    Top = 8
    Width = 82
    Height = 13
    Caption = #1053#1086#1084#1077#1088' '#1087#1088#1080#1073#1086#1088#1072':'
  end
  object Button1: TButton
    Left = 424
    Top = 120
    Width = 75
    Height = 25
    Caption = #1055#1088#1086#1090#1086#1082#1086#1083
    TabOrder = 0
    OnClick = Button1Click
  end
  object Memo1: TMemo
    Left = 8
    Top = 280
    Width = 505
    Height = 169
    Lines.Strings = (
      'Memo1')
    TabOrder = 1
  end
  object Memo2: TMemo
    Left = 8
    Top = 72
    Width = 185
    Height = 193
    Lines.Strings = (
      'Memo2')
    TabOrder = 2
  end
  object Button2: TButton
    Left = 416
    Top = 240
    Width = 89
    Height = 25
    Caption = #1048#1079' .gpf '#1074' Memo1'
    TabOrder = 3
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 8
    Top = 40
    Width = 105
    Height = 25
    Caption = #1042#1077#1097#1077#1089#1074#1072' '#1074' '#1084#1072#1089#1089#1080#1074
    TabOrder = 4
    OnClick = Button3Click
  end
  object Memo3: TMemo
    Left = 208
    Top = 72
    Width = 185
    Height = 193
    Lines.Strings = (
      'Memo3')
    TabOrder = 5
  end
  object Button4: TButton
    Left = 424
    Top = 160
    Width = 75
    Height = 25
    Caption = #1055#1072#1089#1087#1086#1088#1090
    TabOrder = 6
    OnClick = Button4Click
  end
  object Edit1: TEdit
    Left = 168
    Top = 8
    Width = 113
    Height = 21
    TabOrder = 7
    Text = #1042#1074#1077#1076#1080#1090#1077' '#8470#1087#1088#1080#1073#1086#1088#1072
    OnChange = Edit1Change
  end
end
