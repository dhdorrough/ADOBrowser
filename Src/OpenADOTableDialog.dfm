object frmOpenADOTable: TfrmOpenADOTable
  Left = 305
  Top = 495
  Width = 364
  Height = 258
  Caption = 'Open ADO Table'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  DesignSize = (
    356
    231)
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 8
    Width = 84
    Height = 13
    Caption = '&Connection String'
    FocusControl = memoConnectionString
  end
  object Label2: TLabel
    Left = 8
    Top = 152
    Width = 68
    Height = 13
    Anchors = [akLeft, akBottom]
    Caption = '&Table to Open'
    FocusControl = cbTableList
  end
  object btnOk: TButton
    Left = 190
    Top = 198
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = '&Ok'
    Default = True
    Enabled = False
    ModalResult = 1
    TabOrder = 0
  end
  object btnCancel: TButton
    Left = 270
    Top = 198
    Width = 75
    Height = 25
    Anchors = [akRight, akBottom]
    Cancel = True
    Caption = '&Cancel'
    ModalResult = 2
    TabOrder = 1
  end
  object memoConnectionString: TMemo
    Left = 8
    Top = 24
    Width = 337
    Height = 89
    Anchors = [akLeft, akTop, akRight, akBottom]
    Lines.Strings = (
      'memoConnectionString')
    TabOrder = 2
    OnChange = memoConnectionStringChange
  end
  object btnBuild: TButton
    Left = 206
    Top = 118
    Width = 139
    Height = 25
    Anchors = [akRight, akBottom]
    Caption = '&Build Connection String'
    TabOrder = 3
    OnClick = btnBuildClick
  end
  object cbTableList: TComboBox
    Left = 8
    Top = 168
    Width = 145
    Height = 21
    Style = csDropDownList
    Anchors = [akLeft, akBottom]
    ItemHeight = 13
    TabOrder = 4
    OnClick = cbTableListClick
    OnEnter = cbTableListEnter
  end
end
