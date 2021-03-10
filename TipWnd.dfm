object frmTipWnd: TfrmTipWnd
  Left = 736
  Height = 129
  Top = 497
  Width = 500
  BorderIcons = []
  BorderStyle = bsNone
  ClientHeight = 129
  ClientWidth = 500
  Color = clInfoBk
  Constraints.MaxHeight = 600
  Constraints.MaxWidth = 800
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  FormStyle = fsStayOnTop
  KeyPreview = True
  OnClose = FormClose
  OnCloseQuery = FormCloseQuery
  LCLVersion = '2.0.10.0'
  object Panel1: TPanel
    Left = 0
    Height = 129
    Top = 0
    Width = 500
    Align = alClient
    BevelOuter = bvNone
    BevelWidth = 2
    ClientHeight = 129
    ClientWidth = 500
    TabOrder = 0
    object Label1: TLabel
      Left = 5
      Height = 13
      Top = 5
      Width = 31
      Caption = 'Label1'
      Font.Color = clInfoText
      Font.Height = -11
      Font.Name = 'Tahoma'
      ParentColor = False
      ParentFont = False
      Visible = False
    end
    object Tipinfo: TMemo
      Left = 5
      Height = 33
      Top = 5
      Width = 113
      BorderStyle = bsNone
      Color = clInfoBk
      Font.Color = clInfoText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Lines.Strings = (
        'Tipinfo'
      )
      ParentFont = False
      TabOrder = 0
    end
  end
end
