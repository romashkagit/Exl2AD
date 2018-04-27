object Exl2ADfm: TExl2ADfm
  Left = 0
  Top = 0
  Caption = 'Exl2ADfm'
  ClientHeight = 454
  ClientWidth = 514
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object pnlBtn: TPanel
    Left = 0
    Top = 0
    Width = 514
    Height = 27
    Align = alTop
    AutoSize = True
    TabOrder = 0
    object btnLoadExl: TButton
      Left = 1
      Top = 1
      Width = 512
      Height = 25
      Align = alTop
      Caption = #1047#1072#1075#1088#1091#1079#1080#1090#1100' '#1080#1079' Excel'
      TabOrder = 0
      OnClick = btnLoadExlClick
    end
  end
  object reLog: TRichEdit
    Left = 0
    Top = 27
    Width = 514
    Height = 427
    Align = alClient
    Color = clNone
    Font.Charset = RUSSIAN_CHARSET
    Font.Color = clLime
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    Lines.Strings = (
      '')
    ParentFont = False
    ScrollBars = ssBoth
    TabOrder = 1
    Zoom = 100
  end
  object qCn: TADOQuery
    Connection = adoCn
    CursorLocation = clUseServer
    CursorType = ctStatic
    LockType = ltReadOnly
    ParamCheck = False
    Parameters = <>
    SQL.Strings = (
      
        'select  cn, userPrincipalName, department, title, sAMAccountName' +
        ',TelephoneNumber,'
      'employeeType, displayName, physicalDeliveryOfficeName, location'
      'from '#39'LDAP://NEOD.IN'#39
      'where cn='#39'test'#39
      '')
    Left = 255
    Top = 65534
  end
  object adoCn: TADOConnection
    Connected = True
    ConnectionString = 
      'Provider=ADsDSOObject;Encrypt Password=False;Mode=Read;Bind Flag' +
      's=0'
    LoginPrompt = False
    Mode = cmRead
    Provider = 'ADsDSOObject'
    Left = 287
    Top = 4
  end
end
