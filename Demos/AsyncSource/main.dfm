object FormMain: TFormMain
  Left = 366
  Top = 402
  AutoScroll = False
  Caption = 'Async Data Transfer Demo - Drop source'
  ClientHeight = 383
  ClientWidth = 573
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Arial'
  Font.Style = []
  OldCreateOrder = False
  Position = poDefaultPosOnly
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 14
  object Panel2: TPanel
    Left = 0
    Top = 117
    Width = 573
    Height = 232
    Cursor = crHandPoint
    Align = alClient
    BevelOuter = bvNone
    Caption = ' '
    ParentColor = True
    TabOrder = 0
    object Panel4: TPanel
      Left = 109
      Top = 0
      Width = 464
      Height = 232
      Align = alClient
      BevelOuter = bvLowered
      BorderWidth = 8
      Caption = ' '
      TabOrder = 1
      object Label2: TLabel
        Left = 26
        Top = 24
        Width = 423
        Height = 93
        Anchors = [akLeft, akTop, akRight]
        AutoSize = False
        Caption = 
          'During a normal data transfer the source application will be blo' +
          'cked while the transfer takes place because the applicaion is un' +
          'able to process events.'#13'Notice that the timer stops and the form' +
          ' can'#39't be moved or repainted during a normal data transfer.'
        ShowAccelChar = False
        WordWrap = True
      end
      object Label3: TLabel
        Left = 26
        Top = 136
        Width = 427
        Height = 85
        Anchors = [akLeft, akTop, akRight]
        AutoSize = False
        Caption = 
          'During an asynchronous data transfer the data transfer is perfor' +
          'med in a separate thread so the source application appears unaff' +
          'ected by the transfer.'#13'Notice that the timer continues and the f' +
          'orm can be moved and repainted during an asynchronous data trans' +
          'fer.'
        ShowAccelChar = False
        WordWrap = True
      end
      object RadioButtonNormal: TRadioButton
        Left = 8
        Top = 8
        Width = 353
        Height = 17
        Caption = 'Normal synchronous transfer'
        Checked = True
        TabOrder = 0
        TabStop = True
      end
      object RadioButtonAsync: TRadioButton
        Left = 8
        Top = 120
        Width = 361
        Height = 17
        Caption = 'Asynchronous transfer'
        TabOrder = 1
      end
    end
    object Panel3: TPanel
      Left = 0
      Top = 0
      Width = 109
      Height = 232
      Align = alLeft
      BevelOuter = bvLowered
      Caption = ' '
      ParentColor = True
      TabOrder = 0
      object PaintBoxPie: TPaintBox
        Left = 1
        Top = 1
        Width = 107
        Height = 230
        Align = alClient
      end
    end
  end
  object ProgressBar1: TProgressBar
    Left = 0
    Top = 349
    Width = 573
    Height = 15
    Align = alBottom
    Min = 0
    Max = 1000000
    Step = 1024
    TabOrder = 1
  end
  object Panel5: TPanel
    Left = 0
    Top = 0
    Width = 573
    Height = 117
    Align = alTop
    BevelOuter = bvLowered
    Caption = ' '
    TabOrder = 2
    object Label1: TLabel
      Left = 115
      Top = 1
      Width = 457
      Height = 115
      Align = alClient
      Caption = 
        'This example demonstrates how to source a drop from a thread.'#13#13'T' +
        'he advantage of using a thread is that the source application is' +
        'n´t blocked while the target application processes the drop.'#13#13'No' +
        'te that this approach is normally only used when transferring la' +
        'rge amounts of data or when the drop target is very slow.'
      ShowAccelChar = False
      WordWrap = True
    end
    object Label4: TLabel
      Left = 1
      Top = 1
      Width = 114
      Height = 115
      Cursor = crHandPoint
      Align = alLeft
      Alignment = taCenter
      Caption = 'Drag from here and onto Desktop or Explorer'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -13
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      ShowAccelChar = False
      Layout = tlCenter
      WordWrap = True
      OnMouseDown = OnMouseDown
    end
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 364
    Width = 573
    Height = 19
    Panels = <>
    SimplePanel = True
    OnResize = StatusBar1Resize
    object ButtonAbort: TSpeedButton
      Left = 456
      Top = 3
      Width = 104
      Height = 15
      Anchors = [akRight, akBottom]
      Caption = 'Abort transfer'
      Flat = True
      Glyph.Data = {
        C2010000424DC20100000000000036000000280000000B0000000B0000000100
        1800000000008C010000C21E0000C21E00000000000000000000FF00FFFF00FF
        C1C6D75D65951F216B0F115F161853474C66B0B3B7FF00FFFF00FF000000FF00
        FFBEC1DD282B9502038A00008E00008E0000880101711A1C56A3A8A9FF00FF00
        0000CED1EA292CA82323A33636B40505A60202A51B1BA53C3CAA06068623255C
        BDC1C4000000797DCD0202A87B7BB5D4D4ED3E3EC31C1CB5A7A7D8BCBCDA1C1C
        A1040475697186000000282ABB0303BC2121BCA4A4C7DBDBF2C5C5EED6D6E644
        44B90505B302029A3E426F0000001011C20909CF0404D33A3AC4EBEBF4FFFFFF
        7D7DD60000C60202C30303B0383D770000002426D61515E72525E1A4A4DFDEDE
        E9CDCDE3D1D1F03E3EDB0A0AD20808BA51578F0000005B61E82828F88C8CD6D2
        D2DD4444D31C1CCFA0A0C8B6B6E12626E11819BD9AA0C6000000B6B9F74546FA
        6969E46363DD1C1CF61111F43636E25A5AD62929EE5156C9E0E3EB000000FF00
        FFA5AAF96264FB8080FF8383FF7373FD6162FB4546FC565BECD7DCECFF00FF00
        0000FF00FFFF00FFBBBEFA0808BA0808BA0808BA696EF8979CF62929EEFF00FF
        FF00FF000000}
      Visible = False
      OnClick = ButtonAbortClick
    end
  end
  object Timer1: TTimer
    Interval = 100
    OnTimer = Timer1Timer
    Left = 8
    Top = 160
  end
  object DropEmptySource1: TDropEmptySource
    DragTypes = [dtCopy, dtMove]
    OnDrop = DropEmptySource1Drop
    OnAfterDrop = DropEmptySource1AfterDrop
    OnGetData = DropEmptySource1GetData
    AllowAsyncTransfer = True
    Left = 40
    Top = 161
  end
  object DataFormatAdapterSource: TDataFormatAdapter
    DragDropComponent = DropEmptySource1
    DataFormatName = 'TVirtualFileStreamDataFormat'
    Left = 40
    Top = 193
  end
end
