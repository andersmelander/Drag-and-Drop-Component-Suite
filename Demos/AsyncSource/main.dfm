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
  OnDestroy = FormDestroy
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
      Left = 513
      Top = 3
      Width = 59
      Height = 15
      Anchors = [akRight, akBottom]
      Caption = 'Abort'
      Flat = True
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = [fsBold]
      ParentFont = False
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
