object FormMain: TFormMain
  Left = 288
  Top = 117
  BorderStyle = bsDialog
  Caption = 'Custom clipboard format demo 1'
  ClientHeight = 229
  ClientWidth = 400
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDefaultPosOnly
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 201
    Height = 96
    Align = alLeft
    BevelInner = bvLowered
    BevelOuter = bvNone
    BorderWidth = 4
    Caption = ' '
    TabOrder = 0
    object PanelSource: TPanel
      Left = 5
      Top = 24
      Width = 191
      Height = 67
      Cursor = crHandPoint
      Align = alClient
      Caption = '00:00:00.000'
      TabOrder = 0
      OnMouseDown = PanelSourceMouseDown
    end
    object Panel4: TPanel
      Left = 5
      Top = 5
      Width = 191
      Height = 19
      Align = alTop
      Caption = 'Drag source'
      TabOrder = 1
    end
  end
  object Panel2: TPanel
    Left = 201
    Top = 0
    Width = 199
    Height = 96
    Align = alClient
    BevelInner = bvLowered
    BevelOuter = bvNone
    BorderWidth = 4
    Caption = ' '
    TabOrder = 1
    object PanelDest: TPanel
      Left = 5
      Top = 24
      Width = 189
      Height = 67
      Align = alClient
      Caption = 'Drop here'
      TabOrder = 0
    end
    object Panel5: TPanel
      Left = 5
      Top = 5
      Width = 189
      Height = 19
      Align = alTop
      Caption = 'Drop target'
      TabOrder = 1
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 96
    Width = 400
    Height = 133
    Align = alBottom
    BevelOuter = bvNone
    BorderWidth = 4
    Caption = ' '
    TabOrder = 2
    object Memo1: TMemo
      Left = 4
      Top = 4
      Width = 392
      Height = 125
      Align = alClient
      BorderStyle = bsNone
      Lines.Strings = (
        
          'This application demonstrates how to define and drag custom clip' +
          'board formats.'
        ''
        
          'The custom format stores the time-of-day and a color value in a ' +
          'structure. The '
        
          'TGenericDataFormat class is used to add support for this format ' +
          'to the '
        'TDropTextSource and TDropTextTarget components.'
        ''
        
          'To see the custom clipboard format in action, drag from the left' +
          ' panel above and '
        
          'drop on the right. You can also do this between multiple instanc' +
          'es of this '
        'application.'
        '')
      ParentColor = True
      ReadOnly = True
      TabOrder = 0
      WantReturns = False
    end
  end
  object Timer1: TTimer
    OnTimer = Timer1Timer
    Left = 14
    Top = 74
  end
  object DropTextSource1: TDropTextSource
    DragTypes = [dtCopy]
    ImageIndex = 0
    ShowImage = False
    ImageHotSpotX = 16
    ImageHotSpotY = 16
    Left = 16
    Top = 32
  end
  object DropTextTarget1: TDropTextTarget
    Dragtypes = [dtCopy, dtLink]
    GetDataOnEnter = False
    OnDrop = DropTextTarget1Drop
    ShowImage = True
    Target = PanelDest
    Left = 216
    Top = 32
  end
end
