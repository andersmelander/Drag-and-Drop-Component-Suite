object FormMain: TFormMain
  Left = 356
  Top = 128
  ActiveControl = ListViewHeader
  AutoScroll = False
  Caption = 'Netscape Email/News message target demo'
  ClientHeight = 400
  ClientWidth = 597
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDefaultPosOnly
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 597
    Height = 73
    Align = alTop
    BevelOuter = bvNone
    Caption = ' '
    TabOrder = 0
    object Label1: TLabel
      Left = 8
      Top = 20
      Width = 25
      Height = 13
      Caption = 'URL:'
    end
    object Label2: TLabel
      Left = 8
      Top = 4
      Width = 419
      Height = 13
      Caption = 
        'Drag an email or news message from Netscape and drop it in this ' +
        'window.'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object Edit1: TEdit
      Left = 8
      Top = 36
      Width = 581
      Height = 21
      Anchors = [akLeft, akTop, akRight]
      ReadOnly = True
      TabOrder = 0
    end
  end
  object Panel2: TPanel
    Left = 0
    Top = 73
    Width = 597
    Height = 291
    Align = alClient
    BevelOuter = bvNone
    BorderWidth = 8
    Caption = ' '
    TabOrder = 1
    object PageControl1: TPageControl
      Left = 8
      Top = 8
      Width = 581
      Height = 275
      ActivePage = TabSheetHeader
      Align = alClient
      TabOrder = 0
      object TabSheetHeader: TTabSheet
        Caption = 'Header'
        ImageIndex = 2
        object ListViewHeader: TListView
          Left = 0
          Top = 0
          Width = 573
          Height = 247
          Align = alClient
          Columns = <
            item
              Caption = 'Field'
              Width = 150
            end
            item
              AutoSize = True
              Caption = 'Value'
            end>
          ColumnClick = False
          ReadOnly = True
          RowSelect = True
          ParentColor = True
          TabOrder = 0
          ViewStyle = vsReport
        end
      end
      object TabSheetContents: TTabSheet
        Caption = 'Contents'
        object MemoContent: TMemo
          Left = 0
          Top = 0
          Width = 573
          Height = 247
          TabStop = False
          Align = alClient
          ReadOnly = True
          ScrollBars = ssVertical
          TabOrder = 0
          WantReturns = False
        end
      end
      object TabSheetRaw: TTabSheet
        Caption = 'Raw format'
        ImageIndex = 1
        object MemoRaw: TMemo
          Left = 0
          Top = 0
          Width = 573
          Height = 247
          TabStop = False
          Align = alClient
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -11
          Font.Name = 'Courier New'
          Font.Style = []
          ParentFont = False
          ReadOnly = True
          ScrollBars = ssBoth
          TabOrder = 0
          WantReturns = False
          WordWrap = False
        end
      end
    end
  end
  object Panel3: TPanel
    Left = 0
    Top = 364
    Width = 597
    Height = 36
    Align = alBottom
    BevelOuter = bvNone
    Caption = ' '
    TabOrder = 2
    object Button1: TButton
      Left = 8
      Top = 4
      Width = 145
      Height = 25
      Caption = 'Open message in Netscape'
      TabOrder = 0
      OnClick = Button1Click
    end
  end
  object DropURLTarget1: TDropURLTarget
    DragTypes = [dtCopy, dtLink]
    GetDataOnEnter = True
    OnEnter = DropURLTarget1Enter
    OnDrop = DropURLTarget1Drop
    Target = Owner
    Left = 280
    Top = 232
  end
end
