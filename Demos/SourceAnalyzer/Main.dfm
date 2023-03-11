object FormMain: TFormMain
  Left = 571
  Top = 240
  Width = 573
  Height = 631
  Caption = 'Drop Source Analyzer'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object Splitter1: TSplitter
    Left = 0
    Top = 197
    Width = 565
    Height = 3
    Cursor = crVSplit
    Align = alTop
    AutoSnap = False
    MinSize = 1
    ResizeStyle = rsUpdate
  end
  object Panel2: TPanel
    Left = 0
    Top = 200
    Width = 565
    Height = 378
    Align = alClient
    BevelOuter = bvNone
    Caption = ' '
    Constraints.MinHeight = 100
    TabOrder = 0
    object MemoHexView: TRichEdit
      Left = 0
      Top = 0
      Width = 565
      Height = 378
      Align = alClient
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Courier New'
      Font.Style = []
      ParentColor = True
      ParentFont = False
      PlainText = True
      ReadOnly = True
      ScrollBars = ssBoth
      TabOrder = 0
      WantReturns = False
    end
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 578
    Width = 565
    Height = 19
    Panels = <>
    SimplePanel = True
  end
  object ListViewDataFormats: TListView
    Left = 0
    Top = 0
    Width = 565
    Height = 197
    Align = alTop
    Columns = <
      item
        Caption = 'ID'
      end
      item
        AutoSize = True
        Caption = 'Name'
        MinWidth = 50
      end
      item
        Caption = 'Aspect'
        Width = 75
      end
      item
        Caption = 'Medium'
        MinWidth = 50
        Width = 100
      end
      item
        Alignment = taRightJustify
        Caption = 'Size'
      end>
    ColumnClick = False
    Constraints.MinHeight = 100
    Constraints.MinWidth = 100
    HideSelection = False
    ReadOnly = True
    RowSelect = True
    SortType = stText
    TabOrder = 2
    ViewStyle = vsReport
    OnDeletion = ListViewDataFormatsDeletion
    OnSelectItem = ListViewDataFormatsSelectItem
  end
end
