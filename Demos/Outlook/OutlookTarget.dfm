object FormOutlookTarget: TFormOutlookTarget
  Left = 373
  Top = 288
  AutoScroll = False
  Caption = 'Outlook drop target demo'
  ClientHeight = 393
  ClientWidth = 672
  Color = clBtnFace
  Constraints.MinWidth = 200
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poDefaultPosOnly
  ShowHint = True
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object SplitterBrowser: TSplitter
    Left = 250
    Top = 0
    Width = 3
    Height = 374
    Cursor = crHSplit
    AutoSnap = False
    MinSize = 100
    ResizeStyle = rsUpdate
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 374
    Width = 672
    Height = 19
    Panels = <>
    SimplePanel = True
  end
  object ListViewBrowser: TListView
    Left = 0
    Top = 0
    Width = 250
    Height = 374
    Align = alLeft
    Columns = <
      item
        Caption = 'From'
        Width = 100
      end
      item
        AutoSize = True
        Caption = 'Subject'
      end>
    ColumnClick = False
    Constraints.MinWidth = 100
    HideSelection = False
    ReadOnly = True
    RowSelect = True
    TabOrder = 2
    ViewStyle = vsReport
    OnDeletion = ListViewBrowserDeletion
    OnSelectItem = ListViewBrowserSelectItem
  end
  object PanelMain: TPanel
    Left = 253
    Top = 0
    Width = 419
    Height = 374
    Align = alClient
    BevelOuter = bvNone
    Caption = ' '
    TabOrder = 1
    object SplitterAttachments: TSplitter
      Left = 0
      Top = 320
      Width = 419
      Height = 3
      Cursor = crVSplit
      Align = alBottom
      AutoSnap = False
      MinSize = 50
      ResizeStyle = rsUpdate
      Visible = False
    end
    object PanelFrom: TPanel
      Left = 0
      Top = 0
      Width = 419
      Height = 16
      Align = alTop
      AutoSize = True
      BevelOuter = bvNone
      Caption = ' '
      TabOrder = 0
      object Label1: TLabel
        Left = 4
        Top = 0
        Width = 26
        Height = 13
        Caption = 'From:'
      end
      object EditFrom: TEdit
        Left = 64
        Top = 0
        Width = 350
        Height = 16
        Cursor = crHandPoint
        Anchors = [akLeft, akTop, akRight]
        AutoSize = False
        BorderStyle = bsNone
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlue
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = [fsUnderline]
        ParentColor = True
        ParentFont = False
        ReadOnly = True
        TabOrder = 0
      end
    end
    object PanelTo: TPanel
      Left = 0
      Top = 16
      Width = 419
      Height = 16
      Align = alTop
      AutoSize = True
      BevelOuter = bvNone
      Caption = ' '
      TabOrder = 1
      object Label2: TLabel
        Left = 4
        Top = 0
        Width = 16
        Height = 13
        Caption = 'To:'
      end
      object ScrollBox1: TScrollBox
        Left = 64
        Top = 0
        Width = 350
        Height = 16
        Anchors = [akLeft, akTop, akRight]
        AutoSize = True
        BorderStyle = bsNone
        Constraints.MaxHeight = 60
        TabOrder = 0
        object ListViewTo: TListView
          Left = 0
          Top = 0
          Width = 350
          Height = 16
          Anchors = [akLeft, akTop, akRight]
          BorderStyle = bsNone
          Columns = <
            item
              AutoSize = True
              Caption = 'Name'
            end>
          ColumnClick = False
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clBlue
          Font.Height = -11
          Font.Name = 'MS Sans Serif'
          Font.Style = [fsUnderline]
          HotTrackStyles = [htHandPoint]
          ReadOnly = True
          ParentColor = True
          ParentFont = False
          ShowColumnHeaders = False
          TabOrder = 0
          ViewStyle = vsReport
          OnInfoTip = ListViewToInfoTip
        end
      end
    end
    object PanelSubject: TPanel
      Left = 0
      Top = 32
      Width = 419
      Height = 16
      Align = alTop
      AutoSize = True
      BevelOuter = bvNone
      Caption = ' '
      TabOrder = 2
      object Label3: TLabel
        Left = 4
        Top = 0
        Width = 39
        Height = 13
        Caption = 'Subject:'
      end
      object EditSubject: TEdit
        Left = 64
        Top = 0
        Width = 350
        Height = 16
        Anchors = [akLeft, akTop, akRight]
        AutoSize = False
        BorderStyle = bsNone
        ParentColor = True
        ReadOnly = True
        TabOrder = 0
      end
    end
    object MemoBody: TRichEdit
      Left = 0
      Top = 48
      Width = 419
      Height = 272
      Align = alClient
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Courier New'
      Font.Style = []
      Constraints.MinHeight = 50
      Lines.Strings = (
        'Drag one or more items from Outlook and drop them anywhere '
        'on this form.'
        ''
        'Note: Outlook Express is not supported.')
      ParentFont = False
      ReadOnly = True
      ScrollBars = ssBoth
      TabOrder = 3
      WantReturns = False
    end
    object ListViewAttachments: TListView
      Left = 0
      Top = 323
      Width = 419
      Height = 51
      Align = alBottom
      Columns = <
        item
          AutoSize = True
          Caption = 'Name'
        end
        item
          Caption = 'Type'
          Width = 100
        end
        item
          Caption = 'Size'
          Width = 75
        end>
      ColumnClick = False
      Constraints.MinHeight = 40
      IconOptions.AutoArrange = True
      LargeImages = ImageListBig
      ReadOnly = True
      RowSelect = True
      PopupMenu = PopupMenu1
      SmallImages = ImageListSmall
      TabOrder = 4
      Visible = False
      OnDblClick = ListViewAttachmentsDblClick
      OnDeletion = ListViewAttachmentsDeletion
      OnResize = ListViewAttachmentsResize
    end
  end
  object DataFormatAdapterOutlook: TDataFormatAdapter
    DragDropComponent = DropEmptyTarget1
    DataFormatName = 'TOutlookDataFormat'
    Left = 60
    Top = 125
  end
  object DropEmptyTarget1: TDropEmptyTarget
    DragTypes = [dtCopy, dtLink]
    OnDrop = DropTextTarget1Drop
    Target = Owner
    Left = 28
    Top = 125
  end
  object ImageListSmall: TImageList
    ShareImages = True
    Left = 64
    Top = 177
  end
  object ImageListBig: TImageList
    ShareImages = True
    Left = 28
    Top = 176
  end
  object PopupMenu1: TPopupMenu
    Left = 28
    Top = 220
    object MenuAttachmentOpen: TMenuItem
      Caption = 'Open'
      Default = True
      OnClick = ListViewAttachmentsDblClick
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object MenuAttachmentView: TMenuItem
      Caption = 'View'
      object MenuAttachmentViewLargeIcons: TMenuItem
        Caption = 'Large icons'
        Checked = True
        RadioItem = True
        OnClick = MenuAttachmentViewLargeIconsClick
      end
      object MenuAttachmentViewSmallIcons: TMenuItem
        Caption = 'Small icons'
        RadioItem = True
        OnClick = MenuAttachmentViewSmallIconsClick
      end
      object MenuAttachmentViewList: TMenuItem
        Caption = 'List'
        RadioItem = True
        OnClick = MenuAttachmentViewListClick
      end
      object MenuAttachmentViewDetails: TMenuItem
        Caption = 'Details'
        RadioItem = True
        OnClick = MenuAttachmentViewDetailsClick
      end
    end
  end
end
