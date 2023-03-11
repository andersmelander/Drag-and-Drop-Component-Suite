object DataModuleDropHandler: TDataModuleDropHandler
  OldCreateOrder = False
  Left = 269
  Top = 106
  Height = 150
  Width = 215
  object DropHandler1: TDropHandler
    Dragtypes = [dtCopy]
    GetDataOnEnter = False
    OnDrop = DropHandler1Drop
    ShowImage = True
    OptimizedMove = True
    AllowAsyncTransfer = False
    Left = 36
    Top = 16
  end
end
