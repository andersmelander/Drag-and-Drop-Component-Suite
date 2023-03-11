OutlookDemo

This application demonstrates how to receive mail and other items dropped from
Outlook.

The application extends a TDropEmptyTarget component with support for the
TOutlookDataFormat class using a TDataFormatAdapter component.

The TOutlookDataFormat class uses the TStorageDataFormat class to extract an
IStorage (structured storage) object from the data provided by outlook. It then
"converts" the IStorage object to an IMessage MAPI object. Finally the
application retrieves the content and properties of the dropped Outlook message
from the IMessage object.

Note: TOutlookDataFormat only works with Outlook. Outlook Express is not
supported.
If you need Outlook Express support, try the ComboTargetDemo application.
