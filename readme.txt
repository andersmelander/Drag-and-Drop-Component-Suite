
Drag and Drop Component Suite Version 4.1
Release candidate 3, released 2-may-2003
© 1997-2003 Angus Johnson & Anders Melander
http://www.melander.dk/delphi/dragdrop/

-------------------------------------------
Table of Contents:
-------------------------------------------
1. Supported platforms
2. Installation
3. Getting started
4. Known problems
5. Support and feedback
6. Bug reports
7. Upgrades and bug fixes
8. Missing in this release
9. New in version 4.x
10. TODO
11. Licence, Copyright and Disclaimer
12. Release history


-------------------------------------------
1. Supported platforms:
-------------------------------------------
This section has been moved to the help file.


-------------------------------------------
2. Installation:
-------------------------------------------
This section has been moved to the help file.


-------------------------------------------
3. Getting started:
-------------------------------------------
This section has been moved to the help file.


-------------------------------------------
4. Known problems:
-------------------------------------------
* When the demo applications are compiled with Delphi 7, some of them
  will probably emit a lot of "Unsafe type", "Unsafe code" etc. warnings.
  This doesn't mean that there's anything wrong with the demos. It just
  means that Borland would like you to know that they are moving to .NET
  and would like you to do the same (so you can buy the next version of
  Delphi).
  You can turn the warnings of in the project options.

* Some of the demos doesn't work with C++Builder 4 yet.
  Specifically the DetailedDemo and ExtractDemo applications can't compile.

* The Shell Extension components does not support C++Builder 4.
  For some strange reason the components causes a link error.
  I have a single report of problems with C++ Builder 5 too, but I'm unable
  to verify as I haven't got C++ Builder installed anymore.

* There appear to be sporadic problems compiling with C++Builder 5.
  Several user have reported that they occasionally get one or more of the
  following compiler errors:

    [C++ Error] DragDropFile.hpp(178): E2450 Undefined structure
      '_FILEDESCRIPTORW'
    [C++ Error] DropSource.hpp(135): E2076 Overloadable operator expected

  I have not been able to reproduce these errors, but I believe the following
  work around will fix the problem: In the project options of *all* projects
  which uses these components, add the following conditional define:

    NO_WIN32_LEAN_AND_MEAN

  The define *must* be made in the project options. It is not sufficient to
  #define it in the source.

* Delphi's and C++Builder's HWND and THandle types are not compatible.
  For this reason it might be nescessary to cast C++Builder's HWND values to
  Delphi's THandle type when a HWND is passed to a function. E.g.:

    if (DragDetectPlus(reinterpret_cast<THandle>(Handle), Point(X, Y))) {
      ...
    }

* Virtual File Stream formats can only be pasted from the clipboard with live
  data (i.e. FlushClipboard/OleFlushClipboard hasn't been called on the data
  source). This problem affects TFileContentsStreamOnDemandClipboardFormat and
  the VirtualFileStream demo.
  This is believed to be a bug in the Windows clipboard and a work around hasn't
  been found yet.

* When TDropFileTarget.GetDataOnEnter is set to True, the component doesn't work
  with WinZip.
  Although the file names are received correctly by TDropFileTarget, WinZip
  doesn't extract the files and the files thus can't be copied/moved.
  This is caused by a quirk in WinZip; Apparently WinZip doesn't like
  IDataObject.GetData to be called before IDropTarget.Drop is called.

-------------------------------------------
5. Support and feedback:
-------------------------------------------
This section has been moved to the help file.


-------------------------------------------
6. Bug reports:
-------------------------------------------
This section has been moved to the help file.


-------------------------------------------
7. Upgrades and bug fixes:
-------------------------------------------
This section has been moved to the help file.


-------------------------------------------
8. Missing in this release:
-------------------------------------------
* Online help is not complete.
  I'm still working on the online help so many topics are either missing,
  incomplete or out of date (e.g. those that were copied from the v3.7
  help).


-------------------------------------------
9. New in version 4.x:
-------------------------------------------
* Completely redesigned and rewritten.
  Previous versions of the Drag and Drop Component Suite used a very monolithic
  design and flat class hierachy which made it a bit cumbersome to extend the
  existing components or implement new ones.

  Version 4 is a complete rewrite and redesign, but still maintains
  compatibility with previous versions. The new V4 design basically separates
  the library into three layers:

    1) Clipboard format I/O.
    2) Data format conversion and storage.
    3) COM Drag/Drop implementation and VCL component interface.

  The clipboard format layer is responsible for reading and writing data in
  different formats to and from an IDataObject interface. For each different
  clipboard format version 4 implements a specialized class which knows exactly
  how to interpret the clipboard format. For example the CF_TEXT (plain text)
  clipboard format is handled by the TTextClipboardFormat class and the CF_FILE
  (file names) clipboard format is handled by the TFileClipboardFormat class.

  The data format layer is primarily used to render the different clipboard
  formats to and from native Delphi data types. For example the TTextDataFormat
  class represents all text based clipboard formats (e.g. TTextClipboardFormat)
  as a string while the TFileDataFormat class represents a list of file names
  (e.g. TFileClipboardFormat) as a string list.

  The conversion between different data- and clipboard formats is handled by the
  same Assign/AssignTo mechanism as the VCLs TPersistent employes. This makes it
  possible to extend existing data formats with support for new clipboard
  formats without modification to the existing classes.

  The drag/drop component layer has several tasks; It implements the actual COM
  drag/drop functionality (i.e. it implements the IDropSource, IDropTarget and
  IDataObject interfaces (along with several other related interfaces)), it
  surfaces the data provided by the data format layer as component properties
  and it handles the interaction between the whole drag/drop framework and the
  users code.
  The suite provides a multitude of different components. Most are specialized
  for different drag/drop tasks (e.g. the TDropFileTarget and TDropFilesSource
  components for drag/drop of files), but some are either more generic, handling
  multiple unrelated formats, or simply helper components which are used to
  extend the existing components or build new ones.

* Support for Delphi 6 and Delphi 7.
  Version 4.0 was primarily developed on Delphi 6 and then ported back to
  previous versions of Delphi and C++Builder.

* Support for Windows 2000 inter application drag images.
  On Windows platforms which supports it, drag images are now displayed when
  dragging between applications. Currently only Windows 2000 supports this
  feature. On platforms which doesn't support the feature, drag images are only
  displayed whithin the source application.

* Support for Windows 2000 asynchronous data transfers.
  Asynchronous data tranfers allows the drop source and targets to perform slow
  transfers, or to transfer large amounts of data, without blocking the user
  interface while the data is being transfered.
  For platforms prior to Windows 2000, or when the drop target doesn't support
  asynchronous data transfers, the drop source components on their own can 
  provide similar (but more limited) asynchronous data transfer capabilities.

* Support for optimized and non-optimized move.
  When performing drag-move operations, it is now possible to specify if the
  target (optimized move) or the source (non-optimized move) is responsible for
  deleting the source files.

* Support for delete-on-paste.
  When data is cut to the clipboard, it is now possible to defer the deletion of
  the source data until the target actually pastes the data. The source is
  notified by an event when the target pastes the data.

* Extended clipboard support.
  All formats and components (both source and target) now support clipboard
  operations (copy/cut/paste) and the VCL clipboard object.

* Support for shell drop handlers.
  The new TDropHandler component can be used to write drop handler shell
  extensions. A drop handler is a shell extension which is executed when a user
  drags and drops one or more files on a file associated wth your application.

* Support for shell drag drop handlers.
  The new TDragDropHandler component can be used to write drag drop handler
  shell extensions. A drag drop handler is a shell extension which can extend
  the popup menu which is displayed when a user drag and drops files with the
  right mouse button.

* Support for shell context menu handlers.
  The new TDropContextMenu component can be used to write context menu handler
  shell extensions. A context menu handler is a shell extension which can extend
  the popup menu which is displayed when a user right-clicks a file in the
  shell.

* Drop sources can receive data from drop targets.
  It is now possible for drop targets to write data back to the drop source.
  This is used to support optimized-move, delete-on-paste and inter application
  drag images.

* Automatic re-registration of targets when the target window handle is
  recreated.
  In previous versions, target controls would loose their ability to accept
  drops when their window handles were recreated by the VCL (e.g. when changing
  the border style or docking a form). This is no longer a problem.

* Support for run-time definition of custom data formats.
  You can now add support for new clipboard formats without custom components.

* Support for design-time extension of existing source and target components.
  Using the new TDataFormatAdapter component it is now possible to mix and match
  data formats and source and target components at design time. E.g. the
  TDropFileTarget component can be extended with URL support.

* It is now possible to completely customize the target auto-scroll feature.
  Auto scroling can now be completely customized via the OnDragEnter,
  OnDragOver, OnGetDropEffect and OnScroll events and the public NoScrollZone
  and published AutoScroll properties.

* Multiple target controls per drop target component.
  In previous versions you had to use one drop target component per target
  control. With version 4, each drop target component can handle any number of
  target controls.

* It is now possible to specify the target control at design time.
  A published Target property has been added to the drop target components.

* Includes 20 components:
  - TDropFileSource and TDropFileTarget
    Used for drag and drop of files. Supports recycle bin and PIDLs.

  - TDropTextSource and TDropTextTarget
    Used for drag and drop of text.

  - TDropBMPSource and TDropBMPTarget
    Used for drag and drop of bitmaps.

  - TDropPIDLSource and TDropPIDLTarget
    Used for drag and drop of PIDLs in native format.

  - TDropURLSource and TDropURLTarget
    Used for drag and drop of internet shortcuts.

  - TDropDummyTarget
    Used to provide drag/drop cursor feedback for controls which aren't
    registered as drop targets.

  - TDropComboTarget (new)
    Swiss-army-knife target. Accepts text, files, bitmaps, meta files, URLs and
    file contents.

  - TDropMetaFileTarget (new)
    Target which can accept meta files and enhanced meta files.

  - TDropImageTarget (new)
    Target which can accept bitmaps, DIBs, meta files and enhanced meta files.

  - TDragDropHandler (new)
    Used to implement Drag Drop Handler shell extensions.

  - TDropHandler (new)
    Used to implement Shell Drop Handler shell extensions.

  - TDragDropContext (new)
    Used to implement Shell Context Menu Handler shell extensions.

  - TDataFormatAdapter (new)
    Extends the standard source and target components with support for extra
    data formats. An alternative to TDropComboTarget.

  - TDropEmptySource and TDropEmptyTarget (new)
    Target and source components which doesn't support any formats, but can be
    extended with TDataFormatAdapter components.

* Supports 27 standard clipboard formats:
  Text formats:
  - CF_TEXT (plain text)
  - CF_UNICODETEXT (Unicode text)
  - CF_OEMTEXT (Text in the OEM characterset)
  - CF_LOCALE (Locale specification)
  - 'Rich Text Format' (RTF text)
  - 'CSV' (Tabular spreadsheet text)
  File formats:
  - CF_HDROP (list of file names)
  - CF_FILEGROUPDESCRIPTOR, CF_FILEGROUPDESCRIPTORW and CF_FILECONTENTS (list of
    files and their attributes and content).
  - 'Shell IDList Array' (PIDLs)
  - 'FileName' and 'FileNameW' (single filename, used for 16 bit compatibility).
  - 'FileNameMap' and 'FileNameMapW' (used to rename files, usually when
    dragging from the recycle bin)
  Image formats:
  - CF_BITMAP (Windows bitmap)
  - CF_DIB (Device Independant Bitmap)
  - CF_METAFILEPICT (Windows MetaFile)
  - CF_ENHMETAFILE (Enhanced Metafile)
  - CF_PALETTE (Bitmap palette)
  Internet formats:
  - 'UniformResourceLocator' and 'UniformResourceLocatorW' (Internet shortcut)
  - 'Netscape Bookmark' (Netscape bookmark/URL)
  - 'Netscape Image Format' (Netscape image/URL)
  - '+//ISBN 1-887687-00-9::versit::PDI//vCard' (V-Card)
  - 'HTML Format' (HTML text)
  - 'Internet Message (rfc822/rfc1522)' (E-mail message in RFC822 format)
  Misc. formats:
  - CF_PREFERREDDROPEFFECT and CF_PASTESUCCEEDED (mostly used by clipboard)
  - CF_PERFORMEDDROPEFFECT and CF_LOGICALPERFORMEDDROPEFFECT (mostly used for
    optimized-move)
  - 'InShellDragLoop' (used by Windows shell)
  - 'TargetCLSID' (Mostly used when dragging to recycle-bin)

* New source events:
  - OnGetData: Fired when the target requests data.
  - OnSetData: Fired when the target writes data back to the source.
  - OnPaste: Fired when the target pastes data which the source has placed on
    the clipboard.
  - OnAfterDrop: Fired after the drag/drop operation has completed.

* New target events:
  - OnScroll: Fires when the target component is about to perform auto-scroll on
    the target control.
  - OnAcceptFormat: Fires when the target component needs to determine if it
    will accept a given data format. Only surfaced in the TDropComboTarget
    component.

* 20 new demo applications, 31 in total.


-------------------------------------------
10. TODO (may or may not be implemented):
-------------------------------------------
Help file needs a lot of work.


-------------------------------------------
11. Licence, Copyright and Disclaimer:
-------------------------------------------
This section has been moved to the help file.


-------------------------------------------
12. Version 4 release history:
-------------------------------------------
02-may-2003
  * Updated AsyncSource.BCB demo.
    I haven't tested this update since I haven't got BCB installed anymore.

  * Released for test as v4.1 RC3.

25-apr-2003
  * Renamed TFeedbackDataFormat.PasteSucceded to PasteSucceeded.
    Oups. Thanks to Chris Ueberall for pointing this out.

23-apr-2003
  * Made FileGroupDescriptor and FileContents support in TDropTextSource
    (TTextDataFormat actually) conditional.
    Disable the DROPSOURCE_TEXTSCRAP definition in DragDropText.pas to remove
    FileGroupDescriptor and FileContents support from TDropTextSource.
    This change was made to support custom handling of these formats with drag
    of text. In a future version the support will be completely moved from
    TTextDataFormat to another support class and TDataFormatAdapter will become
    the preferred method of adding FileGroupDescriptor and FileContents support
    to TDropTextSource.

  * Modified TFileClipboardFormat (CF_HDROP) to transfer data in widestring
    format on NT and in ansistring format otherwise.

  * Added work around for Nero Express CD burner bug.
    Nero Express incorrectly specifies the desired transfer medium as -1 and
    then chokes if we give it anything other than a HGlobal medium.

  * Modified TDropTextSource so the Locale property isn't updated automatically
    when the Text property is modified at design time.

  * Added work around for Delphi 7 TComboBox bug to PathComboBox.pas used in the
    PIDLDemo application.

14-apr-2003
  * Fixed problem with Delphi 7 package.

  * Added support for Objects property to
    TFileDescriptorToFilenameStrings class.

04-apr-2003
  * Released for test as v4.1 RC2.

19-mar-2003
  * Added work around for Outlook 2000 attachment copy/paste bug.

11-feb-2003
  * Port to Delphi 7.
    Disabled all "Unsafe" .NET warnings in DragDrop.inc.

  * Released for test as v4.1 RC1.

20-jan-2003
  * Fixed reference count problems in OutlookDemo application.

12-jan-2003
  * Redesigned and reimplemented asynchronous data transfer in the source
    components.

06-jan-2003
  * Renamed TInterfaceList to TNamedInterfaceList to avoid conflict with
    the VCL Classes.TInterfaceList.

10-dec-2003
  * Implemented TOutlookDataFormat data format.

  * Added OutlookDropTarget demo application.

01-dec-2002
  * Resuming work after a "short" vacation...

  * Recovered most of the work I lost in the first May 2002 disk crash.
    Will try to recover the rest from updates I have mailed to users.

27-may-2002
  * Another disk crash.
    Burnt out. Going on vacation...

13-may-2002
  * Disk crash.
    Lost all work done after feb 2002. Recovering from March 2002 backup.

04-feb-2002
  * Renamed TPasteSuccededClipboardFormat class to
    TPasteSucceededClipboardFormat.
    Oups - I wish Delphi had a spell checker... 

28-jan-2001
  * Made target auto-scroll cursor sample rate customizable
    (See DragDropScrollVelocitySample).

26-jan-2002
  * Fixed thread termination bug in DetailedDemo.
    Thanks to "anonymous" for spotting this problem.

  * Misc changes for C++ Builder 6 compatibility.
    Thanks to "Deep throat" for help with this.

21-jan-2002
  * Added 3 more C++Builder demos from Jonathan Arnold.
    Some of the demos doesn't work with C++Builder 4 yet.

  * Fixed mouse velocity problem with auto-scroll and TDropDummy component.

  * Added TCustomDropTarget.CanPasteFromClipboard.
    Based on idea by Pieter Zijlstra.

  * Released for test as v4.1 FT6.

17-jan-2002
  * New help file included.
    Completely rewritten from scratch, but not yet completed..

  * Installer now includes an uninstaller.

  * Fixed bug in unregistration of target controls that were being destroyed.
    Thanks to Thomas Reimann for bringing this problem to my attention.

  * TDropFileTarget.OptimizedMove now defaults to False for consistency.

  * When the OnDrop target event is being fired as a result of a call to
    PasteFromClipboard, the Target property will now be set to nil to indicate
    that the OnDrop was the result of a clipboard operation.

13-jan-2002
  * Fixed problems in drop target positioning of drag images.
    Thanks to Pieter Zijlstra for reporting this problem.

  * Fixed default size of auto-scroll zone.
    Thanks to Pieter Zijlstra for reporting this problem.

  * Fixed auto-scroll with drag images.
    Thanks to Pieter Zijlstra for reporting this problem.

  * Fixed problems in TDropContextMenu with the "Directory\Background" shell
    object.

5-jan-2002
  * Improved auto scroll.
    Auto scroll is now only performed if the target's scroll bars are actually
    visible. This fixes some problem with controls that couldn't handle
    WM_SCROLL messages with hidden scroll bars.
    Thanks to Praful Kapadia for spotting this problem.

  * Added OwnsObjects property to TClipboardFormats class.

  * Fixed bug in KeysToShiftStatePlus.
    Alt key was handled incorrectly.
    
  * TScrolDirection and TScrolDirections types has been deprecated and replaced
    with the TScrollDirection and TScrollDirections types.
    This might affect applications that have implemented the OnScroll drop
    target event.

27-dec-2001
  * Added design time warning message for use of TCustomRichEdit as a target
    with the TCustomDropTarget.AutoRegister property set to True.
    If AutoRegister is set to true and the target is a rich edit control, the
    text of the rich edit control can become invisible if the control is
    scrolled or otherwise modified. The only known work around at this time is
    to set AutoRegister to False.

  * Reverted OnGetDropEffect event to pre V4 behaviour.
    Prior to version 4 the Effect parameter of the OnGetDropEffect event
    contained a mask of drop effects supported by the drop source. In V4 this
    was changed so that the Effect parameter contained the default drop effect.
    Since this change made it impossible for the target to determine what drop
    effects the source supported, the event has been changed back to how it was
    prior to V4.
    Thanks to Praful Kapadia for reporting this problem.
    Since it is still desireable to pass the default drop effect on to the
    OnGetDropEffect event, this event will either be redeclared in a future
    version or a new event will be added. This will likely happen in version 5.

  * Fixed a bunch of memory leaks.
    Thanks to Hideo Koiso for reporting these leaks.

  * Installation kit modified to store install location and a bunch of other
    install related info in the registry.
    Since no uninstaller is included, this info will be left in the registry if
    the components are uninstalled/deleted. Shouldn't be a big deal.
    The installer no longer scans the disk for previous versions of the
    components, but instead looks in the registry for a previous version.

22-dec-2001
  * Added SimpleContextMenuHandlerShellExt demo.
    A simple demo that demonstrates use of the TDropContextMenu component but
    doesn't use as many advanced features as the ContextMenuHandlerShellExt
    demo.

  * Fixed problem with TDropContextMenu and image lists.
    I'm still not quite satisfied with all popup menu/bitmap/image list
    combinations, but it'll have to do for now.

  * Added a few more C++Builder demos ported by Jonathan Arnold.

16-dec-2001
  * Ported to C++Builder 4.

  * Released for test as v4.1 FT5.

12-dec-2001
  * Fixed C++Builder name clash between TDropComboTarget.GetMetaFile and the
    GetMetaFile #define in wingdi.h
    
1-dec-2001
  * The IAsyncOperation interface is now also declared as IAsyncOperation2 and
    all references to IAsyncOperation has been replaced with IAsyncOperation2.
    This was done to work around a bug in C++Builder.
    Thanks to Jonathan Arnold for all his help with getting the components to
    work with C++Builder. Without Jonathan's help version 4.1 would prabably
    have shipped witout C++Builder support and certainly without any C++
    Builder demos.

  * Demo applications for C++Builder.
    The C++Builder demos were contributed by Jonathan Arnold.

27-nov-2001
  * TCustomDropTarget.Droptypes property renamed to DropTypes (notice the case).
    Thanks to Krystian Brazulewicz for spotting this.

24-nov-2001
  * The GetURLFromString function in the DragDropInternet unit has been made
    public due to user request.

21-nov-2001
  * Modified MakeHTML function to comply with Microsoft's description of the
    CF_HTML clipboard format.

  * Added MakeTextFromHTML function to convert CF_HTML data to plain HTML.
    Provides the reverse functionality of MakeHTML.

  * Added HTML support to TTextDataFormat class and TDropTextSource and
    TDropTextTarget components. 

  * Fixed C++Builder 5 problem with IAsyncOperation.

  * Released for test as v4.1 FT4.

10-nov-2001
  * Added NetscapeDemo demo application.
    Demonstrates how to receive messages dropped from Netscape.
    This demo was sponsored by ThoughtShare Communications Inc.

  * Released for test as v4.1 FT3.

23-oct-2001
  * Conversion priority of TURLDataFormat has been changed to give the
    File Group Descritor formats priority over the Internet Shortcut format.
    This resolves a problem where dropping an URL on the desktop would cause the
    desktop to assume that an Active Desktop item was to be created instead of
    an Internet Shortcut.
    Thanks to Allen Martin for reporting this problem.
    By luck this modification also happens to work around a bug in Mozilla and
    Netscape 6; Mozilla incorrectly supplies the UniformResourceLocator
    clipboard format in unicode format instead of ANSI format.
    Thanks to Florian Kusche for reporting this problem.

  * Added support for TFileGroupDescritorWClipboardFormat to TURLDataFormat.

  * Added declaration of FD_PROGRESSUI to DragDropFormats.

  * Added TURLWClipboardFormat which implements the "UniformResourceLocatorW"
    (a.k.a. CFSTR_INETURLW) clipboard format. Basically a Unicode version of
    CFSTR_SHELLURL/CFSTR_INETURL.
    The TURLWClipboardFormat class isn't used anywhere yet but will probably be
    supported by TURLDataFormat (and thus TDropURLTarget/TDropURLSource) in a
    later release.

  * Added experimental Shell Drag Image support.
    This relies on undodumented shell32.dll functions and probably won't be
    fully support before v4.2 (if ever). See InitShellDragImage in
    DropSource.pas.
    Thanks to Jim Kueneman for bringning these functions to my attention.

13-oct-2001
  * TCustomDropSource.Destroy and TCustomDropMultiSource.Destroy changed to call
    FlushClipboard instead of EmptyClipboard.
    This means that clipboard contents will be preserved when the source
    application/component is terminated.

  * Added clipboard support to VirtualFileStream demo.

  * Modified VirtualFileStream demo to work around clipboard quirk with IStream
    medium.

  * Modified TCustomSimpleClipboardFormat to disable TYMED_ISTORAGE support by
    default.
    At present TYMED_ISTORAGE is only supported for drop targets and enabling it
    by default in TCustomSimpleClipboardFormat.Create caused a lot of clipboard
    operations (e.g. copy/paste of text) to fail.
    Thanks to Michael J Marshall for bringing this problem to my attention.

  * Modified TCustomSimpleClipboardFormat to read from the the TYMED_ISTREAM
    medium in small (1Mb) chunks and via a global memory buffer.
    This has resultet in a huge performance gain (several orders of magnitude)
    when transferring large amounts of data via the TYMED_ISTREAM medium.

3-oct-2001
  * Fixed bug in TCustomDropSource.SetImageIndex.
    Thanks to Maxim Abramovich for spotting this.

  * Added missing default property values to TCustomDropSource.
    Thanks to Maxim Abramovich for spotting this.

  * DragDrop.pas and DragDropContext.pas updated for Delphi 4.

  * Reimplemented utility to convert DFM form files from Delphi 5/6 test format
    to Delphi 4/5 binary format.

  * Improved unregistration of Shell Extensions.
    Shell extension now completely (and safely) remove their registry entries
    when unregistered.

  * Deprecated support for C++Builder 3.

  * Released for test as v4.1 FT2.

25-sep-2001
  * Rewritten ContextMenuHandlerShellExt demo.
    The demo is now actually a quite useful utility which can be used to
    register and unregister ActiveX controls, COM servers and type libraries. It
    includes the same functionality as Borland's TRegSvr utility.

20-sep-2001
  * Added support for cascading menus, ownerdraw and menu bitmaps to
    TDropContextMenu component.

  * Modified TFileContentsStreamOnDemandClipboardFormat to handle invalid
    parameter value (FormatEtcIn.lindex) when data is copied to clipboard.
    This works around an apparent bug in the Windows clipboard. Thanks to Steve
    Moss for reporting this problem.

  * Modified TEnumFormatEtc class to not enumerate empty clipboard formats.
    Thanks to Steve Moss for this improvement.

1-sep-2001
  * Introduced TCustomDropTarget.AutoRegister property.
    The AutoRegister property is used to control if drop target controls should
    be automatically unregistered and reregistered when their window handle is
    recreated by the VCL. If AutoRegister is True, which is the default, then
    automatic reregistration will be performed.
    This property was introduced because the hidden child control, which is used
    to monitor the drop target control's window handle, can have unwanted side
    effects on the drop target control (e.g. TToolBar).

  * Deprecated support for Delphi 3.

22-jun-2001
  * Redesigned TTextDataFormat to handle RTF, Unicode, CSV and OEM text without
    conversion.
    Moved TTextDataFormat class to DragDropText unit.
    Added support for TLocaleClipboardFormat.

  * Surfaced new text formats as properties in TDropTextSource and
    TDropTextTarget.
    Previous versions of the Text source and target components represented all
    supported text formats via the Text property. In order to enable users to
    handle the different text formats independantly, the text source and target
    components now has individual properties for ANSI, OEM, Unicode and RTF text
    formats.
    The text target component can automatically synthesize some of the formats
    from the others (e.g. OEM text from ANSI text), but applications which
    previously relied on all formats being represented by the Text property will
    have to be modified to handle the new properties.

  * Added work around for problem where TToolBar as a drop target would display
    the invisible target proxy window.

  * Fixed wide string bug in WriteFilesToZeroList.
    Thanks to Werner Lehmann for spotting this.

15-jun-2001
  * Added work-around for Outlook Express IDataObject.QueryGetData quirk.

3-jun-2001
  * Ported to C++Builder 4 and 5.

  * Added missing DragDropDesign.pas unit to design time packages.

  * First attempt at C++Builder 3 port.... failed.

  * Improved handling of oversized File Group Descriptor data.

  * Added support for IStorage medium to TFileContentsStreamClipboardFormat.
    This allows the TDropComboTarget component to accept messages dropped from
    Microsoft Outlook.
    This work was sponsored by ThoughtShare Communications Inc.

23-may-2001
  * Ported to Delphi 4.

  * First attempt at C++Builder 5 port.... failed.

18-may-2001
  * Released as version 4.0.
    Note: Version 4.0 was released exclusively on the Delphi 6 Companion CD.

  * ContextMenuDemo and DropHandlerDemo application has been partially
    rewritten and renamed.
    ContextMenuDemo is now named ContextMenuHandlerShellExt.
    DropHandlerDemo is now named DropHandlerShellExt.

  * TDropContextMenu component has been rewitten.
    The TDropContextMenu now implements a context menu handler shell extension.
    In previous releases it implemented a drag drop handler shell extension.

  * The DragDropHandler.pas unit which implements the TDropHandler component
    has been renamed to DropHandler.pas.

  * Added new TDragDropHandler component.
    The new component, which lives in the DragDropHandler unit, is used to
    implement drag drop handler shell extensions.

  * Added DragDropHandlerShellExt demo application.

  * Removed misc incomplete demos from kit.

  * Fixed minor problem in VirtualFileStream demo which caused drops from
    the VirtualFile demo not to transfer content correctly.

11-may-2001
  * Converted all demo forms to text DFM format.
    This has been nescessary to maintain compatibility between all supported
    versions of Delphi.

  * Fixed a bug in GetPIDLsFromFilenames which caused drag-link of files (dtLink
    with TDropFileSource) not to work.

  * Added readme.txt files to some demo applications.

  * Added missing tlb and C++Builder files to install kit.

  * Released as FT4.

6-may-2001
  * Added missing dfm files to install kit.

  * Tested with Delphi 5.
    Fixed Delphi 5 compatibility error in main.dfm of DragDropDemo.

  * Removed misc compiler warnings.

  * The AsyncTransferTarget and OleObjectDemo demos were incomplete and has been removed
    from the kit for the V4.0 release.
    The demos will be included in a future release.

  * Released as FT3.

3-may-2001
  * Added missing dpr and bpg files to install kit.

  * Updated readme.txt with regard to lack of C++Builder demos.

  * Released as FT2.

29-apr-2001
  * Cleaned up for release.

  * Released as FT1.


23-feb-2001
  * Modified TCustomDropTarget.FindTarget to handle overlapping targets (e.g.
    different targets at the same position but on different pages of a page
    control or notebook).
    Thanks to Roger Moe for spotting this problem.

13-feb-2001
  * Renamed AsyncTransfer2 demo to AsyncTransferSource.

  * Added AsyncTransferTarget demo.

  * Replaced TChart in AsyncTransfer2 demo with homegrown pie-chart-thing.

  * Modified all IStream based target formats to support incremental transfer.

  * URW533 problem has finally been fixed.
    The cause of the problem, which is a bug in Delphi, was found by Stefan
    Hoffmeister.

  * Fixed free notification for TDropContextmenu and TDataFormatAdapter.

27-dec-2000
  * Moved TVirtualFileStreamDataFormat and TFileContentsStreamOnDemandClipboardFormat
    classes from VirtualFileStream demo to DragDropFormats unit.

  * Added TClipboardFormat.DataFormat and TClipboardFormats.DataFormat property.

  * Added TDropEmptySource and TDropEmptyTarget components.
    These are basically do-nothing components for use with TDataFormatAdapter.

  * Rewritten AsyncTransfer2 demo.
    The demo now uses TDropEmptySource, TDataFormatAdapter and
    TVirtualFileStreamDataFormat to transfer 10Mb of data with progress
    feedback.

  * Rewritten VirtualFileStream demo.
    The demo now uses TDropEmptySource, TDropEmptyTarget, TDataFormatAdapter and
    TVirtualFileStreamDataFormat.

  * Fixed memory leak in TVirtualFileStreamDataFormat.
    This leak only affected the old VirtualFileStream demo.

  * Added support for full File Descriptor attribute set to
    TVirtualFileStreamDataFormat.
    It is now possible to specify file attributes such as file size and last
    modified time in addition to the filename. I plan to add similar features to
    the other classes which uses FileDescriptors (e.g. TDropFileSource and
    TDropFileTarget).

21-dec-2000
  * Ported to Delphi 4.

  * Added workaround for design bug in either Explorer or the clipboard.
    Explorer and the clipboard's requirements to the cursor position of an
    IStream object are incompatible. Explorer requires the cursor to be at the
    beginning of stream and the clipboard requires the cursor to be at the end
    of stream.

15-dec-2000
  * Fixed URW533 problem.
    I'll leave the description of the workaround in here for now in case the
    problem resurfaces.

11-dec-2000
  * Fixed bug in filename to PIDL conversion (GetPIDLsFromFilenames) which
    affected TDropFileTarget.
    Thanks to Poul Halgaard Jørgensen for reporting this.

4-dec-2000
  * Added THTMLDataFormat.

  * Fixed a a few small bugs which affected clipboard operations.

  * Added {$ALIGN ON} to dragdrop.inc.
    Apparently COM drag/drop requires some structures to be word alligned.
    This change fixes problems where some of the demos would suddenly stop
    working.

  * The URW533 problem has resurfaced.
    See the "Known problems" section below.

13-nov-2000
  * TCopyPasteDataFormat has been renamed to TFeedbackDataFormat.

  * Added support for the Windows 2000 "TargetCLSID" format with the
    TTargetCLSIDClipboardFormat class and the TCustomDropSource.TargetCLSID
    property.

  * Added support for the "Logical Performed DropEffect" format with the
    TLogicalPerformedDropEffectClipboardFormat class. The class is used
    internally by TCustomDropSource.

30-oct-2000
  * Added ContextMenu demo and TDropContextMenu component.
    Demonstrates how to customize the context menu which is displayed when a
    file is dragged with the right mouse button and dropped in the shell.

  * Added TCustomDataFormat.GetData.
    With the introduction of the GetData method, Data Format classes can now be
    used stand-alone to extract data from an IDataObject.

20-oct-2000
  * Added VirtualFileStream demo.
    Demonstrates how to use the "File Contents" and "File Group Descritor"
    clipboard formats to drag and drop virtual files (files which doesn't exist
    physically) and transfer the data on-demand via a stream.

14-oct-2000
  * Added special drop target registration of TCustomRichEdit controls.
    TCustomRichEdit needs special attention because it implements its own drop
    target handling which prevents it to work with these components.
    TCustomDropTarget now disables a rich edit control's built in drag/drop
    handling when the control is registered as a drop target.

  * Added work around for Windows bug where IDropTarget.DragOver is called
    regardless that the drop has been rejected in IDropTarget.DragEnter.

12-oct-2000
  * Fixed bug that caused docking to interfere with drop targets.
    Thanks to G. Bradley MacDonald for bringing the problem to my attention.

30-sep-2000
  * The DataFormats property has been made public in the TCustomDropMultiTarget
    class.

  * Added VirtualFile demo.
    Demonstrates how to use the TFileContentsClipboardFormat and
    TFileGroupDescritorClipboardFormat formats to drag and drop a virtual file
    (a file which doesn't exist physically).

28-sep-2000
  * Improved drop source detection of optimized move.
    When an optimized move is performed by a drop target, the drop source's
    Execute method will now return drDropMove. Previously drCancel was returned.
    The OnAfterDrop event must still be used to determine if a move operation
    were optimized or not.

  * Modified TCustomDropTarget.GetPreferredDropEffect to get data from the
    current IDataObject instead of from the VCL global clipboard.

18-sep-2000
  * Fixed bug in DropComboTarget caused by the 17-sep-2000 TStreams
    modification.

17-sep-2000
  * Added AsyncTransfer2 demo to demonstrate use of TDropSourceThread.

  * Renamed TStreams class to TStreamList.

29-aug-2000
  * Added TDropSourceThread.
    TDropSourceThread is an alternative to Windows 2000 asynchronous data
    transfers but also works on other platforms than Windows 2000.
    TDropSourceThread is based on code contributed by E. J. Molendijk.

24-aug-2000
  * Added support for Windows 2000 asynchronous data transfers.
    Added IAsyncOperation implementation to TCustomDropSource.
    Added TCustomDropSource.AllowAsyncTransfer and AsyncTransfer properties.

5-aug-2000
  * Added work around for URW533 compiler bug.

  * Fixed D4 and D5 packages and updated a few demos.
    Obsolete DropMultiTarget were still referenced  a few places.

  * Documented work around for C++Builder 5 compiler error.
    See the Known Problems section later in this document for more information.

2-aug-2000
  * The package files provided in the kit is now design-time only packages.
    In previous versions, the packages could be used both at design- and
    run-time. The change was nescessary because the package now contains
    design-time code.

  * Added possible work around for suspected C++Builder bug.
    The bug manifests itself as a "Overloadable operator expected" compile
    time error. See the "Known problems" section of this document.

  * Rewrote CustomFormat1 demo.

  * Added CustomFormat2 demo.

  * TDataDirection members has been renamed from ddGet and ddSet to ddRead and
    ddWrite.

  * All File Group Descritor and File Contents clipboard formats has been moved
    from the DragDropFile unit to the DragDropFormats unit.

  * File Contents support has been added to TTextDataFormat.
    The support is currently only enabled for drop sources.

  * Renamed TDropMultiTarget component to TDropComboTarget.
    Note: This will break applications which uses the TDropMultiTarget
    component. You can use the following technique to port application from
    previous releases:
    1) Install the new components.
    2) Repeat step 3-8 for all units which uses the TDropMultiTarget component.
    3) Make a backup of the unit (both pas and dfm file) just in case...
    4) Open the unit in the IDE.
    5) In the .pas file, replace all occurances of "TDropMultiTarget" with
       "TDropComboTarget".
    6) View the form as text.
    7) Replace all occurances of "TDropMultiTarget" with "TDropComboTarget".
    8) Save the unit.

  * Renamed a lot of demo files and directories.

  * Added work around for yet another bug in TStreamAdapter.

  * Added TCustomStringClipboardFormat as new base class for
    TCustomTextClipboardFormat.
    This changes the class hierachy a bit for classes which previously
    descended from TCustomTextClipboardFormat: All formats which needs zero
    termination now descend from TCustomTextClipboardFormat and the rest descend
    from TCustomStringClipboardFormat.
    Added TrimZeroes property.
    Fixed zero termination bug in TCustomTextClipboardFormat and generally
    improved handling of zero terminated strings.
    Disabled zero trim in TCustomStringClipboardFormat and enabled it in
    TCustomTextClipboardFormat.

23-jul-2000
  * Improved handling of long file names in DropHandler demo.
    Added work around for ParamStr bug.

  * Added TDataFormatAdapter component and adapter demo.
    TDataFormatAdapter is used to extend the existing source and target
    components with additional data format support without modifying them. It
    can be considered an dynamic alternative to the current TDropMultiTarget
    component.

17-jul-2000
  * TDropHandler component and DropHandler demo fully functional.

14-jul-2000
  * Tested with C++Builder 5.

  * Fixed sporadic integer overflow bug in DragDetectPlus function.

  * Added shell drop handler support with TDropHandler component.
    This is a work in progress and is not yet functional.

1-jul-2000
  * Tested with Delphi 4.

  * Support for Windows 2000 inter application drag images.

  * TRawClipboardFormat and TRawDataFormat classes for support of arbitrary
    unknown clipboard formats.
    The classes are used internally in the TCustomDropSource.SetData method to
    support W2K drag images.

