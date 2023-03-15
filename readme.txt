
Project:            Drag and Drop Component Suite.

Base Components:    TDropSource, TDropTarget
Derived Components: TDropFileSource, TDropFileTarget
                    TDropTextSource, TDropTextTarget
                    TDropBMPSource, TDropBMPTarget
                    TDropURLSource, TDropURLTarget
                    (TDropPidlSource, TDropPidlTarget)

Description:        Implements Dragging & Dropping of data between applications.

Version:            3.4
Date:               21-FEB-1999
Target:             Win32, Delphi 3 & 4, C++ Builder 3

Authors:            Angus Johnson, ajohnson@rpi.net.au
                    Anders Melander, anders@melander.dk, http://www.melander.dk
                    Graham Wideman, graham@sdsu.edu, http://www.wideman-one.com

Copyright:          ©1997-1999 Angus Johnson, Anders Melander & Graham Wideman.

You are free to use this source but please give us credit for our work.
FEEDBACK IS WELCOME.


--------------------------------------------------------------------------------
Installation:

Note:
If you are upgrading from a previous version - read 'UpgradingTo304.txt' first.

  1. Move the all the files in the Components directory of the Zip file
     into the folder where you wish save these D'n'D components.

  2. In the Delphi IDE select "File|Open" from the menu.

  3. Open the DragDropD3.dpk package file (or DragDropD4.dpk if you are using
     Delphi 4) from the folder where you saved the components.

  4. In the Package Dialog click the Install button.

  5. Add the path of the component units to Delphi's Library Path.

The newly installed components will appear in the Delphi Component Palette under
a new tab called 'DragDrop'.


--------------------------------------------------------------------------------
Components:

The Drag and Drop Component Suite contains the following units and components:

Unit                 Components        Use
..................   ...............   .........................................
DropSource.pas       TDropFileSource   Drag files from your application.
                     TDropTextSource   Drag text from your application.

DropTarget.pas       TDropFileTarget   Accept files dragged to your application.
                     TDropTextTarget   Accept text dragged to your application.
                     TDropDummy        Just displays a drag image.

DropBMPSource.pas    TDropBMPSource    Drag bitmaps from your application.

DropBMPTarget.pas    TDropBMPTarget    Accept bitmaps dragged to your application.

DropURLSource.pas    TDropURLSource    Drag URLs from your application.

DropURLTarget.pas    TDropURLTarget    Accept URLs dragged to your application.

DropPIDLSource.pas   TDropPIDLSource   Drag PIDLs from your application.

DropPIDLTarget.pas   TDropPIDLTarget   Accept PIDLs dragged to your application.


-----------------------------------------------------------------------------
-----------------------------------------------------------------------------
  History:
  dd/mm/yy  Version  Changes
  --------  -------  ----------------------------------------
  17.02.99  3.4      * NOTE: IF YOU ARE UPGRADING FROM A PREVIOUS VERSION - 
                       READ 'UpgradingTo304.TXT' FIRST.
                     * TDropSource: 
                       1. Images can now be displayed while dragging.
                       (Images will be displayed over registered drop targets
                       - although restricted to target windows in the 
                       same application as the source if running in WinNT.)
                       New properties: 
                         Images: TImageList;
                         ImageHotspotX: integer;
                         ImageHotspotY: integer;
                         ImageIndex: integer;
                         ShowImage: boolean;
                       (See the Detailed Demo for examples.)
                       2.CutToClipboard method added.
                       (CutToClipboard and CopyToClipboard will now also
                       render the CF_PREFERREDDROPEFFECT clipboard format.)
                     * TDropTextSource: 
                       Commented out two clipboard formats
                       (CF_FILECONTENTS, CF_FILEGROUPDESCRIPTOR) 
                       as they seem to prevent text drop into Word97.
                       (They were originally added to enable Text Scrap Files.)
                     * TDropFileSource:
                       CopyToClipboard bug when running in Win98 should now be fixed.
                       Files should no longer be moved instead of copied.
                       The move (cut) or  copy behaviour is now indicated to the 
                       target by setting a CF_PREFERREDDROPEFFECT clipboard dataobject.
                       (See TDropSource above.) 
                     * TDropTarget:
                       1. OnEnter(), OnDragOver() and OnDrop() event declarations have 
                       been changed to reflect the underlying IDropTarget methods.
                       However, the DataObject has not been passed as a parameter 
                       in these events as this interface pointer can be accessed 
                       through the DataObject property if really needed.
                       2. OnGetDropEffect() declaration has also been modified. This event
                       now overrides the 'default' behaviour. (See example in demo.)
                       3. PasteFromClipboard method added (and implemented in
                       descendant classes).
                       4. AutoScrolling when dragging near the Target window edge.
                       This feature also ensures a clean painting of a drag image.
                       5. Target TWinControl property added. Should rarely be needed
                       as the Target control is assigned with the Register() method. 
                       Use with care!
                       6. Images can now be displayed while dragging. 
                       New properties: 
                         ShowImage: boolean;
                     * TDropDummy: (New component) 
                       Target component to be used when just a drag image is 
                       desired to be displayed over a window (eg a form) but 
                       but where no rendering of the dataobject is needed.
                     * TDropBMPTarget:
                       Bug fix: HPalette resource leak fixed.
                     * TDropURLSource:
                       Title property added.
                       Blank spaces now allowed in URL filename (Title).
                     * TDropURLTarget:
                       1. Title property added.
                       2. Minor bug fix.
                     * TDropPIDLSource & TDropPIDLTarget: (New components) 
                     * Improved (detailed) demo.
  16.11.98  3.3      * Help file, and two simple demos added.
                     * TDropBMPSource & TDropBMPTarget:
                       Added DIB support.
  22.10.98  3.2      * TDropSource:
                       1. AddFormatEtc() method added.
                       2. GiveFeedback() method declaration changed
                       to enable use of custom cursors.
                     * TDropFileSource:
                       Fixed bug in Files property.
                     * TDropBMPSource & TDropBMPTarget: (New components) 
                     * TDropURLSource & TDropURLTarget: (New components) 
  01.10.98  3.1      * Removed hooking of source and target windows introduced 
                       in previous versions. Although hooking simplified the use
                       of these components, it became apparent that hooking within 
                       components introduces the risk of a program crash when 
                       another component or the component user also hooks the
                       same window. (To avoid the possiblilty of a crash the order of 
                       unhooking has to be exactly the reverse order of the hooking,
                       but this is beyond the control of a component.)
                     * TDropSource:
                       1. Removed the DropSource TWincontrol property (see above).
                       2. Removed the "AutoDrag" feature (see above).
                     * TDropTarget:
                       1. Removed the TargetWindow TWincontrol property (see above).
                       2. Removed Enabled property.
                       3. Register() and Unregister() methods added.
                     * TDropURLSource & TDropURLTarget: (New components)
  22.09.98  3.0      * TDropSource:
                       1. Shortcuts (link) enabled.
                       2. DropSource property added to implement 
                       autodrag feature. (Removed in Ver 3.1)
                       3. New events - OnStartDrag and OnEndDrag added. (Removed in Ver 3.1)
                       4. DoEnumFormatEtc() no longer declared abstract.
                     * TDropTarget:
                       GetValidDropEffect() moved to protected section.
                     * TDropTextSource:
                       Scrap files enabled. (Disabled again temporarily? in Version 3.4)
                     * TDropFileTarget: 
                       1. Fixed a bug where StgMediums weren't released. (oops!)
                       2. Shortcuts (links) enabled.
                     * TDropTextTarget: 
                       Fixed a bug where StgMediums weren't released.
  08.09.98  2.0      * TDropTarget: (New component)
  31.08.98  1.5      * Improved Demo.
  19.08.98  1.4      * TDropSource:
                       CopyToClipboard method added.
  21.07.98  1.3      * TDropSource:
                       1. OnFeedback event added.
                       2. dtLink added to TDragType enumeration but
                       still not able to get it to work.
                     * TDropTextSource:
                       Bug Fix. Enabled drag text to WordPad.
  19.07.98  1.2      * TDropSource:
                       Changed TDragType enumeration to a
                       TDragTypes set.
  17.07.98  1.1      * TDropSource:
                       Reenabled end-user option to select either
                       Copy or Move operation while dragging.
  15.07.98  1.0      * Initial Delphi 3 component implementation of AM's
                       DropSource unit.
-----------------------------------------------------------------------------
-----------------------------------------------------------------------------
