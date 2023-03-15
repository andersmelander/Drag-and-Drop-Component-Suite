  // -----------------------------------------------------------------------------
  // Project:         Drag and Drop Source Components
  // Component Names: TDropTextSource, TDropFileSource
  // Module:          DropSource
  // Description:     Implements Dragging & Dropping of text and files
  //                  FROM your application TO another.
  // Version:	    1.4
  // Date:            19-AUG-1998
  // Target:          Win32, Delphi 3 & 4
  // Authors:         Angus Johnson, ajohnson@rpi.net.au
  //                  Anders Melander, anders@melander.dk
  //                                   http://www.melander.dk
  // Copyright        ©1998 Angus Johnson & Anders Melander
  // -----------------------------------------------------------------------------
  // You are free to use this source but please give us credit for our work.
  // If you make improvements or derive new components from this code,
  // we would very much like to see your improvements. FEEDBACK IS WELCOME.
  // -----------------------------------------------------------------------------
  
  // History:
  // dd/mm/yy  Version  Changes
  // --------  -------  ----------------------------------------
  // 19.08.98  1.4      * CopyToClipboard method added.
  //                    * Should now compile in Delphi 4. (see below)
  //                    * Another tidy up of the code.
  // 21.07.98  1.3      * Fixed a bug in Ver 1.2 where OnDrop event was never called.
  //                    * Now able to drag text to WordPad.
  //                    * Added OnFeedback Event.
  //                    * Added dtLink to TDragType enumeration but
  //                      still not able to get it to work.
  //                    * Code tidy up.
  // 19.07.98  1.2      * Changed TDragType enumeration type to a
  //                      TDragTypes set (AM's suggestion) ready to
  //                      implement dtLink later.
  //                    * Added known bugs to header and demo.
  // 17.07.98  1.1      * Reenabled end-user option to select either
  //                      Copy or Move operation while dragging.
  // 15.07.98  1.0      * Initial Delphi 3 component implementation of AM's
  //                      DropSource unit. I released a Delphi 2 D'n'D
  //                      component (TDragFilesSrc) over 12 months ago. However
  //                      with the significant changes in COM implementations
  //                      between Delphi versions I decided to use AM's code
  //                      as the springboard for my new Delphi 3 D'n'D components.
  //                      Thanks to Anders for the excellent start and
  //                      suggestions along the way!
  //                      
  // -----------------------------------------------------------------------------

  // Future Plans -
  // 1. Implement drag and drop of Links and Scrap Files.
  //    (So far I've drawn a blank. Any hints VERY welcome!)
  // -----------------------------------------------------------------------------
  
  // TDropTextSource -
  //   Public
  //      ....
  //      Text: string;
  //      function Execute: TDragResult; //drDropCopy, drDropMove, drDropCancel ...
  //      function CopyToClipboard: boolean;
  //    published
  //      property DragTypes: TDragTypes;  // [dtCopy, dtMove]
  //      property OnDrop: TDropEvent;
  //      property OnFeedback: TFeedbackEvent;
  //      ....

  // TDropFileSource -
  //   Public
  //      ....
  //      property Files: TStrings;
  //      function Execute: TDragResult; //drDropCopy, drDropMove, drDropCancel ...
  //      function CopyToClipboard: boolean;
  //    published
  //      property DragTypes: TDragTypes;  // [dtCopy, dtMove]
  //      property OnDrop: TDropEvent;
  //      property OnFeedback: TFeedbackEvent;
  //      ....

  //
  // Very brief examples of usage (see included demo for more detailed examples):
  //
  // TDropTextSource -
  // DropFileSource1.DragTypes := [dtCopy];
  // DropTextSource1.text := edit1.text;
  // if DropTextSource1.execute = drDropCopy then ShowMessage('It worked!');
  //
  // TDropFileSource -
  // DropFileSource1.DragTypes := [dtCopy, dtMove];
  // //  ie: let the user decide - Copy or Move.
  // //      Hold Ctrl down during drag -> Copy
  // //      Hold Shift down during drag -> Move
  // DropFileSource1.files.clear;
  // DropFileSource1.files.add('c:\autoexec.bat');
  // DropFileSource1.files.add('c:\config.sys');
  // res := DropFileSource1.execute;
  // if res = drDropCopy then ShowMessage('Files Copied!')
  // else if res = drDropMove then ShowMessage('Files Moved!');
  // -----------------------------------------------------------------------------
