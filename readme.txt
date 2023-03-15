// -----------------------------------------------------------------------------
// Project:         Drag and Drop Components
// Component Names: TDropTextSource, TDropFileSource
//                  TDropTextTarget, TDropFileTarget.
// Modules:         DropSource, DropTarget.
// Description:     Implements Dragging & Dropping of text and files
//                  TO and FROM your application.
// Version:	        3.0
// Date:            22-SEP-1998
// Target:          Win32, Delphi 3 & 4
// Authors:         Angus Johnson, ajohnson@rpi.net.au
//                  (TDropTextSource, TDropFileSource,
//                  TDropTextTarget, TDropFileTarget, and demo.)
//
//                  Anders Melander, anders@melander.dk ::: http://www.melander.dk
//                  (TDropTextSource, TDropFileSource, and Delphi 4 compatability.)
// Copyright:       ©1998 Angus Johnson & Anders Melander

// -----------------------------------------------------------------------------
// You are free to use this source but please give us credit for our work.
// If you make improvements or derive new components from this code,
// we would very much like to see your improvements. FEEDBACK IS WELCOME.
// -----------------------------------------------------------------------------

// -----------------------------------------------------------------------------
// History & Usage: Also see module headers.
// -----------------------------------------------------------------------------

// History:
// dd/mm/yy  Version  Changes
// --------  -------  ----------------------------------------
// 22.09.98  3.0      * Shortcuts (links) now enabled.
//                    * Scrap files now enabled.
//                    * Demo modified - separate thread to monitor directory changes.
//                    * TDropTarget bug fix where StgMediums not cleaned up.
//                    * Some bugs still with NT4 :-)
// 08.09.98  2.0      * DropTarget module included.
// -----------------------------------------------------------------------------
