AsyncTransferSource

This application demonstrates how to transfer data asynchronously via a stream.

The application uses a TDropEmptySource component and extends it with a
TDataFormatAdapter component.

While the drag/drop operation takes place in the main thread as always, the
actual data transfer is performed in a separate thread. The advantage of using a
thread is that the source application isn't blocked while the target application
processes the drop.
Note that this approach is normally only used when transferring large amounts of
data (e.g. file contents) or when the drop target is very slow.

While asynchronous drop targets (see the AsyncTransferTarget application)
requires the cooperation of the drop source, an asynchronous drop source can
perform its magic independently of the drop target.

Note that since the Drag and Drop Component Suite, and thus this demo, supports
asynchronous drop targets, you will not notice any difference between the two
modes of operation demonstrated by this application if you drop onto an
application that supports asynchronous transfer, such as the Windows Explorer.
