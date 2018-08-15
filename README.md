
# FileMonitor

FileMonitor is a Windows service that monitors one or more folders and subfolders for new and modified files. The typical flow is:

* Notification for a file being created or modified.
* Rules may be defined for dependencies; for example, a **`.map`** file requires a like-named **`.xtr`** file. The syntax for this dependency is **`"map>xtr"`**.
* Rules may also be defined for exclusion; for example, if a **`.mmv`** file exists then do not process any events for these files. The syntax is **`"map>xtr<mmv"`**. 
* Regular expressions are supported for folders and files to monitor.
* Define a **`verify`** procedure to run to ensure that the I/O operations are complete. Verify procedures may be specified for each type of file or folder monitored and may be a method in a COM object or Statit macro.
* Define an **`action`** procedure to run after a successful verify. Action procedures may be specified for each type of file or folder monitored and may be a method in a COM object or Statit macro.
* Define a **`dispatcher`** procedure to run a procedure or Statit macro based on a Windows event. For example, if the action procedure uses a queue then the dispatcher may fire on items in the queue.
* Each folder monitored utilizes 3 threads.
    * A thread for the monitoring.
    * A thread for the verification and action procedures.
    * A thread for scavenging. Scavenging is performed at startup to process folders and files that were created or modified while the service wasn't running. Surges from large copies or downloads may result in Windows not firing an event for all changes. When a surge is detected a scavenge operation is initiated. Scavenges may also be configured to run on an interval; e.g., every hour. Scavenging may be I/O intensive so a **`governor`** may be defined to limit the number of simultaneous scavenges.


