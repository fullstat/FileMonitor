
#include "stdafx.h"
#include <stdio.h>
#include <process.h>
#include "monitorshare.h"
#include "eserver.h"
#include "smmerrors.h"
#include "regexp.h"
#include <time.h>
#include <limits.h>
#include <math.h>
#include <eh.h>

// the document for GetExitCodeThread() references a STILL_ALIVE
// value and states that it is 259; however, I cannot locate it in
// the VC++ or Platform SDK header files.

#ifndef STILL_ALIVE
#define STILL_ALIVE 259
#endif

#define SCAVENGE_TIMEOUT 5      // 5 seconds (interval)

// the default action object

#import "imports/smmapactions.dll" no_namespace, raw_interfaces_only

// default extensions, MAP and XTR for EG SECS2 (not regular expression)

const char * const defaultExts = "map>xtr<mmv";
const unsigned long numDefaultExts = 1;

// default lot directory (regular expression)

const char * const defaultLotDir = "[Ll][Oo][Tt][0-9][0-9]*";

const unsigned long maxRetryCount = 3;
const unsigned long retryDelay = 50;   // milliseconds

void DirList::allocateProc(size_t first, size_t last)
{
   do
   {
      DirEntry *p = (DirEntry *) nodeFast(first);

      p->pathName[0] = '\0';
      p->fileName[0] = '\0';
      p->modifiedTime = 0.0;
      p->creationTime = 0.0;
      p->fileAttrib = 0;
      p->notify = false;
   }
   while (++first <= last);
}

void DirList::deallocateProc(size_t first, size_t last)
{
   do
   {
      DirEntry *p = (DirEntry *) nodeFast(first);
      if (p)
      {
         p->pathName[0] = '\0';
         p->fileName[0] = '\0';
         p->modifiedTime = 0.0;
         p->creationTime = 0.0;
         p->fileAttrib = 0;
         p->notify = false;
      }
   }
   while (++first <= last);
}


void HoldForAssoc::allocateProc(size_t first, size_t last)
{
   do
   {
      HoldForAssocPath *p = (HoldForAssocPath *) nodeFast(first);

      p->inUse = false;
      p->pathName = 0;
      p->timeStamp = 0.0;
   }
   while (++first <= last);
}

void HoldForAssoc::deallocateProc(size_t first, size_t last)
{
   do
   {
      HoldForAssocPath *p = (HoldForAssocPath *) nodeFast(first);
      if (p)
      {
         delete p->pathName;
         p->pathName = 0;
         p->timeStamp = 0.0;

         p->inUse = false;
      }
   }
   while (++first <= last);
}


class LocalPointersThread
{
   public:
      LocalPointersThread() {}
      ~LocalPointersThread() { CoUninitialize(); ExitThread(0); }

};

DWORD WINAPI shareThread(LPVOID lpParam)
{
   static const char *funcName = "shareThread()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   LocalPointersThread LP;

   MonitorShare *pShare = (MonitorShare *) lpParam;

   CoInitializeEx(NULL,COINIT_MULTITHREADED);

   try
   {
      pShare->run();
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
      return Error_Failure;
   }
   catch (...)
   {
      char localBuf[128];

      sprintf(localBuf, "Exception in prober share monitoring thread: Id=%i",
            GetCurrentThreadId());
      logErrorMsg(errSmmUnknownException, funcName, localBuf);
      return Error_Failure;
   }

   return Error_Success;
}

DWORD WINAPI queueThread(LPVOID lpParam)
{
   static const char *funcName = "queueThread()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   LocalPointersThread LP;

   MonitorShare *pShare = (MonitorShare *) lpParam;

   CoInitializeEx(NULL,COINIT_MULTITHREADED);

   try
   {
      pShare->runQueue();
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
      return Error_Failure;
   }
   catch (...)
   {
      char localBuf[128];

      sprintf(localBuf, "Exception in queue monitoring thread: Id=%i",
            GetCurrentThreadId());
      logErrorMsg(errSmmUnknownException, funcName, localBuf);
      return Error_Failure;
   }

   return Error_Success;
}

DWORD WINAPI scavengeThread(LPVOID lpParam)
{
   static const char *funcName = "scavengeThread()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   LocalPointersThread LP;

   MonitorShare *pShare = (MonitorShare *) lpParam;

   CoInitializeEx(NULL,COINIT_MULTITHREADED);

   try
   {
      pShare->runScavenge();
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
      return Error_Failure;
   }
   catch (...)
   {
      char localBuf[128];

      sprintf(localBuf, "Exception in scavenge monitoring thread: Id=%i",
            GetCurrentThreadId());
      logErrorMsg(errSmmUnknownException, funcName, localBuf);
      return Error_Failure;
   }

   return Error_Success;
}



MonitorShare::MonitorShare()
{
   static const char *funcName = "MonitorShare::MonitorShare()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   m_initialized = false;
   m_didSetup = false;
   m_useServerPool = false;
   m_inShutdown = false;
   m_requestShutdown = false;

   m_monitorId = 0;
   m_dirPath = 0;
   m_verifyMap = 0;
   m_verifyLot = 0;
   m_mapAction = 0;
   m_lotAction = 0;
   m_scavenge = 0;
   m_readChangeBuffer = 0;

   m_lenVerifyMap = 0;
   m_lenMapAction = 0;
   m_lenVerifyLot = 0;
   m_lenLotAction = 0;
   m_lenScavenge = 0;

   m_scavengeStart = 30;  // first scavenge starts 30 seconds after monitoring
   m_scavengeTime = 7200; // default is 2 hours

   m_compPort = 0;

   m_threadHandle = 0;
   m_threadEvent = 0;
   m_dirHandle = 0;
   m_threadId = 0;

   m_queueSemaphore = 0;
   m_queueThreadId = 0;
   m_queueThreadHandle = 0;
   m_queueEvent = 0;

   m_scavengeThreadId = 0;
   m_scavengeThreadHandle = 0;
   m_scavengeEvent = 0;

   m_scavengeSurgeTrigger = 0;
   m_scavengeSurgeExceeded = false;
   m_triggerScavenge = false;

   m_queueHead = 0;
   m_queueTail = 0;
   m_queueDepth = 0;
   m_maxQueueDepth = 0;

   m_extInfo.setCapacity(0);
   m_mapHold.setCapacity(0);

   m_numberOnHold = 0;

   InitializeCriticalSection(&m_sLock);
   InitializeCriticalSection(&m_qLock);
}

MonitorShare::~MonitorShare()
{
   static const char *funcName = "MonitorShare::~MonitorShare()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   shutdown();

   DeleteCriticalSection(&m_qLock);
   DeleteCriticalSection(&m_sLock);
}


class LocalPointersSetup
{
   public:
      char *string;
      HANDLE handle;

      LocalPointersSetup() : string(0), handle(0) {}
      ~LocalPointersSetup() { if (string) delete string;
                              if (handle) CloseHandle(handle); }
};

long MonitorShare::setup(const char *monitorId, const char *dirPath,
               unsigned long numExtensions, const char **extensions,
               const char *lotDir, const char *verifyMap,
               const char *mapAction, const char *verifyLot,
               const char *lotAction, const char *scavenge,
               unsigned long scavengeTime, unsigned long scavengeStart,
               unsigned long scavengeSurgeTrigger, bool useServerPool)
{
   static const char *funcName = "MonitorShare::setup()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   LocalPointersSetup LP;
   unsigned long ls;

   if (m_didSetup)
   {
      shutdown();
      m_didSetup = false;
   }

   if (dirPath == 0 || strlen(dirPath) == 0)
   {
      logErrorMsg(errSmmBadParameter, funcName,
         "A null or empty directory path was specified.");
      return Error_Failure;
   }

   try
   {
      ls = strlen(dirPath);
      LP.string = new char[ls+3];      // leave room for trailing '\'
      if (LP.string == 0)
      {
         logErrorMsg(errSmmResource, funcName,
            "Failed to allocate memory for the directory path.");
         return Error_Failure;
      }

      strcpy(LP.string, dirPath);
      if (dirPath[ls-1] == '\\')       // remove for CreateFile()
      {
         LP.string[ls-1] = '\0';
         ls = ls - 1;
      }

      LP.handle = CreateFile(LP.string, FILE_LIST_DIRECTORY,
             FILE_SHARE_READ | FILE_SHARE_DELETE | FILE_SHARE_WRITE, 0,
             OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OVERLAPPED, 0);
      if (LP.handle == INVALID_HANDLE_VALUE)
      {
         char localBuf[2048];
         char errBuf[1024];

         DWORD eCode = GetLastError();
         if (! FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, eCode, 0,
               errBuf, 1024, 0))
         {
            sprintf(localBuf,"Failed to open '%s' for monitoring. Error code=%d.",
               LP.string, eCode);
         }
         else
         {
            sprintf(localBuf,
               "Failed to open '%s' for monitoring (%s). Error code=%d.",
                  LP.string, errBuf, eCode);
         }
         logErrorMsg(errSmmIO, funcName, localBuf);
         return Error_Failure;
      }

      m_dirHandle = LP.handle;
      LP.handle = 0;     // for dtor

      LP.string[ls++] = '\\';
      LP.string[ls] = '\0';
      m_dirPath = LP.string;
      LP.string = 0;     // for dtor

      ls = strlen(monitorId);
      m_monitorId = new char[ls+1];
      if (m_monitorId == 0)
      {
         logErrorMsg(errSmmResource, funcName,
            "Failed to allocate memory for the monitor id.");
         return Error_Failure;
      }

      strcpy(m_monitorId, monitorId);

// process the extensions

      if (! parseExtensions(extensions, numExtensions))
         return Error_Failure;

      if (lotDir && strlen(lotDir) > 0)
      {
         if (m_lotDir.set(lotDir))
         {
            logErrorMsg(errSmmResource, funcName,
               "Failed to allocate memory for the lot directory name.");
            return Error_Failure;
         }

         if (m_lotDir.testRegExp() == false)
         {
            char localBuf[_MAX_PATH+101];

            sprintf(localBuf,
                  "Invalid lot folder regular expression '%s' for %s.",
                  m_lotDir.value(),m_monitorId);
            logErrorMsg(errSmmBadRegExp, 0, localBuf);
            return Error_Failure;
         }
      }
      else
      {
         if (m_lotDir.set(defaultLotDir))
         {
            logErrorMsg(errSmmResource, funcName,
               "Failed to allocate memory for the lot directory name.");
            return Error_Failure;
         }
      }

      if (verifyMap && strlen(verifyMap) > 0)
      {
         m_lenVerifyMap = strlen(verifyMap);
         m_verifyMap = new char[m_lenVerifyMap+1];
         if (m_verifyMap == 0)
         {
            logErrorMsg(errSmmResource, funcName,
               "Failed to allocate memory for the verify map proc name.");
            return Error_Failure;
         }

         strcpy(m_verifyMap, verifyMap);
      }

      if (mapAction && strlen(mapAction) > 0)
      {
         m_lenMapAction = strlen(mapAction);
         m_mapAction = new char[m_lenMapAction+1];
         if (m_mapAction == 0)
         {
            logErrorMsg(errSmmResource, funcName,
               "Failed to allocate memory for the map action.");
            return Error_Failure;
         }

         strcpy(m_mapAction, mapAction);
      }

      if (verifyLot && strlen(verifyLot) > 0)
      {
         m_lenVerifyLot = strlen(verifyLot);
         m_verifyLot = new char[m_lenVerifyLot+1];
         if (m_verifyLot == 0)
         {
            logErrorMsg(errSmmResource, funcName,
               "Failed to allocate memory for the verify lot proc name.");
            return Error_Failure;
         }

         strcpy(m_verifyLot, verifyLot);
      }

      if (lotAction && strlen(lotAction) > 0)
      {
         m_lenLotAction = strlen(lotAction);
         m_lotAction = new char[m_lenLotAction+1];
         if (m_lotAction == 0)
         {
            logErrorMsg(errSmmResource, funcName,
               "Failed to allocate memory for the lot action.");
            return Error_Failure;
         }

         strcpy(m_lotAction, lotAction);
      }

// need to look up CLSID for map action object if either are default

      if (m_lenLotAction == 0 || m_lenMapAction == 0)
      {
         HRESULT hr = CLSIDFromProgID(OLESTR("SMMapActions.clsMapActions"),
                           &m_clsIdMapAction);
         if (FAILED(hr))
         {
            char errBuf[BUFSIZ+1];
            char localBuf[BUFSIZ+256];

            if (! FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, hr, 0,
                  errBuf, BUFSIZ, 0))
            {
               sprintf(localBuf,
                  "Failed to lookup the SMMapActions.clsMapActions CLSID. Error code=%x.",
                  hr);
            }
            else
            {
               sprintf(localBuf,
                  "Failed to lookup the SMMapActions.clsMapActions CLSID (%s). Error code=%x.",
                  errBuf, hr);
            }
            logErrorMsg(errSmmConfiguration, funcName, localBuf);
            return Error_Failure;
         }
      }

      if (scavenge && strlen(scavenge) > 0)
      {
         m_lenScavenge = strlen(scavenge);
         m_scavenge = new char[m_lenScavenge+1];
         if (m_scavenge == 0)
         {
            logErrorMsg(errSmmResource, funcName,
               "Failed to allocate memory for the scavenge action.");
            return Error_Failure;
         }

         strcpy(m_scavenge, scavenge);
      }

      if (scavengeTime >= 0)
         m_scavengeTime = scavengeTime;

      if (scavengeStart >= 0)
         m_scavengeStart = scavengeStart;

      if (scavengeSurgeTrigger >= 0)
         m_scavengeSurgeTrigger = scavengeSurgeTrigger;

      m_useServerPool = useServerPool;

      m_readChangeBuffer = new unsigned char[ReadChangeBufferLen+1];
      if (m_readChangeBuffer == 0)
      {
         logErrorMsg(errSmmResource, funcName,
            "Failed to allocate memory for the read change buffer.");
         return Error_Failure;
      }

      memset(&m_readChangeOverlapped, 0, sizeof(OVERLAPPED));

      if (m_notifyCheck.init() == false)
         return Error_Failure;

      if (logTraceLevel >= logTraceLevelDetail)
      {
         char localBuf[_MAX_PATH+151];

         sprintf(localBuf, "Monitoring: '%s', id: %s", m_dirPath, m_monitorId);
         logTrace(funcName, localBuf);

         sprintf(localBuf, "   Lot directory: '%s'", m_lotDir.value());
         logTrace(funcName, localBuf);

         if (m_verifyMap)
         {
            sprintf(localBuf, "   Verify map: '%s'", m_verifyMap);
            logTrace(funcName, localBuf);
         }

         if (m_mapAction)
         {
            sprintf(localBuf, "   Map action: '%s'", m_mapAction);
            logTrace(funcName, localBuf);
         }

         if (m_verifyLot)
         {
            sprintf(localBuf, "   Verify lot: '%s'", m_verifyLot);
            logTrace(funcName, localBuf);
         }

         if (m_lotAction)
         {
            sprintf(localBuf, "   Lot action: '%s'", m_lotAction);
            logTrace(funcName, localBuf);
         }
      }

      m_didSetup = true;
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
      return Error_Failure;
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName,
            "Exception occurred during setup()");
      return Error_Failure;
   }
   return Error_Success;
}


long MonitorShare::start()
{
   static const char *funcName = "MonitorShare::start()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   if (! m_didSetup)
   {
      logErrorMsg(errSmmSetupNotCalled, funcName,
         "The setup() method must be call prior to the start() method.");
      return Error_Failure;
   }

   m_compPort = CreateIoCompletionPort(m_dirHandle, 0, (DWORD) this, 0);
   if (! m_compPort)
   {
      char localBuf[128];

      DWORD eCode = GetLastError();
      sprintf(localBuf,"CreateIoCompletionPort() returned %ld.",eCode);
      logErrorMsg(errSmmResource, funcName, localBuf);
      return Error_Failure;
   }

   m_queueSemaphore = CreateSemaphore(0, 0, LONG_MAX, 0);
   if (! m_queueSemaphore)
   {
      char localBuf[128];

      DWORD eCode = GetLastError();
      sprintf(localBuf,"CreateSemaphore() returned %ld.",eCode);
      logErrorMsg(errSmmResource, funcName, localBuf);
      return Error_Failure;
   }

   m_queueEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
   if (m_queueEvent)
   {
      m_queueThreadHandle = CreateThread(NULL, 0, queueThread, (void *) this,
                           0,(unsigned long *) &m_queueThreadId);

      if (m_queueThreadHandle)
      {
         WaitForSingleObject(m_queueEvent, INFINITE);
         ResetEvent(m_queueEvent);
      }
      else
      {
         char localBuf[128];

         DWORD eCode = GetLastError();
         sprintf(localBuf,"CreateThread(queue) returned %ld.",eCode);
         logErrorMsg(errSmmResource, funcName, localBuf);
         return Error_Failure;
      }
   }
   else
   {
      char localBuf[128];

      DWORD eCode = GetLastError();
      sprintf(localBuf,"CreateEvent(queue) returned %ld.",eCode);
      logErrorMsg(errSmmResource, funcName, localBuf);
      return Error_Failure;
   }

   if (m_scavengeTime > 0 && m_scavengeStart > 0)
   {
      m_scavengeEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
      if (m_scavengeEvent)
      {
         m_scavengeThreadHandle = CreateThread(NULL, 0, scavengeThread,
                      (void *) this, 0,(unsigned long *) &m_scavengeThreadId);

         if (m_scavengeThreadHandle)
         {
            WaitForSingleObject(m_scavengeEvent, INFINITE);
            ResetEvent(m_scavengeEvent);
         }
         else
         {
            char localBuf[128];

            DWORD eCode = GetLastError();
            sprintf(localBuf,"CreateThread(scavenge) returned %ld.",eCode);
            logErrorMsg(errSmmResource, funcName, localBuf);
            return Error_Failure;
         }
      }
      else
      {
         char localBuf[128];

         DWORD eCode = GetLastError();
         sprintf(localBuf,"CreateEvent(scavenge) returned %ld.",eCode);
         logErrorMsg(errSmmResource, funcName, localBuf);
         return Error_Failure;
      }
   }

   m_threadEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
   if (m_threadEvent)
   {
      m_threadHandle = CreateThread(NULL, 0, shareThread, (void *) this,
                           0,(unsigned long *) &m_threadId);

      if (m_threadHandle)
      {
         WaitForSingleObject(m_threadEvent, INFINITE);
         ResetEvent(m_threadEvent);
      }
      else
      {
         char localBuf[128];

         DWORD eCode = GetLastError();
         sprintf(localBuf,"CreateThread(monitor) returned %ld.",eCode);
         logErrorMsg(errSmmResource, funcName, localBuf);
         return Error_Failure;
      }
   }
   else
   {
      char localBuf[128];

      DWORD eCode = GetLastError();
      sprintf(localBuf,"CreateEvent(monitor) returned %ld.",eCode);
      logErrorMsg(errSmmResource, funcName, localBuf);
      return Error_Failure;
   }

   DWORD notUsed = 0;
   BOOL rStatus = ReadDirectoryChangesW(m_dirHandle, m_readChangeBuffer,
         ReadChangeBufferLen, TRUE, FILE_NOTIFY_CHANGE_LAST_WRITE |
         FILE_NOTIFY_CHANGE_DIR_NAME | FILE_NOTIFY_CHANGE_FILE_NAME,
         &notUsed, &m_readChangeOverlapped, 0);
   if (rStatus == FALSE)
   {
      char localBuf[BUFSIZ+256];
      char errBuf[BUFSIZ+1];

      DWORD eCode = GetLastError();
      if (! FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, eCode, 0,
            errBuf, BUFSIZ, 0))
      {
         sprintf(localBuf, "ReadDirectoryChangesW() returned %ld for %s '%s'.",
               eCode, m_monitorId, m_dirPath);
      }
      else
      {
         sprintf(localBuf,
            "ReadDirectoryChangesW() returned %ld (%s) for %s '%s'.",
               eCode, errBuf, m_monitorId, m_dirPath);
      }
      logErrorMsg(errSmmIO, funcName, localBuf);
      return Error_Failure;
   }

   return Error_Success;
}


void MonitorShare::run()
{
   static const char *funcName = "MonitorShare::run()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   m_initialized = true;
   SetEvent(m_threadEvent);

   if (logTraceLevel >= logTraceLevelDetail)
   {
      char localBuf[_MAX_PATH+151];

      sprintf(localBuf, "Monitoring thread started for %s '%s'", m_monitorId,
            m_dirPath);
      logTrace(funcName, localBuf);
   }

   OVERLAPPED *pOverlapped = 0;
   DWORD numBytes = 0;
   DWORD cbOffset = 0;
   MonitorShare *pKey = 0;
   bool done = false;

   do
   {
      try
      {
         BOOL rStatus = GetQueuedCompletionStatus(m_compPort, &numBytes,
               (LPDWORD) &pKey, &pOverlapped, INFINITE);

         if (rStatus && pKey)
         {
            FILE_NOTIFY_INFORMATION *notifyBuf;
            char fileName[_MAX_PATH*2+1];
            unsigned long fCount = 0;

            notifyBuf = (FILE_NOTIFY_INFORMATION *) pKey->readBuffer();

            do
            {
               if (m_inShutdown || _Module.inShutdown())
               {
                  done = true;
                  break;
               }

               cbOffset = notifyBuf->NextEntryOffset;

               if (notifyBuf->FileNameLength > 0)
               {
                  char pathBuf[_MAX_PATH*2+1];

                  wcstombs(fileName, notifyBuf->FileName,
                              notifyBuf->FileNameLength);
                  fileName[notifyBuf->FileNameLength/2] = '\0';
                  fCount++;

                  strcpy(pathBuf,m_dirPath);
                  strcat(pathBuf,fileName);

                  if (notifyBuf->Action == FILE_ACTION_ADDED ||
                      notifyBuf->Action == FILE_ACTION_MODIFIED ||
                      notifyBuf->Action == FILE_ACTION_RENAMED_NEW_NAME)
                  {
                     addToQueue(pathBuf, notifyBuf->Action);
                  }
               }

               notifyBuf = (FILE_NOTIFY_INFORMATION *)
                     ((unsigned char *) notifyBuf + cbOffset);
               if ((unsigned char *) notifyBuf - pKey->readBuffer() >
                        ReadChangeBufferLen)
                  cbOffset = 0;         // gone too far
            }
            while (cbOffset > 0);

            if (! done)
            {
               DWORD notUsed = 0;

               rStatus = ReadDirectoryChangesW(pKey->dirHandle(),
                  pKey->readBuffer(), ReadChangeBufferLen, TRUE,
                  FILE_NOTIFY_CHANGE_LAST_WRITE | FILE_NOTIFY_CHANGE_DIR_NAME |
                  FILE_NOTIFY_CHANGE_FILE_NAME, &notUsed,
                  pKey->readOverlapped(), 0);
               if (rStatus == FALSE)
               {
                  char localBuf[BUFSIZ+256];
                  char errBuf[BUFSIZ+1];

                  DWORD eCode = GetLastError();
                  if (! FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, eCode, 0,
                        errBuf, BUFSIZ, 0))
                  {
                     sprintf(localBuf,
                        "ReadDirectoryChangesW() returned %ld for %s '%s'.",
                           eCode, m_monitorId, m_dirPath);
                  }
                  else
                  {
                     sprintf(localBuf,
                        "ReadDirectoryChangesW() returned %ld (%s) for %s '%s'.",
                           eCode, errBuf, m_monitorId, m_dirPath);
                  }

                  logErrorMsg(errSmmIO, funcName, localBuf);
                  setRequestShutdown(true);
               }
            }
         }
      }
      catch (_com_error &e)
      {
         logErrorMsg(e.Error(), funcName, e.ErrorMessage());
      }
      catch (...)
      {
         logErrorMsg(errSmmUnknownException, funcName, 0);
      }
   } while ((! done) && pKey);

   if (logTraceLevel >= logTraceLevelDetail)
   {
      char localBuf[_MAX_PATH+151];

      sprintf(localBuf, "Monitoring thread completed for %s '%s'", m_monitorId,
            m_dirPath);
      logTrace(funcName, localBuf);
   }
}


const int THREAD_SHUTDOWN_WAIT = 5000;   // 5 seconds

void MonitorShare::shutdown()
{
   static const char *funcName = "MonitorShare::shutdown()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   EnterCriticalSection(&m_sLock);

   try
   {
      if (m_initialized)
      {
         m_initialized = false;
         m_inShutdown = true;

         // stop the monitoring thread first

         BOOL postSuccess = PostQueuedCompletionStatus(m_compPort, 0, 0, 0);
         DWORD reason = WAIT_TIMEOUT;

         if (postSuccess == TRUE)
         {
            reason = MsgWaitForMultipleObjects(1, &m_threadHandle, FALSE,
                           THREAD_SHUTDOWN_WAIT, QS_ALLINPUT);
         }

         if (postSuccess == FALSE)
         {
            char localBuf[256];

            DWORD eCode = GetLastError();
            sprintf(localBuf,
                  "PostQueuedCompletionStatus() returned %ld for %s.",
                  eCode, m_monitorId);
            logErrorMsg(errSmmPostThread, funcName, localBuf);
         }
         else if (reason == WAIT_TIMEOUT)
         {
            char localBuf[256];

            sprintf(localBuf,
                  "MonitorShare::shutdown() monitoring thread timeout for %s.",
                  m_monitorId);
            logErrorMsg(errSmmWaitTimeout, funcName, localBuf);
         }

         // stop the scavenge worker thread, if there is one.

         // this thread may have already
         // exited due to detecting the shutdown so if the post fails I
         // won't worry about it.

         DWORD exitCode = 0;
         if (m_scavengeThreadId > 0)
         {
            if (! GetExitCodeThread(m_scavengeThreadHandle, &exitCode))
               exitCode = 0;

            if (exitCode == STILL_ALIVE)
            {
               reason = WAIT_TIMEOUT;
               postSuccess = PostThreadMessage(m_scavengeThreadId, WM_QUIT, 0, 0);
               if (postSuccess == TRUE)
               {
                  reason = MsgWaitForMultipleObjects(1, &m_scavengeThreadHandle,
                                 FALSE, THREAD_SHUTDOWN_WAIT, QS_ALLINPUT);
               }

               if (postSuccess == FALSE)
               {
                  char localBuf[256];

                  DWORD eCode = GetLastError();
                  if (eCode != ERROR_INVALID_THREAD_ID)
                  {
                     sprintf(localBuf,
                        "PostThreadMessage(scavenge) returned %ld for %s.",
                        eCode, m_monitorId);
                     logErrorMsg(errSmmPostThread, funcName, localBuf);
                  }
               }
               else if (reason == WAIT_TIMEOUT)
               {
                  char localBuf[256];
   
                  sprintf(localBuf,
                        "MonitorShare::shutdown() scavenge thread timeout for %s.",
                        m_monitorId);
                  logErrorMsg(errSmmWaitTimeout, funcName, localBuf);
               }
            }
         }

         // stop the queue worker thread.  this thread may have already
         // exited due to detecting the shutdown so if the post fails I
         // won't worry about it.

         exitCode = 0;
         if (! GetExitCodeThread(m_queueThreadHandle, &exitCode))
            exitCode = 0;

         if (exitCode == STILL_ALIVE)
         {
            reason = WAIT_TIMEOUT;
            postSuccess = PostThreadMessage(m_queueThreadId, WM_QUIT, 0, 0);
            if (postSuccess == TRUE)
            {
               reason = MsgWaitForMultipleObjects(1, &m_queueThreadHandle,
                           FALSE, THREAD_SHUTDOWN_WAIT, QS_ALLINPUT);
            }

            if (postSuccess == FALSE)
            {
               char localBuf[256];

               DWORD eCode = GetLastError();
               if (eCode != ERROR_INVALID_THREAD_ID)
               {
                  sprintf(localBuf,
                     "PostThreadMessage(queue) returned %ld for %s.",eCode,
                     m_monitorId);
                  logErrorMsg(errSmmPostThread, funcName, localBuf);
               }
            }
            else if (reason == WAIT_TIMEOUT)
            {
               char localBuf[256];

               sprintf(localBuf,
                     "MonitorShare::shutdown() queue thread timeout for %s.",
                     m_monitorId);
               logErrorMsg(errSmmWaitTimeout, funcName, localBuf);
            }
         }

         // in case of a problem during startup signal the event so that
         // the infinite wait will proceed

         if (m_threadEvent)
            SetEvent(m_threadEvent);

         if (m_scavengeEvent)
            SetEvent(m_scavengeEvent);

         if (m_queueEvent)
            SetEvent(m_queueEvent);
      }
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName,
            "Exception occurred during shutdown()");
   }

   LeaveCriticalSection(&m_sLock);

   if (m_threadHandle)
   {
      CloseHandle(m_threadHandle);
      m_threadHandle = 0;
   }

   if (m_threadEvent)
   {
      CloseHandle(m_threadEvent);
      m_threadEvent = 0;
   }

   if (m_dirHandle)
   {
      CloseHandle(m_dirHandle);
      m_dirHandle = 0;
   }

   if (m_scavengeThreadHandle)
   {
      CloseHandle(m_scavengeThreadHandle);
      m_scavengeThreadHandle = 0;
   }

   if (m_scavengeEvent)
   {
      CloseHandle(m_scavengeEvent);
      m_scavengeEvent = 0;
   }

   if (m_queueSemaphore)
   {
      CloseHandle(m_queueSemaphore);
      m_queueSemaphore = 0;
   }

   if (m_queueThreadHandle)
   {
      CloseHandle(m_queueThreadHandle);
      m_queueThreadHandle = 0;
   }

   if (m_queueEvent)
   {
      CloseHandle(m_queueEvent);
      m_queueEvent = 0;
   }

   if (m_compPort)
   {
      CloseHandle(m_compPort);
      m_compPort = 0;
   }

   m_scavengeThreadId = 0;
   m_queueThreadId = 0;
   m_threadId = 0;

   delete m_verifyMap;
   m_verifyMap = 0;
   m_lenVerifyMap = 0;

   delete m_mapAction;
   m_mapAction = 0;
   m_lenMapAction = 0;

   delete m_verifyLot;
   m_verifyLot = 0;
   m_lenVerifyLot = 0;

   delete m_lotAction;
   m_lotAction = 0;
   m_lenLotAction = 0;

   delete m_scavenge;
   m_scavenge = 0;
   m_lenScavenge = 0;

   delete m_readChangeBuffer;
   m_readChangeBuffer = 0;

   EnterCriticalSection(&m_qLock);

   while (m_queueHead)
   {
      QueueEntry *old = m_queueHead;
      m_queueHead = m_queueHead->m_next;

      delete old;
   }

   m_queueTail = 0;

   LeaveCriticalSection(&m_qLock);

   if (logTraceLevel >= logTraceLevelDetail)
   {
      char localBuf[_MAX_PATH+51];

      sprintf(localBuf, "Current non-scavenged queue depth = %d",
            m_queueDepth);
      logTrace(funcName, localBuf);
      sprintf(localBuf, "Maximum non-scavenged queue depth = %d",
            m_maxQueueDepth);
      logTrace(funcName, localBuf);

      if (logTraceLevel >= logTraceLevelFiles)
      {
         sprintf(localBuf, "Items in map assocation hold buffer = %d",
               m_numberOnHold);
         logTrace(funcName, localBuf);

         if (m_numberOnHold > 0)
         {
            for (unsigned long i = 0; i < m_mapHold.length(); i++)
            {
               if (m_mapHold[i].inUse)
               {
                  sprintf(localBuf, "   Held association: %s",
                     m_mapHold[i].pathName);
                  logTrace(funcName, localBuf);
               }
            }
         }
         m_notifyCheck.dump(m_monitorId, m_dirPath);
         m_notifyCheck.clear();
      }
   }

   m_queueDepth = 0;
   m_maxQueueDepth = 0;

   delete m_monitorId;
   m_monitorId = 0;

   delete m_dirPath;
   m_dirPath = 0;

   m_mapHold.setCapacity(0);
}

bool MonitorShare::alive()
{
   static const char *funcName = "MonitorShare::alive()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   bool retval = false;

   EnterCriticalSection(&m_sLock);

   try
   {
      retval = m_initialized;
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
      retval = false;
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
      retval = false;
   }

   LeaveCriticalSection(&m_sLock);

   return retval;
}

bool MonitorShare::requestShutdown()
{
   static const char *funcName = "MonitorShare::requestShutdown()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   bool retval = false;

   EnterCriticalSection(&m_sLock);

   try
   {
      retval = m_requestShutdown;
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
      retval = false;
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
      retval = false;
   }

   LeaveCriticalSection(&m_sLock);

   return retval;
}

void MonitorShare::setRequestShutdown(bool b)
{
   static const char *funcName = "MonitorShare::setRequestShutdown()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   if (m_inShutdown == false)
   {
      EnterCriticalSection(&m_sLock);

      try
      {
         m_requestShutdown = b;
      }
      catch (_com_error &e)
      {
         logErrorMsg(e.Error(), funcName, e.ErrorMessage());
      }
      catch (...)
      {
         logErrorMsg(errSmmUnknownException, funcName, 0);
      }

      LeaveCriticalSection(&m_sLock);
   }
}

/*
 * this function parses the extensions and builds the extInfo tree
 * for monitoring.  if there are no extensions, then the defaults
 * are used.
 */

class LocalPointersParseExt
{
   public:
      SmmExtension *pSmmExt;

      LocalPointersParseExt() : pSmmExt(0) {}
      ~LocalPointersParseExt() { delete pSmmExt; }
};

bool MonitorShare::parseExtensions(const char **extensions,
         unsigned long numExtensions)
{
   static const char *funcName = "MonitorShare::parseExtensions()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   bool retval = false;
   bool useDefault = false;

   if (extensions == 0 || numExtensions == 0)
   {
      useDefault = true;
      numExtensions = 1;
   }

   try
   {
      for (unsigned long i = 0; i < numExtensions; i++)
      {
         const char *pExt = 0;

         if (useDefault)
            pExt = defaultExts;
         else
            pExt = extensions[i];

         if (pExt)
         {
            unsigned long ln = strtrimlen(pExt);

            if (logTraceLevel >= logTraceLevelFiles)
            {
               char localBuf[_MAX_PATH+1];

               sprintf(localBuf,"Parsing extension string '%s'", pExt);
               logTrace(funcName, localBuf);
            }

   // this is the tricky part and makes a big assumption that there will
   // not be a '>' or '<' in an extension (unless enclosed in [] or {})

            unsigned long idx = 0;
            unsigned long j = 0;

            char extBuf[_MAX_PATH+1];
            char lastModifier = 0;
            long lPos = -1;
            bool firstExt = true;
            LocalPointersParseExt LP;

            while (idx < ln)
            {
               unsigned long lnExt = 0;
               bool gotExt = false;

               long nSqrBracket = 0;
               long nCurlyBracket = 0;

               for (j = idx; j < ln; j++)
               {
                  if (pExt[j] == '<' || pExt[j] == '>')
                  {
                     if (nSqrBracket == 0 && nCurlyBracket == 0)
                     {
                        if (lPos >= 0)
                           lastModifier = pExt[lPos];

                        lPos = j;
                        lnExt = j - idx;

                        strncpy(extBuf,pExt+idx,lnExt);
                        extBuf[lnExt] = '\0';
                        gotExt = true;
                        break;
                     }
                  }
                  else if (pExt[j] == '[')
                  {
                     nSqrBracket++;
                  }
                  else if (pExt[j] == ']')
                  {
                     nSqrBracket--;
                  }
                  else if (pExt[j] == '{')
                  {
                     nCurlyBracket++;
                  }
                  else if (pExt[j] == '}')
                  {
                     nCurlyBracket--;
                  }
               }

               if (gotExt == false && ln > idx)
               {
                  if (lPos >= 0)
                     lastModifier = pExt[lPos];

                  lnExt = ln - idx;
                  strncpy(extBuf,pExt+idx,lnExt);
                  extBuf[lnExt] = '\0';
                  gotExt = true;
               }

               idx = j + 1;      // either way, this is the next 'idx'

               if (gotExt)
               {
                  if (firstExt)
                  {
                     LP.pSmmExt = new SmmExtension;
                     if (LP.pSmmExt == 0)
                     {
                        logErrorMsg(errSmmResource, funcName,
                           "Failed to allocate memory for SmmExtension.");
                        return false;
                     }

                     if (logTraceLevel >= logTraceLevelFiles)
                     {
                        char localBuf[_MAX_PATH+1];

                        sprintf(localBuf,"Monitoring extension '%s'", extBuf);
                        logTrace(funcName, localBuf);
                     }

                     if (LP.pSmmExt->set(extBuf, lnExt) != 0)
                        return false;

                     firstExt = false;
                  }
                  else if (lastModifier == '>')
                  {
                     if (logTraceLevel >= logTraceLevelFiles)
                     {
                        char localBuf[_MAX_PATH+1];

                        sprintf(localBuf,"Adding association '%s'", extBuf);
                        logTrace(funcName, localBuf);
                     }

                     if (LP.pSmmExt->addAssoc(extBuf, lnExt) != 0)
                        return false;
                  }
                  else if (lastModifier == '<')
                  {
                     if (logTraceLevel >= logTraceLevelFiles)
                     {
                        char localBuf[_MAX_PATH+1];

                        sprintf(localBuf,"Adding exclusion '%s'", extBuf);
                        logTrace(funcName, localBuf);
                     }

                     if (LP.pSmmExt->addExclude(extBuf, lnExt) != 0)
                        return false;
                  }
                  else      // should not happen
                  {
                     logErrorMsg(errSmmInternalError, funcName,
                        "Unknown extension modifier.");
                     return false;
                  }
               }
               else
               {
                  char localBuf[_MAX_PATH+50];

                  sprintf(localBuf,"Invalid extension string '%s'", pExt);
                  logErrorMsg(errSmmConfiguration, funcName, localBuf);
                  return false;
               }
            }

            if (LP.pSmmExt)
            {
               if (m_extInfo.append(LP.pSmmExt) != 0)
                  return false;

               LP.pSmmExt = 0;      // for dtor (m_extInfo has ownership)
            }
         }
      }

// now validate the regular expressions, if any

      {
         SmmExtension **pExt;
         SmmExtensionIterator it(&m_extInfo);

         while (retval == false && (pExt = it()))
         {
            if (*pExt)
            {
               if ((*pExt)->testRegExp() == false)
               {
                  char localBuf[256];

                  sprintf(localBuf,
                     "One or more invalid regular expressions found for %s.",
                     m_monitorId);
                  logErrorMsg(errSmmBadRegExp, 0, localBuf);
                  return false;
               }
            }
         }
      }

      retval = true;
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}

/*
 * this function takes a path name and checks against the extension
 * information to determine if the extension is being monitored.  it
 * also checks of there are required associations or exclusions
 */

bool MonitorShare::checkMap(const char *pathName, double timeStamp)
{
   static const char *funcName = "MonitorShare::checkMap()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   bool retval = false;

   try
   {
      bool noMonitor = true;

      SmmExtension **pExt;
      SmmExtensionIterator it(&m_extInfo);

      while (retval == false && (pExt = it()))
      {
         if (*pExt)
         {
            SmmExtension::MonitorAction fStatus = (*pExt)->find(pathName);
            if (fStatus == SmmExtension::ActionOK)
            {
               if (m_notifyCheck.isADupMap(pathName, timeStamp))
               {
                  if (logTraceLevel >= logTraceLevelFiles)
                  {
                     char localBuf[_MAX_PATH+50];

                     sprintf(localBuf,"File '%s' has already been processed.",
                           pathName);
                     logTrace(funcName, localBuf);
                  }
               }
               else
               {
                  if (logTraceLevel >= logTraceLevelDetail)
                  {
                     char localBuf[_MAX_PATH+50];

                     sprintf(localBuf,"File '%s' will be processed.", pathName);
                     logTrace(funcName, localBuf);
                  }

                  if (processMap(pathName, timeStamp) == Error_Failure)
                  {
                     // FIX: handle e-Server, other issues?
                  }
               }
               noMonitor = false;
               retval = true;
            }
            else if (fStatus == SmmExtension::ActionMissingAssoc)
            {
               if (logTraceLevel >= logTraceLevelDetail)
               {
                  char localBuf[_MAX_PATH+50];

                  sprintf(localBuf,"File '%s' is missing association(s).",
                              pathName);
                  logTrace(funcName, localBuf);
               }

               addToHold(pathName, timeStamp);
               noMonitor = false;
            }
            else if (fStatus == SmmExtension::ActionAssociation)
            {
               if (logTraceLevel >= logTraceLevelDetail)
               {
                  char localBuf[_MAX_PATH+100];

                  sprintf(localBuf,
                     "File '%s' is associated with a monitored extension.",
                     pathName);
                  logTrace(funcName, localBuf);
               }

               noMonitor = false;

               long slot = findBaseInHold(pathName);
               if (slot >= 0)
               {
                  if ((*pExt)->find(m_mapHold[slot].pathName) ==
                           SmmExtension::ActionOK)
                  {
                     if (m_notifyCheck.isADupMap(m_mapHold[slot].pathName,
                           m_mapHold[slot].timeStamp))
                     {
                        if (logTraceLevel >= logTraceLevelFiles)
                        {
                           char localBuf[_MAX_PATH+50];

                           sprintf(localBuf,
                          "File '%s' (from hold) has already been processed.",
                                 m_mapHold[slot].pathName);
                           logTrace(funcName, localBuf);
                        }
                     }
                     else
                     {
                        if (logTraceLevel >= logTraceLevelDetail)
                        {
                           char localBuf[_MAX_PATH+50];

                           sprintf(localBuf,
                              "File '%s' (from hold) will be processed.",
                              m_mapHold[slot].pathName);
                           logTrace(funcName, localBuf);
                        }

                        if (processMap(m_mapHold[slot].pathName,
                              m_mapHold[slot].timeStamp) == Error_Failure)
                        {
                           // FIX: handle e-Server, other issues?
                        }
                     }

                     removeFromHold(slot);
                     retval = true;
                  }
               }
            }
            else if (fStatus == SmmExtension::ActionExclusion)
            {
               if (logTraceLevel >= logTraceLevelFiles)
               {
                  char localBuf[_MAX_PATH+100];

                  sprintf(localBuf,
                     "File '%s' skipped due an extension exclusion.",
                     pathName);
                  logTrace(funcName, localBuf);
               }

               long slot = findBaseInHold(pathName);
               while (slot >= 0)
               {
                  if (logTraceLevel >= logTraceLevelFiles)
                  {
                     char localBuf[_MAX_PATH+50];

                     sprintf(localBuf,
                        "File '%s' (from hold) skipped due to exclusion.",
                        m_mapHold[slot].pathName);
                     logTrace(funcName, localBuf);
                  }

                  removeFromHold(slot);
                  slot = findBaseInHold(pathName);
               }
            }
         }
      }

      if (noMonitor == true && logTraceLevel >= logTraceLevelFiles)
      {
         char localBuf[_MAX_PATH+50];

         sprintf(localBuf,"File '%s' will be ignored",pathName);
         logTrace(funcName, localBuf);
      }
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
      retval = false;
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
      retval = false;
   }

   return retval;
}

bool MonitorShare::checkLot(const char *pathName, double timeStamp)
{
   static const char *funcName = "MonitorShare::checkLot()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   bool retval = false;

   try
   {
      bool noMonitor = true;

// the full path is passed in but the comparison is done only against
// the folder name so peel it off for the comparison

      char lotName[_MAX_PATH+1];

      long ls = strlen(pathName);
      long idx = -1;
      long lnLot = 0;
      long i;

      for (i = ls-1; i > 0; i--)
      {
         if (pathName[i] == '\\' || pathName[i] == '/')
         {
            idx = i;
            break;
         }
      }

      if (idx >= 0)
      {
         for (i = idx+1; i < ls; i++)
            lotName[lnLot++] = pathName[i];

         lotName[lnLot] = '\0';
      }

      if (lnLot > 0 && m_lotDir.compare(lotName))
      {
         if (logTraceLevel >= logTraceLevelFiles)
         {
            char localBuf[_MAX_PATH+50];

            sprintf(localBuf,"Lot folder match '%s'", m_lotDir.value());
            logTrace(funcName, localBuf);
         }

         if (m_notifyCheck.isADupLot(pathName))
         {
            if (logTraceLevel >= logTraceLevelFiles)
            {
               char localBuf[_MAX_PATH+50];

               sprintf(localBuf,"Folder '%s' has already been processed.",
                     pathName);
               logTrace(funcName, localBuf);
            }
         }
         else
         {
            if (logTraceLevel >= logTraceLevelDetail)
            {
               char localBuf[_MAX_PATH+50];

               sprintf(localBuf,"Folder '%s' will be processed.", pathName);
               logTrace(funcName, localBuf);
            }

            if (processLot(pathName) == Error_Failure)
            {
               // FIX: handle e-Server, other issues?
            }
         }
         noMonitor = false;
         retval = true;
      }

      if (noMonitor == true && logTraceLevel >= logTraceLevelDetail)
      {
         char localBuf[_MAX_PATH+50];

         sprintf(localBuf,"Folder '%s' will be ignored",pathName);
         logTrace(funcName, localBuf);
      }
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
      retval = false;
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
      retval = false;
   }

   return retval;
}

class LocalPointersMapAction
{
   public:
      _clsMapActions *iMapActions;

      LocalPointersMapAction() : iMapActions(0) {}
      ~LocalPointersMapAction() { if (iMapActions) iMapActions->Release(); }
};

// this method is the main processing when a file has been identified as
// a map by extension, etc.  the verifyMap proc, if defined, and the
// mapAction proc, if defined, are referenced.  if a mapAction proc
// isn't defined, then the default map action object will be utilized.

long MonitorShare::processMap(const char *pathName, double timeStamp)
{
   static const char *funcName = "MonitorShare::processMap()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   try
   {
      // first, check if it has been recently processed.  this is necessary
      // because the file notification event is sent multiple times when
      // a file is dropped or written to the folder

      bool verifyOK = true;

      if (m_verifyMap && m_lenVerifyMap > 0)
      {
         eServer statitLib(m_useServerPool);

         if (statitLib.init() == Error_Failure)
            return retval;

         if (statitLib.putGlobalString("%SMM_MapPath", pathName) ==
               Error_Failure)
            return retval;

         if (statitLib.putGlobalString("%SMM_MonitorId", m_monitorId) ==
               Error_Failure)
            return retval;

         if (statitLib.putGlobalString("%SMM_VerifyMap", m_verifyMap) ==
                  Error_Failure)
            return retval;

         if (statitLib.putGlobalInt("$SMM_VerifyMap", 0) == Error_Failure)
            return retval;

         if (logTraceLevel >= logTraceLevelDetail)
         {
            char localBuf[_MAX_PATH*2+100];

            sprintf(localBuf,"Executing verify map '%s' on '%s'.", m_verifyMap,
                  pathName);
            logTrace(funcName, localBuf);
         }

         if (statitLib.runCmd("smm_runVerifyMap") == Error_Failure)
            return retval;

         long smm_verifyMap = 0;
         if (statitLib.getGlobalInt("$SMM_VerifyMap", smm_verifyMap) ==
                  Error_Failure)
            return retval;

         if (smm_verifyMap == 0)
            verifyOK = false;

         if (logTraceLevel >= logTraceLevelDetail)
         {
            char localBuf[_MAX_PATH+100];

            sprintf(localBuf,"Verify map on '%s' returned $SMM_VerifyMap = %d.",
                  pathName, smm_verifyMap);
            logTrace(funcName, localBuf);
         }
      }

      if (verifyOK)
      {
         if (m_mapAction && m_lenMapAction > 0)
         {
            eServer statitLib(m_useServerPool);

            if (statitLib.init() == Error_Failure)
               return retval;

            if (statitLib.putGlobalString("%SMM_MapPath", pathName) ==
                     Error_Failure)
               return retval;

            if (statitLib.putGlobalString("%SMM_MonitorId", m_monitorId) ==
                  Error_Failure)
               return retval;

            if (statitLib.putGlobalInt("$SMM_MapAction", 1) == Error_Failure)
               return retval;

            if (statitLib.putGlobalString("%SMM_MapAction", m_mapAction) ==
                     Error_Failure)
               return retval;

            if (statitLib.runCmd("smm_runMapAction") == Error_Failure)
               return retval;

            m_notifyCheck.addMap(pathName, timeStamp);
         }
         else
         {
            LocalPointersMapAction LP;

            long aoStatus = Error_Failure;
            unsigned long tryCount = 0;

            do
            {
               if (LP.iMapActions)
               {
                  LP.iMapActions->Release();
                  LP.iMapActions = 0;
               }

               if (tryCount > 0)               // very brief delay
                  ::Sleep(retryDelay);

               HRESULT hr = CoCreateInstance(m_clsIdMapAction, 0, CLSCTX_SERVER,
                     __uuidof(_clsMapActions), (void **) &LP.iMapActions);
               if (FAILED(hr))
               {
                  if (tryCount >= maxRetryCount)
                  {
                     char errBuf[BUFSIZ+1];
                     char localBuf[BUFSIZ+256];

                     if (! FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, hr, 0,
                           errBuf, BUFSIZ, 0))
                     {
                        sprintf(localBuf,
                           "Failed to create an instance of SMMapActions.clsMapActions. Error code=%x.",
                           hr);
                     }
                     else
                     {
                        sprintf(localBuf,
                           "Failed to create an instance of SMMapActions.clsMapActions (%s). Error code=%x.",
                           errBuf, hr);
                     }
                     logErrorMsg(errSmmConfiguration, funcName, localBuf);
                  }
               }
               else
               {
                  long rStatus = 0;   // returned status, will be ignored
                  _bstr_t bPath(pathName);
                  _bstr_t bId(m_monitorId);

                  hr = LP.iMapActions->AddAOActions(bPath, bId, _bstr_t(""),
                              _bstr_t("map"), &rStatus);
                  if (FAILED(hr))
                  {
                     if (tryCount >= maxRetryCount)
                     {
                        char localBuf[256];

                        sprintf(localBuf,
                           "SMMapActions.AddAOActions() failed. Error code=%x.", hr);
                        logErrorMsg(errSmmMapAction, funcName, localBuf);
                     }
                  }
                  else
                  {
                     m_notifyCheck.addMap(pathName, timeStamp);
                     aoStatus = Error_Success;
                  }
               }
               tryCount++;
            }
            while (aoStatus == Error_Failure && tryCount <= maxRetryCount);

            if (aoStatus == Error_Failure)
               return retval;
         }
      }

      retval = Error_Success;
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}

// this method is the main processing when a folder has been identified as
// a lot.  the verifyLot proc, if defined, and the lotAction proc, if defined,
// are referenced.  if a lotAction proc isn't defined, then the default lot
// action object will be utilized.

long MonitorShare::processLot(const char *pathName)
{
   static const char *funcName = "MonitorShare::processLot()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   try
   {
      // first, check if it has been recently processed.  this is necessary
      // because the file notification event is sent multiple times when
      // a file is dropped or written to the folder

      bool verifyOK = true;

      if (m_verifyLot && m_lenVerifyLot > 0)
      {
         eServer statitLib(m_useServerPool);

         if (statitLib.init() == Error_Failure)
            return retval;

         if (statitLib.putGlobalString("%SMM_LotPath", pathName) ==
               Error_Failure)
            return retval;

         if (statitLib.putGlobalString("%SMM_MonitorId", m_monitorId) ==
               Error_Failure)
            return retval;

         if (statitLib.putGlobalString("%SMM_VerifyLot", m_verifyLot) ==
                  Error_Failure)
            return retval;

         if (statitLib.putGlobalInt("$SMM_VerifyLot", 0) == Error_Failure)
            return retval;

         if (logTraceLevel >= logTraceLevelDetail)
         {
            char localBuf[_MAX_PATH*2+100];

            sprintf(localBuf,"Executing verify lot '%s' on '%s'.", m_verifyLot,
                  pathName);
            logTrace(funcName, localBuf);
         }

         if (statitLib.runCmd("smm_runVerifyLot") == Error_Failure)
            return retval;

         long smm_verifyLot = 0;
         if (statitLib.getGlobalInt("$SMM_VerifyLot", smm_verifyLot) ==
                  Error_Failure)
            return retval;

         if (smm_verifyLot == 0)
            verifyOK = false;

         if (logTraceLevel >= logTraceLevelDetail)
         {
            char localBuf[_MAX_PATH+100];

            sprintf(localBuf,"Verify lot on '%s' returned $SMM_VerifyLot = %d.",
                  pathName, smm_verifyLot);
            logTrace(funcName, localBuf);
         }
      }

      if (verifyOK)
      {
         if (m_lotAction && m_lenLotAction > 0)
         {
            eServer statitLib(m_useServerPool);

            if (statitLib.init() == Error_Failure)
               return retval;

            if (statitLib.putGlobalString("%SMM_LotPath", pathName) ==
                     Error_Failure)
               return retval;

            if (statitLib.putGlobalString("%SMM_MonitorId", m_monitorId) ==
                  Error_Failure)
               return retval;

            if (statitLib.putGlobalInt("$SMM_LotAction", 1) == Error_Failure)
               return retval;

            if (statitLib.putGlobalString("%SMM_LotAction", m_lotAction) ==
                     Error_Failure)
               return retval;

            if (statitLib.runCmd("smm_runLotAction") == Error_Failure)
               return retval;

            m_notifyCheck.addLot(pathName);
         }
         else
         {
            LocalPointersMapAction LP;

            HRESULT hr = CoCreateInstance(m_clsIdMapAction, 0, CLSCTX_SERVER,
                     __uuidof(_clsMapActions), (void **) &LP.iMapActions);
            if (FAILED(hr))
            {
               char errBuf[BUFSIZ+1];
               char localBuf[BUFSIZ+256];

               if (! FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, hr, 0,
                     errBuf, BUFSIZ, 0))
               {
                  sprintf(localBuf,
                     "Failed to create an instance of SMMapActions.clsMapActions. Error code=%x.",
                     hr);
               }
               else
               {
                  sprintf(localBuf,
                     "Failed to create an instance of SMMapActions.clsMapActions (%s). Error code=%x.",
                     errBuf, hr);
               }
               logErrorMsg(errSmmConfiguration, funcName, localBuf);
               return retval;
            }

            long rStatus = 0;   // returned status, will be ignored
            _bstr_t bPath(pathName);
            _bstr_t bId(m_monitorId);

            hr = LP.iMapActions->AddAOActions(bPath, bId, _bstr_t(""),
                        _bstr_t("lot"), &rStatus);
            if (FAILED(hr))
            {
               char localBuf[256];

               sprintf(localBuf,
                  "SMMapActions.AddAOActions() failed. Error code=%x.", hr);
               logErrorMsg(errSmmMapAction, funcName, localBuf);
               return retval;
            }

            m_notifyCheck.addLot(pathName);
         }
      }

      retval = Error_Success;
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}

// these next functions maintain a simple queue of maps that do not
// have the associated file(s) available so they are not harvested.
// yes, everything is linear but the plan is that the length should
// be quite small.  if not, changes will be made.

long MonitorShare::findInHold(const char *pathName)
{
   static const char *funcName = "MonitorShare::findInHold()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long slot = -1;

   if (m_numberOnHold > 0)
   {
      for (unsigned long i = 0; i < m_mapHold.length(); i++)
      {
         if (m_mapHold[i].inUse)
         {
            if (stricmp(m_mapHold[i].pathName, pathName) == 0)
            {
               slot = i;
               break;
            }
         }
      }
   }

   return slot;
}

long MonitorShare::findBaseInHold(const char *pathName)
{
   static const char *funcName = "MonitorShare::findBaseInHold()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long slot = -1;

   if (m_numberOnHold > 0)
   {
      unsigned long i;
      long idx = -1;
      unsigned long ls = strlen(pathName);
      for (i = ls-1; i >= 0; i--)
      {
         if (pathName[i] == '.')
         {
            idx = i;
            break;
         }
      }

      if (idx > 0)
      {
         for (unsigned long i = 0; i < m_mapHold.length(); i++)
         {
            if (m_mapHold[i].inUse)
            {
               if (strnicmp(m_mapHold[i].pathName, pathName, idx) == 0)
               {
                  slot = i;
                  break;
               }
            }
         }
      }
   }

   return slot;
}

long MonitorShare::addToHold(const char *pathName, double timeStamp)
{
   static const char *funcName = "MonitorShare::addToHold()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   long openSlot = findInHold(pathName);
   if (openSlot < 0)
   {
      for (unsigned long i = 0; i < m_mapHold.length(); i++)
      {
         if (m_mapHold[i].inUse == false)
         {
            openSlot = (long) i;
            break;
         }
      }

      if (openSlot < 0)
      {
         openSlot = m_mapHold.length();
         if (m_mapHold.setLength(openSlot+1) < 0)
         {
            logErrorMsg(errSmmResource, funcName, "Map queue allocation");
            openSlot = -1;
         }
      }

      if (openSlot >= 0)
      {
         unsigned long ls = strlen(pathName);
         char *p = new char[ls+1];
         if (p == 0)
         {
            logErrorMsg(errSmmResource, funcName, "Map queue entry allocation");
         }
         else
         {
            strcpy(p, pathName);
            m_mapHold[openSlot].pathName = p;
            m_mapHold[openSlot].timeStamp = timeStamp;
            m_mapHold[openSlot].inUse = true;
            m_numberOnHold++;
            retval = openSlot;
         }
      }
   }
   else
   {
      retval = openSlot;
   }

   return retval;
}

long MonitorShare::removeFromHold(const char *pathName)
{
   static const char *funcName = "MonitorShare::removeFromHold(const char *)";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   long slot = findInHold(pathName);
   if (slot >= 0)
      retval = removeFromHold(slot);
   else                        // not in the queue
      retval = Error_Success;

   return retval;
}

long MonitorShare::removeFromHold(long slot)
{
   static const char *funcName = "MonitorShare::removeFromHold(long)";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (slot >= 0 && slot < (long) m_mapHold.length())
   {
      if (m_mapHold[slot].inUse)
      {
         delete m_mapHold[slot].pathName;
         m_mapHold[slot].pathName = 0;
         m_mapHold[slot].timeStamp = 0.0;
         m_mapHold[slot].inUse = false;
         m_numberOnHold--;
      }

      retval = Error_Success;
   }

   return retval;
}

long MonitorShare::addToQueue(const char *pathName, DWORD action, bool
         fromScavenge)
{
   static const char *funcName = "MonitorShare::addToQueue()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;
   bool didAdd = false;
   bool haveLock = false;

   try
   {
      QueueEntry *qe = new QueueEntry(pathName, action, fromScavenge);
      if (qe)
      {
         EnterCriticalSection(&m_qLock);
         haveLock = true;

         if (m_queueTail == 0)
         {
            m_queueHead = m_queueTail = qe;
         }
         else
         {
            m_queueTail->m_next = qe;
            m_queueTail = m_queueTail->m_next;
         }

         if (fromScavenge == false)
         {
            m_queueDepth++;
            if (m_queueDepth > m_maxQueueDepth)
               m_maxQueueDepth = m_queueDepth;

            if (m_scavengeSurgeExceeded == false &&
                     m_scavengeSurgeTrigger > 0 &&
                     m_queueDepth >= m_scavengeSurgeTrigger)
               m_scavengeSurgeExceeded = true;
         }

         didAdd = true;
      }
      else
      {
         logErrorMsg(errSmmResource, funcName,
            "Failed to allocate memory for a queue entry.");
      }
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   if (haveLock)
   {
      LeaveCriticalSection(&m_qLock);
      haveLock = false;
   }


   if (didAdd)
   {
      if (ReleaseSemaphore(m_queueSemaphore, 1, 0))
      {
         retval = Error_Success;
      }
      else
      {
         char localBuf[128];

         DWORD eCode = GetLastError();
         sprintf(localBuf,"ReleaseSemaphore() returned %ld.",eCode);
         logErrorMsg(errSmmResource, funcName, localBuf);
      }
   }

   return retval;
}

class LocalPointersRunQueue
{
   public:
      QueueEntry *qe;

      LocalPointersRunQueue() : qe(0) {}
      ~LocalPointersRunQueue() { delete qe; }
};

void MonitorShare::runQueue()
{
   static const char *funcName = "MonitorShare::runQueue()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   m_initialized = true;
   SetEvent(m_queueEvent);

   if (logTraceLevel >= logTraceLevelDetail)
   {
      char localBuf[_MAX_PATH+151];

      sprintf(localBuf, "Queue thread started for %s '%s'", m_monitorId,
            m_dirPath);
      logTrace(funcName, localBuf);
   }

   bool haveLock = false;
   bool done = false;

   while (! done)
   {
      DWORD waitResult = MsgWaitForMultipleObjectsEx(1, &m_queueSemaphore,
               INFINITE, QS_ALLINPUT, MWMO_ALERTABLE);
      if (waitResult == WAIT_OBJECT_0)
      {
         if (m_inShutdown || _Module.inShutdown())
         {
            done = true;
         }
         else
         {
            EnterCriticalSection(&m_qLock);
            haveLock = true;

            try
            {
               LocalPointersRunQueue LP;

               if (m_queueHead)
               {
                  LP.qe = m_queueHead;
                  m_queueHead = m_queueHead->m_next;
                  if (m_queueHead == 0)
                     m_queueTail = 0;

                  if (LP.qe->fromScavenge() == false)
                  {
                     m_queueDepth--;
                     if (m_queueDepth == 0 && m_scavengeSurgeExceeded)
                        m_triggerScavenge = true;
                  }

                  LeaveCriticalSection(&m_qLock);
                  haveLock = false;

// some processes (tools) create temporary files that are found by the
// monitor but are then deleted before additional processing.  a second
// try is made to get the attributes and if both return that the file
// or path is not found it will be ignored.

                  long fileAttr = GetFileAttributes(LP.qe->pathName());
                  if (fileAttr == (-1))
                  {
                     DWORD eCode = GetLastError();
                     if (eCode == ERROR_FILE_NOT_FOUND ||
                         eCode == ERROR_PATH_NOT_FOUND)      // try again
                     {
                        fileAttr = GetFileAttributes(LP.qe->pathName());
                        if (fileAttr == (-1))
                        {
                           eCode = GetLastError();
                           if (! (eCode == ERROR_FILE_NOT_FOUND ||
                                  eCode == ERROR_PATH_NOT_FOUND)) // report this
                           {
                              char localBuf[_MAX_PATH+100];

                              sprintf(localBuf,
                                    "GetFileAttributes('%s') returned %ld.",
                                    LP.qe->pathName(), eCode);
                              logErrorMsg(errSmmWarning, funcName, localBuf);
                           }
                        }
                     }
                  }

                  if (fileAttr != (-1))
                  {
                     if (fileAttr & FILE_ATTRIBUTE_DIRECTORY)  // directory
                     {
                        if (LP.qe->action() == FILE_ACTION_ADDED ||
                            LP.qe->action() == FILE_ACTION_RENAMED_NEW_NAME)
                        {
                           if (checkLot(LP.qe->pathName(),LP.qe->timeStamp()))
                           {      // FIX?
                           }
                        }
                     }
                     else
                     {
                        if (checkMap(LP.qe->pathName(),LP.qe->timeStamp()))
                        {         // FIX?
                        }
                     }
                  }

                  delete LP.qe;
                  LP.qe = 0;      // for dtor
               }
            }
            catch (_com_error &e)
            {
               if (haveLock)
                  LeaveCriticalSection(&m_qLock);

               logErrorMsg(e.Error(), funcName, e.ErrorMessage());
               setRequestShutdown(true);
            }
            catch (...)
            {
               if (haveLock)
                  LeaveCriticalSection(&m_qLock);

               logErrorMsg(errSmmUnknownException, funcName, 0);
               setRequestShutdown(true);
            }
         }
      }
      else if (waitResult == WAIT_OBJECT_0 + 1)
      {
         MSG msg;

         while ((! done) && PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
         {
            if (logTraceLevel >= logTraceLevelDetail)
            {
               char localBuf[512];

               sprintf(localBuf,"Message %lu received for %s", msg.message,
                     m_monitorId);
               logTrace(funcName, localBuf);
            }

            if (msg.message == WM_QUIT)
            {
               done = true;
            }
            else
            {
               DispatchMessage(&msg);
            }
         }
      }
      else
      {
         if (logTraceLevel)
         {
            char localBuf[128];

            DWORD eCode = GetLastError();
            sprintf(localBuf,"Unexpected queue semaphore wait result = %lu (%ld)",
                  waitResult, eCode);
            logTrace(funcName, localBuf);

            setRequestShutdown(true);
         }
      }
   }

   if (logTraceLevel >= logTraceLevelDetail)
   {
      char localBuf[_MAX_PATH+151];

      sprintf(localBuf, "Queue thread completed for %s '%s'", m_monitorId,
            m_dirPath);
      logTrace(funcName, localBuf);
   }
}


SmmString::SmmString()
{
   m_str = 0;
   m_regexp = false;
}

SmmString::SmmString(const char *s)
{
   set(s);
}

SmmString::SmmString (const SmmString& rhs)
{
   clear();

   if (rhs.m_str)
   {
      unsigned long ls = strlen(rhs.m_str);
      m_str = new char[ls+1];
      if (m_str)
         strcpy(m_str,rhs.m_str);
   }
   else
   {
      m_str = 0;
   }

   setRegExpFlag();
}

SmmString::~SmmString()
{
   clear();
}

void SmmString::clear()
{
   delete m_str;

   m_str = 0;
   m_regexp = false;
}

int SmmString::set (const char * s)
{
   char * sd = 0;
   int retval = 0;

   clear();

   if (s)
   {
      unsigned long ls = strlen(s);
      sd = new char[ls+1];
      if (sd)
         strcpy(sd,s);
      else
         retval = 1;
   }

   m_str = sd;

   setRegExpFlag();
   return retval;
}


int SmmString::set (const char * s, long ln)
{
   clear();
   int retval = 0;

   if (s)
   {
      char * s2 = 0;

      if (ln <= 0)
         ln = strlen(s);

      s2 = new char[ln+1];
      if (s2)
      {
         strncpy(s2, s, ln);
         s2[ln] = '\0';

         m_str = s2;

         setRegExpFlag();
      }
      else
      {
         retval = 1;
      }
   }
   return retval;
}

void SmmString::setRegExpFlag()
{
   if (m_str)
   {
      m_regexp = false;         // assume no regular expression
      unsigned long ln = strlen(m_str);
      for (unsigned long i = 0; i < ln; i++)
      {
         if (! ((m_str[i] >= 'A' && m_str[i] <= 'Z') ||
                (m_str[i] >= 'a' && m_str[i] <= 'z') ||
                (m_str[i] >= '0' && m_str[i] <= '9')))
         {
            m_regexp = true;
            break;
         }
      }
   }
}

bool SmmString::compare(const char *s)
{
   bool retval = false;

   if (m_str)
   {
      if (m_regexp)
      {
         SMRegExp re;

         if (re.init() == Error_Success)
         {
            if (re.setPattern(m_str) == Error_Success)
            {
               if (re.match(s) > 0)
                  retval = true;
            }
         }
      }
      else
      {
         retval = (stricmp(m_str, s) == 0);
      }
   }

   return retval;
}

// the following was added to test the regular expression so that problems
// with users specifying the strings can be determined at startup rather
// than everytime the string is tested.

bool SmmString::testRegExp()
{
   bool retval = false;

   if (m_str)
   {
      if (m_regexp)
      {
         SMRegExp re;

         if (re.init() == Error_Success)
         {
            if (re.setPattern(m_str) == Error_Success)
            {
               if (re.match("ABCDEF") >= 0)   // as long as it doesn't fail
                  retval = true;
            }
         }
      }
      else
      {
         retval = true;   // return ok if not a regular expression
      }
   }

   return retval;
}


int SmmExtNameVec :: append (const char * s)
{
   register int retVal = setLength (length() + 1);
   if (retVal == 0)
   {
      retVal = (*this)[length()-1].set (s);
      if (retVal != 0)
         setLength (length() - 1);
      else
         (*this)[length()-1].clearRegExpFlag();

   }
   return retVal;
}


int SmmExtNameVec :: append (const char * s, long ln)
{
   register int retVal = setLength (length() + 1);
   if (retVal == 0)
   {
      retVal = (*this)[length()-1].set (s, ln);
      if (retVal != 0)
         setLength (length() - 1);
      else
         (*this)[length()-1].clearRegExpFlag();
   }
   return retVal;
}


void SmmExtNameVec :: allocateProc (size_t first, size_t last)
{
   do
   {
      SmmString * p = (SmmString *) nodeFast (first);
      if (p != 0)
      {
         p->m_str = 0;
         p->m_regexp = false;
      }
   } while (++first <= last);
}

void SmmExtNameVec :: deallocateProc (size_t i, size_t last)
{
   do
   {
      SmmString * p = (SmmString *) nodeFast (i);
      if (p == 0)
         break;
      p->SmmString::~SmmString();
   } while (++i <= last);
}


SmmExtension::SmmExtension()
{
}

SmmExtension::~SmmExtension()
{
}

bool SmmExtension::testRegExp()
{
   static const char *funcName = "SmmExtension::testRegExp()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   bool retval = true;

   if (m_ext.testRegExp() == false)
   {
      logErrorMsg(errSmmBadRegExp, 0, m_ext.value());
      retval = false;
   }

// check associations

   {
      SmmString *extName = 0;
      SmmStringIterator it(&m_assocExt);

      while (extName = it())
      {
         if (extName->testRegExp() == false)
         {
            logErrorMsg(errSmmBadRegExp, 0, extName->value());
            retval = false;
         }
      }
   }

// check exclusions

   {
      SmmString *extName = 0;
      SmmStringIterator it(&m_excludeExt);

      while (extName = it())
      {
         if (extName->testRegExp() == false)
         {
            logErrorMsg(errSmmBadRegExp, 0, extName->value());
            retval = false;
         }
      }
   }

   return retval;
}

/*
 * return values:  0 = do not monitor
 *                 1 = passes assocations and exclusions
 *                 2 = missing association
 *                 3 = association for monitored extension
 */

SmmExtension::MonitorAction SmmExtension::find(const char *pathName)
{
   static const char *funcName = "SmmExtension::find()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long i, ln;
   char extBuf[_MAX_PATH+1];
   char baseName[_MAX_PATH+1];

   MonitorAction retval = ActionNone;

   try
   {
      // first, parse out the extension

      unsigned long lnExt = 0;
      unsigned long ls = strlen(pathName);
      long idx = -1;

      for (i = ls-1; i >= 0; i--)
      {
         if (pathName[i] == '.')
         {
            idx = i;
            break;
         }
      }

      if (idx > 0)
         ln = ls - idx - 1;

      if (ln < 0)
         ln = 0;

      if (ln > 0)
      {
         strcpy(extBuf, &pathName[idx+1]);
         strncpy(baseName, pathName, idx+1);
         baseName[idx+1] = '\0';
      }
      else
      {
         extBuf[0] = '\0';
         strcpy(baseName, pathName);
      }

   // a file without an extension should still be matched with a .* regular
   // expression

      if (logTraceLevel >= logTraceLevelFiles)
      {
         char localBuf[_MAX_PATH+50];

         sprintf(localBuf,"Examining file '%s'", pathName);
         logTrace(funcName, localBuf);
      }

   // next, check against the extension

      if (m_ext.compare(extBuf))
      {
         if (logTraceLevel >= logTraceLevelFiles)
         {
            char localBuf[_MAX_PATH+50];

            sprintf(localBuf,"Extension match '%s'", extBuf);
            logTrace(funcName, localBuf);
         }

         retval = ActionOK;  // set to monitor b4 checking assoc & exclusions

         {
            // check associations

            SmmString *extName = 0;
            SmmStringIterator it(&m_assocExt);

            while (retval == ActionOK && (extName = it()))
            {
               char assocBuf[_MAX_PATH+1];

               strcpy(assocBuf, baseName);
               strcat(assocBuf, extName->value());

               long fileAttr = GetFileAttributes(assocBuf);
               if (fileAttr == (-1))
               {
                  retval = ActionMissingAssoc;
               }
               else
               {
                  if (fileAttr & FILE_ATTRIBUTE_DIRECTORY)   // no dirs allowed
                     retval = ActionMissingAssoc;
               }

               if (logTraceLevel >= logTraceLevelFiles)
               {
                  char localBuf[_MAX_PATH+50];

                  if (retval == ActionOK)
                     sprintf(localBuf,"Association match '%s'", assocBuf);
                  else
                     sprintf(localBuf,"Association not found '%s'", assocBuf);
                  logTrace(funcName, localBuf);
               }
            }
         }

         if (retval == ActionOK)         // check exclusions
         {
            SmmString *extName = 0;
            SmmStringIterator it(&m_excludeExt);

            while (retval == ActionOK && (extName = it()))
            {
               char excludeBuf[_MAX_PATH+1];

               strcpy(excludeBuf, baseName);
               strcat(excludeBuf, extName->value());

               long fileAttr = GetFileAttributes(excludeBuf);
               if (fileAttr != (-1))
               {
                  if (! (fileAttr & FILE_ATTRIBUTE_DIRECTORY))
                     retval = ActionExclusion;          // exclusion found
               }

               if (logTraceLevel >= logTraceLevelFiles)
               {
                  char localBuf[_MAX_PATH+50];

                  if (retval == ActionOK)
                     sprintf(localBuf,"Exclusion not found '%s'", excludeBuf);
                  else
                     sprintf(localBuf,"Exclusion match '%s'", excludeBuf);
                  logTrace(funcName, localBuf);
               }
            }
         }
      }
      else         // see if it is an associated extension
      {
         if (logTraceLevel >= logTraceLevelFiles)
         {
            char localBuf[_MAX_PATH+50];

            sprintf(localBuf,"Extension did not match '%s'", extBuf);
            logTrace(funcName, localBuf);
         }

         SmmString *extName = 0;
         SmmStringIterator it(&m_assocExt);

         while (retval == ActionNone && (extName = it()))
         {
            if (extName->compare(extBuf))
            {
               retval = ActionAssociation;

               if (logTraceLevel >= logTraceLevelFiles)
               {
                  char localBuf[_MAX_PATH+50];

                  sprintf(localBuf,"Association match '%s'", extName->value());
                  logTrace(funcName, localBuf);
               }
            }
         }
      }
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
      retval = ActionNone;
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
      retval = ActionNone;
   }

   return retval;
}

int SmmExtension::set(const char *s, long ln)
{
   m_assocExt.setLength(0);
   m_excludeExt.setLength(0);

   return m_ext.set(s, ln);
}

int SmmExtension::addAssoc(const char *s, long ln)
{
   return m_assocExt.append(s, ln);
}

int SmmExtension::addExclude(const char *s, long ln)
{
   return m_excludeExt.append(s, ln);
}

void SmmExtensionVec :: allocateProc (size_t first, size_t last)
{
   do
   {
      SmmExtension ** p = (SmmExtension **) nodeFast (first);
      if (p && *p)
      {
         *p = 0;
      }
   } while (++first <= last);
}

void SmmExtensionVec :: deallocateProc (size_t i, size_t last)
{
   do
   {
      SmmExtension ** p = (SmmExtension **) nodeFast (i);
      if (p == 0 || *p == 0)
         break;

      delete *p;
   } while (++i <= last);
}

void FileNotifyHistory::allocateProc(size_t first, size_t last)
{
   do
   {
      FileNotifyInfo *p = (FileNotifyInfo *) nodeFast(first);

      p->pathName[0] = '\0';
      p->timeStamp = 0.0;
   }
   while (++first <= last);
}

void FileNotifyHistory::deallocateProc(size_t first, size_t last)
{
   do
   {
      FileNotifyInfo *p = (FileNotifyInfo *) nodeFast(first);
      if (p)
      {
         p->pathName[0] = '\0';
         p->timeStamp = 0.0;
      }
   }
   while (++first <= last);
}

FileNotifyCheck::FileNotifyCheck()
{
   m_next = 0;
   m_nItems = 0;
   m_ageLimit = 0.0;

   m_lotNext = 0;
   m_lotNItems = 0;
   m_didInit = false;
}

FileNotifyCheck::~FileNotifyCheck()
{
   m_fileHistory.setCapacity(0);
   m_lotHistory.setCapacity(0);
}

void FileNotifyCheck::dump(const char *id, const char *dirPath)
{
   static const char *funcName = "FileNotifyCheck::dump()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   if (m_didInit)
   {
      char localBuf[_MAX_PATH+101];

      if (id && dirPath)
      {
         sprintf(localBuf, "Id: %s '%s', maps recently processed = %d",
            id, dirPath, m_nItems);
         logTrace(funcName, localBuf);
      }

      if (m_nItems > 0)
      {
         long nLoc = m_next - 1;
         SYSTEMTIME sTime;

         for (unsigned long i = 0; i < m_nItems; i++)
         {
            if (nLoc < 0)
               nLoc = m_fileHistory.length() - 1;

            VariantTimeToSystemTime(m_fileHistory[nLoc].timeStamp, &sTime);

            sprintf(localBuf, "   %s @ %.4d/%-.2d/%-.2d %.2d:%.2d:%.2d.%.3d",
               m_fileHistory[nLoc].pathName, sTime.wYear, sTime.wMonth,
               sTime.wDay, sTime.wHour, sTime.wMinute, sTime.wSecond,
               sTime.wMilliseconds);
            logTrace(funcName, localBuf);

            nLoc--;
         }
      }

      if (id && dirPath)
      {
         sprintf(localBuf, "Id: %s '%s', lots recently processed = %d",
            id, dirPath, m_lotNItems);
         logTrace(funcName, localBuf);
      }

      if (m_lotNItems > 0)
      {
         long nLoc = m_lotNext - 1;
         SYSTEMTIME sTime;

         for (unsigned long i = 0; i < m_lotNItems; i++)
         {
            if (nLoc < 0)
               nLoc = m_lotHistory.length() - 1;

            VariantTimeToSystemTime(m_lotHistory[nLoc].timeStamp, &sTime);

            sprintf(localBuf, "   %s @ %.4d/%-.2d/%-.2d %.2d:%.2d:%.2d.%.3d",
               m_lotHistory[nLoc].pathName, sTime.wYear, sTime.wMonth,
               sTime.wDay, sTime.wHour, sTime.wMinute, sTime.wSecond,
               sTime.wMilliseconds);
            logTrace(funcName, localBuf);

            nLoc--;
         }
      }
   }
}

bool FileNotifyCheck::init(unsigned int size, long ageLimit)
{
   static const char *funcName = "FileNotifyCheck::init()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   if (m_didInit)
   {
      m_fileHistory.setCapacity(0);
      m_lotHistory.setCapacity(0);
      m_didInit = false;
   }

   if (m_fileHistory.setLength(size) < 0)
   {
      logErrorMsg(errSmmResource, funcName, "FileNotifyCheck allocation.");
      return m_didInit;
   }

   m_next = 0;
   m_nItems = 0;

// as per discussion with Ed the size of the lot notification history
// is reduced to 1

   if (m_lotHistory.setLength(1) < 0)
   {
      logErrorMsg(errSmmResource, funcName, "FileNotifyCheck allocation.");
      return m_didInit;
   }

   m_lotNext = 0;
   m_lotNItems = 0;

// the ageLimit parameter is a long (seconds).  this is converted
// to a variant format (double) so it will be fractional (only used for maps)

   m_ageLimit = ((double) ageLimit) / 86400.0;
   m_didInit = true;

   return m_didInit;
}

void FileNotifyCheck::addMap(const char *pathName, double timeStamp)
{
   static const char *funcName = "FileNotifyCheck::addMap()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   if (! m_didInit)
      return;

   strcpy(m_fileHistory[m_next].pathName, pathName);
   m_fileHistory[m_next].timeStamp = timeStamp;

   if (m_nItems < m_fileHistory.length())
      m_nItems++;

   m_next++;
   if (m_next >= m_fileHistory.length())
      m_next = 0;
}

// could do these linear but will be a little more efficient and look
// backward from the most recent

bool FileNotifyCheck::isADupMap(const char *pathName, double timeStamp)
{
   static const char *funcName = "FileNotifyCheck::isADupMap()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   if (! m_didInit)
      return false;

   long nLoc = m_next - 1;

   for (unsigned long i = 0; i < m_nItems; i++)
   {
      if (nLoc < 0)
         nLoc = m_fileHistory.length() - 1;

      if (stricmp(m_fileHistory[nLoc].pathName, pathName) == 0)
      {
         // check the timestamp
         if (fabs(timeStamp - m_fileHistory[nLoc].timeStamp) <= m_ageLimit)
         {
            return true;
         }
      }

      nLoc--;
   }
   return false;
}

void FileNotifyCheck::addLot(const char *pathName)
{
   static const char *funcName = "FileNotifyCheck::addLot()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   if (! m_didInit)
      return;

   strcpy(m_lotHistory[m_lotNext].pathName, pathName);

   if (m_lotNItems < m_lotHistory.length())
      m_lotNItems++;

   m_lotNext++;
   if (m_lotNext >= m_lotHistory.length())
      m_lotNext = 0;
}

// could do these linear but will be a little more efficient and look
// backward from the most recent

bool FileNotifyCheck::isADupLot(const char *pathName)
{
   static const char *funcName = "FileNotifyCheck::isADupLot()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   if (! m_didInit)
      return false;

   long nLoc = m_lotNext - 1;

   for (unsigned long i = 0; i < m_lotNItems; i++)
   {
      if (nLoc < 0)
         nLoc = m_lotHistory.length() - 1;

      if (stricmp(m_lotHistory[nLoc].pathName, pathName) == 0)
      {
         return true;
      }

      nLoc--;
   }
   return false;
}

QueueEntry::QueueEntry()
{
   m_next = 0;
   m_pathName[0] = '\0';
   m_timeStamp = 0.0;
   m_action = 0;
   m_fromScavenge = false;
}

QueueEntry::QueueEntry(const char *pathName, DWORD action, bool fromScavenge)
{
   SYSTEMTIME sTime;

   strcpy(m_pathName, pathName);
   m_next = 0;
   m_action = action;
   m_fromScavenge = fromScavenge;

   GetSystemTime(&sTime);
   SystemTimeToVariantTime(&sTime, &m_timeStamp);
}

void MonitorShare::runScavenge()
{
   static const char *funcName = "MonitorShare::runScavenge()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   m_initialized = true;
   SetEvent(m_scavengeEvent);

   if (logTraceLevel >= logTraceLevelDetail)
   {
      char localBuf[_MAX_PATH+151];

      sprintf(localBuf, "Scavenge thread started for %s '%s'", m_monitorId,
            m_dirPath);
      logTrace(funcName, localBuf);
   }

   bool done = false;

// the timeout period will be set to 5 seconds maximum and a count will
// be used to determine when to initiate a scavenge

   unsigned long nTimeouts = 0;
   unsigned long retryTime = 0;
   unsigned long maxTimeouts = (m_scavengeStart + (SCAVENGE_TIMEOUT - 1)) /
                     SCAVENGE_TIMEOUT;
   unsigned long scavengeTimeout = SCAVENGE_TIMEOUT * 1000;

// there is also a limit placed on the number of simultaneous scavenge
// operations.  the implemented solution isn't optimal and doesn't resolve
// ordering or priority when the maximum is reached; however, it does
// keep it under control

   while (! done)
   {
      DWORD waitResult = MsgWaitForMultipleObjectsEx(0, NULL,
               scavengeTimeout, QS_ALLINPUT, MWMO_ALERTABLE);
      if (waitResult == WAIT_OBJECT_0)
      {
         MSG msg;

         while ((! done) && PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
         {
            if (logTraceLevel >= logTraceLevelDetail)
            {
               char localBuf[512];

               sprintf(localBuf,"Message %lu received for %s", msg.message,
                     m_monitorId);
               logTrace(funcName, localBuf);
            }

            if (msg.message == WM_QUIT)
            {
               done = true;
            }
            else
            {
               DispatchMessage(&msg);
            }
         }
      }
      else if (waitResult == WAIT_TIMEOUT)
      {
         nTimeouts++;
         if (nTimeouts >= maxTimeouts || m_triggerScavenge)
         {
            if (_Module.okayToScavenge(retryTime))
            {
               if (processScavenge() == Error_Failure)
               {
                  setRequestShutdown(true);
                  done = true;
               }
               else
               {
                  nTimeouts = 0;
                  maxTimeouts = (m_scavengeTime + (SCAVENGE_TIMEOUT - 1)) /
                        SCAVENGE_TIMEOUT; // reset the timeout period
                  scavengeTimeout = SCAVENGE_TIMEOUT * 1000;
               }

               _Module.scavengeDone();   // notify done so others may scavenge
            }
            else
            {
               if (logTraceLevel >= logTraceLevelDetail)
               {
                  char localBuf[256];

                  sprintf(localBuf,"Scavenge will be retried for %s in %d ms",
                        m_monitorId, retryTime);
                  logTrace(funcName, localBuf);
               }

               scavengeTimeout = retryTime;
            }
         }
      }
      else
      {
         if (logTraceLevel)
         {
            char localBuf[128];

            DWORD eCode = GetLastError();
            sprintf(localBuf,"Unexpected queue wait result = %lu (%ld)",
                  waitResult, eCode);
            logTrace(funcName, localBuf);

            setRequestShutdown(true);
         }
      }
   }

   if (logTraceLevel >= logTraceLevelDetail)
   {
      char localBuf[_MAX_PATH+151];

      sprintf(localBuf, "Scavenge thread completed for %s '%s'", m_monitorId,
            m_dirPath);
      logTrace(funcName, localBuf);
   }
}

long MonitorShare::processScavenge()
{
   static const char *funcName = "MonitorShare::processScavenge()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (logTraceLevel >= logTraceLevelDetail)
   {
      char localBuf[_MAX_PATH+256];

      sprintf(localBuf,"Running scavenge for %s '%s'", m_monitorId, m_dirPath);
      logTrace(funcName, localBuf);
   }

   if (m_inShutdown || _Module.inShutdown())
      return retval;

   try
   {
      DirList dList;
      char searchString[_MAX_PATH+1];
      long notifyCount = 0;

      strcpy(searchString, m_dirPath);
      strcat(searchString, "*");

      if (buildFileDirList(dList, searchString, true) == Error_Success)
      {
         DirEntry *pEntry;
         DirEntryIterator itDir(&dList);

         while (pEntry = itDir())
         {
            if (m_inShutdown || _Module.inShutdown())
               return retval;

            if (pEntry->fileAttrib & FILE_ATTRIBUTE_DIRECTORY)
            {
               if (m_lotDir.compare(pEntry->fileName))
               {
                  pEntry->notify = true;
                  notifyCount++;
               }
            }
            else
            {
               SmmExtension **pExt;
               SmmExtensionIterator it(&m_extInfo);
               bool foundMatch = false;

               while (foundMatch == false && (pExt = it()))
               {
                  if (*pExt)
                  {
                     SmmExtension::MonitorAction fStatus = (*pExt)->find(
                           pEntry->pathName);
                     if (fStatus == SmmExtension::ActionOK)
                     {
                        pEntry->notify = true;
                        foundMatch = true;
                        notifyCount++;
                     }
                  }
               }
            }
         }
         
         if (notifyCount > 0)
         {
            double newestLotTimeStamp = 0.0;

            if (m_inShutdown || _Module.inShutdown())
               return retval;

            if (m_scavenge && m_lenScavenge > 0)
            {
               if (runScavengeScript(dList) == Error_Failure)
                  return retval;
            }
            else
            {
               DirEntry *pEntry, *pNewestLot = 0;
               DirEntryIterator itDir(&dList);

               while (pEntry = itDir())
               {
                  if (m_inShutdown || _Module.inShutdown())
                     return retval;

                  if (pEntry->notify)
                  {

// the default for lots is to only notify the action manager for the newest
// lot.

                     if (pEntry->fileAttrib & FILE_ATTRIBUTE_DIRECTORY)
                     {
                        if (pEntry->modifiedTime > newestLotTimeStamp)
                        {
                           pNewestLot = pEntry;
                           newestLotTimeStamp = pEntry->modifiedTime;
                        }
                        else if (pEntry->creationTime > newestLotTimeStamp)
                        {
                           pNewestLot = pEntry;
                           newestLotTimeStamp = pEntry->creationTime;
                        }

                        pEntry->notify = false;
                     }
                  }
               }

               if (pNewestLot)      // turn on notification for only this lot
                  pNewestLot->notify = true;
            }

// hand the maps (and lots) off to the queue

            {
               DirEntry *pEntry;
               DirEntryIterator itDir(&dList);

               while (pEntry = itDir())
               {
                  if (m_inShutdown || _Module.inShutdown())
                     return retval;

                  if (pEntry->notify)
                  {
                     if (logTraceLevel >= logTraceLevelFiles)
                     {
                        char localBuf[_MAX_PATH+256];

                        sprintf(localBuf,
                           "Scavenge found %s (may need processing).",
                           pEntry->pathName);
                        logTrace(funcName, localBuf);
                     }

                     addToQueue(pEntry->pathName, FILE_ACTION_ADDED, true);
                  }
               }
            }
         }
         retval = Error_Success;
      }
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
      return Error_Failure;
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
      return Error_Failure;
   }

   EnterCriticalSection(&m_qLock);

   m_triggerScavenge = false;
   m_scavengeSurgeExceeded = false;

   LeaveCriticalSection(&m_qLock);

   return retval;
}

class LocalPointersFileDirList
{
   public:
      HANDLE dirHandle;

      LocalPointersFileDirList() : dirHandle(0) {}
      ~LocalPointersFileDirList() { if (dirHandle) FindClose(dirHandle); }
};

long MonitorShare::buildFileDirList(DirList& dList, const char *searchString,
                        bool recurseSubDirs)
{
   static const char *funcName = "MonitorShare::buildFileDirList()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   LocalPointersFileDirList LP;
   WIN32_FIND_DATA fileData;
   SYSTEMTIME sTime;
   char baseName[_MAX_PATH+1];

   long retval = Error_Failure;

// copy the base folder name to a local buffer (including the trailing '\')

   try
   {
      long ls = strlen(searchString);
      long idx = -1;
      long lnBase = 0;
      long i;

      for (i = ls-1; i > 0; i--)
      {
         if (searchString[i] == '\\' || searchString[i] == '/')
         {
            idx = i;
            break;
         }
      }

      for (i = 0; i <= idx; i++)
         baseName[lnBase++] = searchString[i];
      baseName[lnBase] = '\0';

      LP.dirHandle = FindFirstFile(searchString, &fileData);
      if (LP.dirHandle == INVALID_HANDLE_VALUE)
      {
         char localBuf[BUFSIZ+256];
         char errBuf[BUFSIZ+1];

         DWORD eCode = GetLastError();
         if (! FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, eCode, 0,
               errBuf, BUFSIZ, 0))
         {
            sprintf(localBuf,
               "FindFirstFile() returned %ld for %s '%s'.", eCode, m_monitorId,
               m_dirPath);
         }
         else
         {
            sprintf(localBuf,
               "FindFirstFile() returned %ld (%s) for %s '%s'.", eCode,
               errBuf, m_monitorId, m_dirPath);
         }
         logErrorMsg(errSmmIO, funcName, localBuf);
         return retval;
      }

      double timeStamp = 0.0;
      unsigned long fCount = dList.length();
      BOOL moreFiles = FALSE;   // this is a Win32 BOOL, not a C++ bool

      do
      {
         if (m_inShutdown || _Module.inShutdown())
            return retval;

         unsigned long lf = strlen(fileData.cFileName);
         bool skipDotDir = false;

         if (lf <= 2)
         {
            if (strcmp(fileData.cFileName, ".") == 0)
               skipDotDir = true;
            else if (strcmp(fileData.cFileName, "..") == 0)
               skipDotDir = true;
         }

         if (skipDotDir == false)
         {
            if (dList.setLength(fCount+1))
            {
               logErrorMsg(errSmmResource, funcName,
                     "Directory and folder list allocation.");
               return retval;
            }

            strcpy(dList[fCount].pathName, baseName);
            strcat(dList[fCount].pathName, fileData.cFileName);
            strcpy(dList[fCount].fileName, fileData.cFileName);

            dList[fCount].fileAttrib = fileData.dwFileAttributes;

            if (FileTimeToSystemTime(&fileData.ftCreationTime, &sTime))
            {
               SystemTimeToVariantTime(&sTime, &timeStamp);
               dList[fCount].creationTime = timeStamp;
            }

            if (FileTimeToSystemTime(&fileData.ftLastWriteTime, &sTime))
            {
               SystemTimeToVariantTime(&sTime, &timeStamp);
               dList[fCount].modifiedTime = timeStamp;
            }

            fCount++;

            if (fileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
            {
               if (recurseSubDirs)
               {
                  char newSearch[_MAX_PATH+1];

                  strcpy(newSearch, dList[fCount-1].pathName);
                  strcat(newSearch, "\\*");

                  retval = buildFileDirList(dList, newSearch, recurseSubDirs);
                  if (retval != Error_Success)
                     return retval;

                  fCount = dList.length();
               }
            }
         }

         moreFiles = FindNextFile(LP.dirHandle, &fileData);
      }
      while (moreFiles);

      retval = Error_Success;
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}

class LocalPointersScavengeScript
{
   public:
      char **pathArray;
      char *strValue;
      double *attribArray;
      double *notifyArray;
      unsigned long nPaths;

      LocalPointersScavengeScript() : pathArray(0), nPaths(0), strValue(0),
                                      attribArray(0), notifyArray(0) {}
      ~LocalPointersScavengeScript() { delete strValue; delete attribArray;
                                       delete notifyArray;
                         if (pathArray) {
                            for (unsigned long i = 0; i < nPaths; i++)
                               delete pathArray[i];
                            delete pathArray;
                         }
                      }
};

long MonitorShare::runScavengeScript(DirList& dList)
{
   static const char *funcName = "MonitorShare::runScavengeScript()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   try
   {
      eServer statitLib(m_useServerPool);

      if (statitLib.init() == Error_Failure)
         return retval;

      if (statitLib.putGlobalString("%SMM_MonitorId", m_monitorId) ==
            Error_Failure)
         return retval;

      if (statitLib.putGlobalInt("$SMM_Scavenge", 1) == Error_Failure)
         return retval;

      if (statitLib.putGlobalString("%SMM_Scavenge", m_scavenge) ==
               Error_Failure)
         return retval;

// load the data into the workspace

      if (statitLib.createStringVar("Scavenge.DirPath", _MAX_PATH) ==
               Error_Failure)
         return retval;

      if (statitLib.createDoubleVar("Scavenge.PathAttributes") == Error_Failure)
         return retval;

      if (statitLib.createDoubleVar("Scavenge.Notify") == Error_Failure)
         return retval;

      {
         DirEntry *pEntry;
         DirEntryIterator itDir(&dList);
         long rowNum = 1;      // initially it is row 1, then append (-2)

         while (pEntry = itDir())
         {
            if (m_inShutdown || _Module.inShutdown())
               return retval;

            if (pEntry->notify)
            {
               if (statitLib.putString("Scavenge.DirPath", rowNum,
                     pEntry->pathName) == Error_Failure)
                  return retval;

               if (statitLib.putValue("Scavenge.PathAttributes", rowNum,
                           (double) pEntry->fileAttrib) == Error_Failure)
                  return retval;

               if (statitLib.putValue("Scavenge.Notify", rowNum, 1.0) ==
                           Error_Failure)
                  return retval;

               rowNum = -2;   // append
            }
         }
      }

      if (m_inShutdown || _Module.inShutdown())
         return retval;

      if (statitLib.runCmd("smm_runScavenge") == Error_Failure)
         return retval;

      long smm_scavenge = 0;
      if (statitLib.getGlobalInt("$SMM_Scavenge", smm_scavenge) ==
               Error_Failure)
         return retval;

      dList.setLength(0);   // flush the existing data, will be reloaded or none

// if the scavenge value is set, reload the data

      if (smm_scavenge > 0)
      {
         LocalPointersScavengeScript LP;
         unsigned long nCases = 0;
         long fCount = 0;

         if (statitLib.getStringVar("Scavenge.DirPath", LP.nPaths,
                     LP.pathArray) == Error_Failure)
            return retval;

         if (statitLib.getDoubleVar("Scavenge.PathAttributes", nCases,
                     LP.attribArray) == Error_Failure)
            return retval;

         if (statitLib.getDoubleVar("Scavenge.Notify", nCases,
                     LP.notifyArray) == Error_Failure)
            return retval;

         for (unsigned long i = 0; i < LP.nPaths; i++)
         {
            if (m_inShutdown || _Module.inShutdown())
               return retval;

            if (LP.notifyArray[i] > 0.0)
            {
               if (dList.setLength(fCount+1))
               {
                  logErrorMsg(errSmmResource, funcName,
                        "Directory and folder list allocation.");
                  return retval;
               }

               strcpy(dList[fCount].pathName, LP.pathArray[i]);
               dList[fCount].fileAttrib = (DWORD) LP.attribArray[i];
               dList[fCount].notify = true;

               // unused fields
               dList[fCount].fileName[0] = '\0';
               dList[fCount].modifiedTime = 0.0;
               dList[fCount].creationTime = 0.0;

               fCount++;
            }
         }
      }

      retval = Error_Success;
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}
