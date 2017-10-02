// SMMonitor.cpp : Implementation of WinMain


// Note: Proxy/Stub Information
//      To build a separate proxy/stub DLL, 
//      run nmake -f SMMonitorps.mk in the project directory.

#include "stdafx.h"
#include "resource.h"
#include <initguid.h>
#include "SMMonitor.h"

#include "SMMonitor_i.c"

#include <stdio.h>
#include <stdlib.h>
#include <time.h>
#include <io.h>
#include "EGEventLog.h"
#include "eserver.h"
#include "smmerrors.h"

#define MAIN_LOOP_TIMEOUT 5000

const char * const smmDispatchEventName = "Global\\smm_Dispatch_Ready";

const unsigned long maxRetryCount = 3;
const unsigned long retryDelay = 50;   // milliseconds

const double scavengeRetryMax = 10000.0; // 10 seconds

CServiceModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
END_OBJECT_MAP()

void MonitorList::allocateProc(size_t first, size_t last)
{
   do
   {
      MonitorRecord *p = (MonitorRecord *) nodeFast(first);

      p->inUse = false;
      p->monitorShare = 0;
      p->monitorId = 0;
      p->dirPath = 0;
      p->numExtensions = 0;
      p->extensions = 0;
      p->lotDir = 0;
      p->verifyMap = 0;
      p->mapAction = 0;
      p->verifyLot = 0;
      p->lotAction = 0;
      p->scavenge = 0;
      p->scavengeTime = defaultScavengeTime;
      p->scavengeStart = defaultScavengeStart;
   }
   while (++first <= last);
}

void MonitorList::deallocateProc(size_t first, size_t last)
{
   do
   {
      MonitorRecord *p = (MonitorRecord *) nodeFast(first);
      if (p)
      {
         if (p->monitorShare)
         {
            p->monitorShare->shutdown();
            delete p->monitorShare;
            p->monitorShare = 0;
         }

         delete p->monitorId;
         p->monitorId = 0;

         delete p->dirPath;
         p->dirPath = 0;

         if (p->extensions)
         {
            for (unsigned long i = 0; i < p->numExtensions; i++)
            {
               delete p->extensions[i];
               p->extensions[i] = 0;
            }

            delete p->extensions;
            p->extensions = 0;
         }

         delete p->lotDir;
         p->lotDir = 0;

         delete p->verifyMap;
         p->verifyMap = 0;

         delete p->mapAction;
         p->mapAction = 0;

         delete p->verifyLot;
         p->verifyLot = 0;

         delete p->lotAction;
         p->lotAction = 0;

         delete p->scavenge;
         p->scavenge = 0;

         p->scavengeTime = defaultScavengeTime;
         p->scavengeStart = defaultScavengeStart;

         p->inUse = false;
      }
   }
   while (++first <= last);
}

void MonitorFailureList::allocateProc(size_t first, size_t last)
{
   do
   {
      MonitorFailure *p = (MonitorFailure *) nodeFast(first);

      p->inUse = false;
      p->monitorId = 0;
      p->dirPath = 0;
      p->timeStamp = 0.0;
      p->eCode = 0;
      p->notADir = false;
   }
   while (++first <= last);
}

void MonitorFailureList::deallocateProc(size_t first, size_t last)
{
   do
   {
      MonitorFailure *p = (MonitorFailure *) nodeFast(first);
      if (p)
      {
         delete p->monitorId;
         p->monitorId = 0;

         delete p->dirPath;
         p->dirPath = 0;

         p->timeStamp = 0;
         p->eCode = 0;
         p->notADir = false;
         p->inUse = false;
      }
   }
   while (++first <= last);
}

unsigned long strtrimlen(const char *s)
{
   unsigned long ls = strlen(s);
   while (ls > 0 && s[ls-1] <= ' ')
      ls--;

   return ls;
}

LPCTSTR FindOneOf(LPCTSTR p1, LPCTSTR p2)
{
   while (p1 != NULL && *p1 != NULL)
   {
      LPCTSTR p = p2;
      while (p != NULL && *p != NULL)
      {
         if (*p1 == *p)
            return CharNext(p1);
         p = CharNext(p);
      }
      p1 = CharNext(p1);
   }
   return NULL;
}

class LocalPointersDispatchThread
{
   public:
      LocalPointersDispatchThread() {}
      ~LocalPointersDispatchThread() { CoUninitialize(); ExitThread(0); }

};


DWORD WINAPI dispatchThread(LPVOID lpParam)
{
   static const char *funcName = "dispatchThread()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   LocalPointersDispatchThread LP;

   SMMDispatcher *pDispatch = (SMMDispatcher *) lpParam;

   CoInitializeEx(NULL,COINIT_MULTITHREADED);

   try
   {
      pDispatch->run();
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
      return Error_Failure;
   }
   catch (...)
   {
      char localBuf[128];

      sprintf(localBuf, "Exception in dispatcher thread: Id=%i",
            GetCurrentThreadId());
      logErrorMsg(errSmmUnknownException, funcName, localBuf);
      return Error_Failure;
   }

   return Error_Success;
}

SMMDispatcher::SMMDispatcher()
{
   m_dispatchThreadHandle = 0;
   m_dispatchEvent = 0;
   m_dispatchThreadId = 0;

   m_useAltDispatcher = false;
   m_altDispatcher[0] = '\0';

   m_didSetup = false;
   m_initialized = false;
   m_useServerPool = false;

   m_iDispatcher = 0;

   InitializeCriticalSection(&m_lock);
}

SMMDispatcher::~SMMDispatcher()
{
   static const char *funcName = "SMMDispatcher::~SMMDispatcher()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   shutdown();

   DeleteCriticalSection(&m_lock);
}

long SMMDispatcher::setup(bool useAltDispatcher, const char *altDispatcher,
               unsigned long timeOut, bool useServerPool)
{
   static const char *funcName = "SMMDispatcher::setup()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   try
   {
      if (m_didSetup)
      {
         shutdown();
         m_didSetup = false;
      }

      m_useAltDispatcher = useAltDispatcher;
      if (m_useAltDispatcher)
         strcpy(m_altDispatcher, altDispatcher);

      m_timeout = timeOut;
      m_useServerPool = useServerPool;
      m_didSetup = true;
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

long SMMDispatcher::start()
{
   static const char *funcName = "SMMDispatcher::start()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   if (! m_didSetup)
   {
      logErrorMsg(errSmmSetupNotCalled, funcName,
         "The setup() method must be call prior to the start() method.");
      return Error_Failure;
   }

// create the global dispatch event

   if (m_globalDispatchEvent)
   {
      CloseHandle(m_globalDispatchEvent);
      m_globalDispatchEvent = 0;
   }

   if (m_useAltDispatcher)
      m_globalDispatchEvent = CreateEvent(0, TRUE, TRUE, "smm_Dispatch_Ready");
   else
      m_globalDispatchEvent = CreateEvent(0, TRUE, TRUE, smmDispatchEventName);

   if (! m_globalDispatchEvent)
   {
      char localBuf[128];

      DWORD eCode = GetLastError();
      sprintf(localBuf,"CreateEvent(global dispatch) returned %ld.",eCode);
      logErrorMsg(errSmmResource, funcName, localBuf);
      return Error_Failure;
   }

   m_dispatchEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
   if (m_dispatchEvent)
   {
      m_dispatchThreadHandle = CreateThread(NULL, 0, dispatchThread,
            (void *) this, 0,(unsigned long *) &m_dispatchThreadId);

      if (m_dispatchThreadHandle)
      {
         WaitForSingleObject(m_dispatchEvent, INFINITE);
         ResetEvent(m_dispatchEvent);
      }
      else
      {
         char localBuf[128];

         DWORD eCode = GetLastError();
         sprintf(localBuf,"CreateThread(dispatch) returned %ld.",eCode);
         logErrorMsg(errSmmResource, funcName, localBuf);
         return Error_Failure;
      }
   }
   else
   {
      char localBuf[128];

      DWORD eCode = GetLastError();
      sprintf(localBuf,"CreateEvent(dispatch) returned %ld.",eCode);
      logErrorMsg(errSmmResource, funcName, localBuf);
      return Error_Failure;
   }

   return Error_Success;
}


long SMMDispatcher::releaseInstance()
{
   static const char *funcName = "SMMDispatcher::releaseInstance()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   try
   {
      if (m_iDispatcher)
      {
         m_iDispatcher->Release();
         m_iDispatcher = 0;
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


bool SMMDispatcher::haveInstance()
{
   static const char *funcName = "SMMDispatcher::haveInstance()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   bool retval = false;

   if (m_iDispatcher)
      retval = true;

   return retval;
}

long SMMDispatcher::getInstance()
{
   static const char *funcName = "SMMDispatcher::getInstance()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   try
   {
      if (haveInstance())
         return Error_Success;

      unsigned long tryCount = 0;

      do
      {
         if (m_iDispatcher)
         {
            m_iDispatcher->Release();
            m_iDispatcher = 0;
         }

         if (tryCount > 0)               // very brief delay
            ::Sleep(retryDelay);

         HRESULT hr = CoCreateInstance(m_clsIdDispatcher, 0, CLSCTX_SERVER,
                           __uuidof(_clsDispatcher), (void **) &m_iDispatcher);
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
                     "Failed to create an instance of SMDispatcher.clsDispatcher. Error code=%x.",
                     hr);
               }
               else
               {
                  sprintf(localBuf,
                     "Failed to create an instance of SMDispatcher.clsDispatcher (%s). Error code=%x.",
                     errBuf, hr);
               }
               logErrorMsg(errSmmConfiguration, funcName, localBuf);
            }
         }
         else
         {
            long rStatus = 0;
            hr = m_iDispatcher->init(&rStatus);
            if (FAILED(hr))
            {
               if (tryCount >= maxRetryCount)
               {
                  char localBuf[256];

                  sprintf(localBuf,
                     "SMDispatcher.init() failed. Error code=%x.", hr);
                  logErrorMsg(errSmmDispatcher, funcName, localBuf);
               }
            }
            else
            {
               if (rStatus != 0)
               {
                  if (tryCount >= maxRetryCount)
                  {
                     logErrorMsg(errSmmDispatcher, funcName,
                        "SMDispatcher.init() returned failure.");
                  }
               }
               else
               {
                  retval = Error_Success;
               }
            }
         }
         tryCount++;
      }
      while (retval == Error_Failure && tryCount <= maxRetryCount);
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   if (retval == Error_Failure && m_iDispatcher)
   {
      m_iDispatcher->Release();
      m_iDispatcher = 0;
   }

   return retval;
}

void SMMDispatcher::run()
{
   static const char *funcName = "SMMDispatcher::run()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   m_initialized = true;
   SetEvent(m_dispatchEvent);

   try
   {
      if (! m_useAltDispatcher)
      {
         HRESULT hr = CLSIDFromProgID(OLESTR("SMDispatcher.clsDispatcher"),
                           &m_clsIdDispatcher);
         if (FAILED(hr))
         {
            char errBuf[BUFSIZ+1];
            char localBuf[BUFSIZ+256];

            if (! FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, hr, 0,
                  errBuf, BUFSIZ, 0))
            {
               sprintf(localBuf,
                  "Failed to lookup the SMDispatcher.clsDispatcher CLSID. Error code=%x.",
                  hr);
            }
            else
            {
               sprintf(localBuf,
                  "Failed to lookup the SMDispatcher.clsDispatcher CLSID (%s). Error code=%x.",
                  errBuf, hr);
            }
            logErrorMsg(errSmmConfiguration, funcName, localBuf);
            return;
         }
      }

      bool done = false;
      bool moreRecords = false;
      unsigned long lTimeout = m_timeout * 1000;

      while (! done)
      {
         if (moreRecords)
         {
            moreRecords = false;
            lTimeout = 100;                // timeout quickly, but allows a quit
         }
         else
         {
            lTimeout = m_timeout * 1000;   // back to normal
         }

         DWORD waitResult = MsgWaitForMultipleObjectsEx(1, &m_globalDispatchEvent,
                  lTimeout, QS_ALLINPUT, MWMO_ALERTABLE);
         if (waitResult == WAIT_OBJECT_0)
         {
/*
 * don't reset the event until after the dispatcher has completed its work
 * since it will already pick up most new arrivals
 */
            runDispatcher(moreRecords);
            ResetEvent(m_globalDispatchEvent);
         }
         else if (waitResult == WAIT_OBJECT_0 + 1)
         {
            MSG msg;

            while ((! done) && PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
            {
               if (logTraceLevel >= logTraceLevelDetail)
               {
                  char localBuf[128];

                  sprintf(localBuf,"Message %lu received for dispatcher thread",
                        msg.message);
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
         else if (waitResult == WAIT_TIMEOUT)      // run to flush anything missed
         {
            runDispatcher(moreRecords);
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

   if (m_iDispatcher)
   {
      m_iDispatcher->Release();
      m_iDispatcher = 0;
   }
}

// return -1 for failure, 0 for success

long SMMDispatcher::runDispatcher(bool& moreRecords)
{
   static const char *funcName = "SMMDispatcher::runDispatcher()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;
   moreRecords = false;

   if (_Module.inShutdown())
      return retval;

   try
   {
      if (m_useAltDispatcher)
      {
         eServer statitLib(m_useServerPool);

         if (statitLib.init() == Error_Failure)
            return retval;

// the loadMonitorData buffer contains the proc or macro to execute.  the
// SMM proc, sw_getSMMData, will be used to execute it and validate the data
// global StatitScript variables are used to exchange run information and
// the workspace must contain the dataset Monitor.

         if (statitLib.putGlobalString("%SMM_Dispatcher", m_altDispatcher)
                  == Error_Failure)
            return retval;

         if (logTraceLevel >= logTraceLevelDetail)
         {
            char localBuf[512];

            sprintf(localBuf,"Running dispatcher '%s'.", m_altDispatcher);
            logTrace(funcName, localBuf);
         }

         if (statitLib.runCmd("smm_runDispatcher") == Error_Failure)
            return retval;

         retval = Error_Success;
      }
      else
      {
         if (getInstance() == Error_Failure)
            return retval;

         long rStatus = 0;

         HRESULT hr = m_iDispatcher->ProcessQueue(&rStatus);
         if (FAILED(hr))
         {
            char localBuf[256];

            sprintf(localBuf,
                  "SMDispatcher.ProcessQueue() failed. Error code=%x.", hr);
            logErrorMsg(errSmmDispatcher, funcName, localBuf);
         }
         else
         {
            retval = Error_Success;
            if (rStatus < 0)            // < 0 indicates more records waiting
               moreRecords = true;
         }

         if (moreRecords == false)      // or an error
            releaseInstance();
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

   return retval;
}


const int THREAD_SHUTDOWN_WAIT = 15000;   // 15 seconds

void SMMDispatcher::shutdown()
{
   static const char *funcName = "SMMDispatcher::shutdown()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   EnterCriticalSection(&m_lock);

   if (logTraceLevel >= logTraceLevelDetail)
   {
      logTrace(funcName, "Shutting down the dispatcher thread.");
   }

   try
   {
      if (m_initialized)
      {
         m_initialized = false;

         DWORD reason = WAIT_TIMEOUT;
         BOOL postSuccess = PostThreadMessage(m_dispatchThreadId, WM_QUIT,
                        0, 0);
         if (postSuccess == TRUE)
         {
            reason = MsgWaitForMultipleObjects(1, &m_dispatchThreadHandle,
                        FALSE, THREAD_SHUTDOWN_WAIT, QS_ALLINPUT);
         }

         if (postSuccess == FALSE)
         {
            char localBuf[128];

            DWORD eCode = GetLastError();
            sprintf(localBuf,"PostThreadMessage(dispatcher) returned %ld.",
                  eCode);
            logErrorMsg(errSmmPostThread, funcName, localBuf);
         }
         else if (reason == WAIT_TIMEOUT)
         {
            logErrorMsg(errSmmWaitTimeout, funcName,
                  "SMMDispatcher::shutdown() thread timeout");
         }

         // in case of a problem during startup signal the event so that
         // the infinite wait will proceed

         if (m_dispatchEvent)
            SetEvent(m_dispatchEvent);

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

   if (m_dispatchThreadHandle)
   {
      CloseHandle(m_dispatchThreadHandle);
      m_dispatchThreadHandle = 0;
   }

   if (m_dispatchEvent)
   {
      CloseHandle(m_dispatchEvent);
      m_dispatchEvent = 0;
   }

   if (m_globalDispatchEvent)
   {
      CloseHandle(m_globalDispatchEvent);
      m_globalDispatchEvent = 0;
   }

   LeaveCriticalSection(&m_lock);
}

bool SMMDispatcher::alive()
{
   static const char *funcName = "SMMDispatcher::alive()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   bool retval = false;

   EnterCriticalSection(&m_lock);

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

   LeaveCriticalSection(&m_lock);

   return retval;
}

CServiceModule::CServiceModule() : CComModule()
{
   logTraceLevel = logTraceLevelNone;
   logTraceFile[0] = '\0';

   m_loadMonitorData[0] = '\0';
   m_useServerPool = true;
   m_failureReportInterval = 1.0;

   m_useAltDispatcher = false;
   m_altDispatcher[0] = '\0';
   m_dispatchTimeout = 60;           // seconds

   m_inShutdown = false;

   m_loadTimer = 0;
   m_loadInterval = 1000 * 60 * 5;   // 5 minutes

// the following is the threshold that is used to determine if a scavenge
// needs to be triggered from a surge in file or folder notifications.  if
// the number of items in the queue for a folder exceeds this value then
// when the queue finally drops to 0, a scavenge will be initiated.

// a smaller value is now being used to handle tools that drop blocks of
// files from their local drive to the mapped drive (or UNC).  the dir
// changed notification is triggered when the files are dropped, but they
// cannot be read yet.  a second notification isn't triggered with a large
// enough margin

   m_scavengeSurgeTrigger = 10;

// the following is used to limit the number of simultaneous scavenges

   m_scavengeMax = 5;
   m_scavengeCount = 0;

   srand(time(NULL));   // seed the random number generator

   InitializeCriticalSection(&m_scavengeLock);
}

CServiceModule::~CServiceModule()
{
   static const char *funcName = "CServiceModule::~CServiceModule()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   stopTimers();

   m_monitorList.setCapacity(0);
   m_failureList.setCapacity(0);

   DeleteCriticalSection(&m_scavengeLock);

   LogEvent(EVENTLOG_INFORMATION_TYPE,
      _T("SORTmanager Map Management service stopped"));
}

long CServiceModule::createTimers()
{
   static const char *funcName = "CServiceModule::createTimers()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

// set up the timers.  the load timer is used to determine how often to
// refresh the list of folders to be monitored.

   m_loadTimer = SetTimer(NULL, 0, m_loadInterval, (TIMERPROC) NULL);
   if (m_loadTimer == 0)
   {
      logErrorMsg(errSmmSetTimerFailed, funcName, "Load monitoring data timer");
      return Error_Failure;
   }

   return Error_Success;
}

void CServiceModule::stopTimers()
{
   static const char *funcName = "CServiceModule::stopTimers()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   if (m_loadTimer)
   {
      KillTimer(NULL, m_loadTimer);
      m_loadTimer = 0;
   }
}

// Although some of these functions are big they are declared inline since they are only used once

inline HRESULT CServiceModule::RegisterServer(BOOL bRegTypeLib, BOOL bService)
{
   HRESULT hr = CoInitialize(NULL);
   if (FAILED(hr))
      return hr;

   // Remove any previous service since it may point to
   // the incorrect file
   Uninstall();

   // Add service entries
   UpdateRegistryFromResource(IDR_SMMonitor, TRUE);

   // Adjust the AppID for Local Server or Service
   CRegKey keyAppID;
   LONG lRes = keyAppID.Open(HKEY_CLASSES_ROOT, _T("AppID"), KEY_WRITE);
   if (lRes != ERROR_SUCCESS)
      return lRes;

   CRegKey key;
   lRes = key.Open(keyAppID, _T("{53D6D93C-69F1-4D90-B7A0-8A4148F2AC1D}"), KEY_WRITE);
   if (lRes != ERROR_SUCCESS)
      return lRes;
   key.DeleteValue(_T("LocalService"));
    
   if (bService)
   {
      key.SetValue(_T("SMMonitor"), _T("LocalService"));
      key.SetValue(_T("-Service"), _T("ServiceParameters"));
      // Create service
      Install();
   }

   // Add object entries
   hr = CComModule::RegisterServer(bRegTypeLib);

   CoUninitialize();
   return hr;
}

inline HRESULT CServiceModule::UnregisterServer()
{
   HRESULT hr = CoInitialize(NULL);
   if (FAILED(hr))
      return hr;

   // Remove service entries
   UpdateRegistryFromResource(IDR_SMMonitor, FALSE);
   // Remove service
   Uninstall();
   // Remove object entries
   CComModule::UnregisterServer(TRUE);
   CoUninitialize();
   return S_OK;
}

inline void CServiceModule::Init(_ATL_OBJMAP_ENTRY* p, HINSTANCE h, UINT nServiceNameID, const GUID* plibid)
{
   CComModule::Init(p, h, plibid);

   m_bService = TRUE;
   strcpy(m_szServiceName, "SMMonitor");
   strcpy(m_szServiceLabel,"SORTmanager Map Management");


   // set up the initial service status 
   m_hServiceStatus = NULL;
   m_status.dwServiceType = SERVICE_WIN32_OWN_PROCESS;
   m_status.dwCurrentState = SERVICE_STOPPED;
   m_status.dwControlsAccepted = SERVICE_ACCEPT_STOP | SERVICE_ACCEPT_SHUTDOWN;
   m_status.dwWin32ExitCode = 0;
   m_status.dwServiceSpecificExitCode = 0;
   m_status.dwCheckPoint = 0;
   m_status.dwWaitHint = 0;
}

LONG CServiceModule::Unlock()
{
   LONG l = CComModule::Unlock();
   if (l == 0 && !m_bService)
      PostThreadMessage(dwThreadID, WM_QUIT, 0, 0);
   return l;
}

BOOL CServiceModule::IsInstalled()
{
   BOOL bResult = FALSE;

   SC_HANDLE hSCM = ::OpenSCManager(NULL, NULL, SC_MANAGER_ALL_ACCESS);
   if (hSCM != NULL)
   {
      SC_HANDLE hService = ::OpenService(hSCM, m_szServiceName,
               SERVICE_QUERY_CONFIG);
      if (hService != NULL)
      {
         bResult = TRUE;
         ::CloseServiceHandle(hService);
      }
      ::CloseServiceHandle(hSCM);
   }
   return bResult;
}

BOOL CServiceModule::Install()
{
   if (IsInstalled())
      return TRUE;

// check if this is a SORTmanager installation to determine if there should
// be a dependency on the sortWMMServer database service.

   BOOL smInstall = FALSE;

   try
   {
      CRegKey keySW;

      long lRes = keySW.Open(HKEY_LOCAL_MACHINE, _T("Software"), KEY_READ);
      if (lRes == ERROR_SUCCESS)
      {
         CRegKey keyCompany;

         lRes = keyCompany.Open(keySW, _T("Electroglas"), KEY_READ);
         if (lRes == ERROR_SUCCESS)
         {
            TCHAR szValue[_MAX_PATH+1];
            DWORD dwLen = _MAX_PATH;

            lRes = keyCompany.QueryValue(szValue, _T("InstallTopDir"), &dwLen);
            if (lRes == ERROR_SUCCESS)
            {
               char installIniFile[_MAX_PATH+1];

               strcpy(installIniFile,szValue);
               long ln = strtrimlen(installIniFile);
               if (ln > 0)
               {
                  if (! (installIniFile[ln-1] == '\\' ||
                         installIniFile[ln-1] == '/'))
                     strcat(installIniFile,"\\");

                  strcat(installIniFile,"EGInstallation.ini");
                  if (_access(installIniFile,0) == 0)
                     smInstall = TRUE;
               }
            }

            keyCompany.Close();   
         }
         keySW.Close();
      }
   }
   catch (...)
   {
      smInstall = FALSE;
   }

   char *dependArray = 0;
   char dbServer[] = "ASANYs_sortWMMServer";
   char rpcss[] = "RPCSS";

   dependArray = new char[strlen(dbServer) + strlen(rpcss) + 10];
   int i = 0, j;

   for (j = 0; j < (int) strlen(rpcss); j++)
      dependArray[i++] = rpcss[j];
   dependArray[i++] = '\0';

   if (smInstall)
   {
      for (j = 0; j < (int) strlen(dbServer); j++)
         dependArray[i++] = dbServer[j];
   }
   dependArray[i++] = '\0';
   dependArray[i++] = '\0';      // need an extra null

   SC_HANDLE hSCM = ::OpenSCManager(NULL, NULL, SC_MANAGER_ALL_ACCESS);
   if (hSCM == NULL)
   {
      LogEvent(EVENTLOG_ERROR_TYPE, _T("Couldn't open service manager"));
      return FALSE;
   }

   // Get the executable file path
   TCHAR szFilePath[_MAX_PATH];
   ::GetModuleFileName(NULL, szFilePath, _MAX_PATH);

   DWORD dwStartType = SERVICE_DEMAND_START;
   if (smInstall)
      dwStartType = SERVICE_AUTO_START;

   SC_HANDLE hService = ::CreateService(
       hSCM, m_szServiceName, m_szServiceLabel,
       SERVICE_ALL_ACCESS, SERVICE_WIN32_OWN_PROCESS,
       dwStartType, SERVICE_ERROR_NORMAL,
       szFilePath, NULL, NULL, dependArray, NULL, NULL);

   if (hService == NULL)
   {
      ::CloseServiceHandle(hSCM);
      LogEvent(EVENTLOG_ERROR_TYPE, _T("Couldn't create service"));
      return FALSE;
   }

   ::CloseServiceHandle(hService);
   ::CloseServiceHandle(hSCM);
   return TRUE;
}

inline BOOL CServiceModule::Uninstall()
{
   if (! IsInstalled())
      return TRUE;

   SC_HANDLE hSCM = ::OpenSCManager(NULL, NULL, SC_MANAGER_ALL_ACCESS);
   if (hSCM == NULL)
   {
      LogEvent(EVENTLOG_ERROR_TYPE, _T("Couldn't open service manager"));
      return FALSE;
   }

   SC_HANDLE hService = ::OpenService(hSCM, m_szServiceName, SERVICE_STOP | DELETE);
   if (hService == NULL)
   {
      ::CloseServiceHandle(hSCM);
      LogEvent(EVENTLOG_ERROR_TYPE, _T("Couldn't open service"));
      return FALSE;
   }

   SERVICE_STATUS status;
   ::ControlService(hService, SERVICE_CONTROL_STOP, &status);

   BOOL bDelete = ::DeleteService(hService);
   ::CloseServiceHandle(hService);
   ::CloseServiceHandle(hSCM);

   if (bDelete)
      return TRUE;

   LogEvent(EVENTLOG_ERROR_TYPE, _T("Service could not be deleted"));
   return FALSE;
}

// Logging functions
void CServiceModule::LogEvent(WORD evtType, LPCTSTR pFormat, ...)
{
   TCHAR chMsg[256];
   va_list pArg;

   va_start(pArg, pFormat);
   _vstprintf(chMsg, pFormat, pArg);
   va_end(pArg);

   EGEventLog eventLog("SORTMapManager");

   eventLog.postMessage(evtType, chMsg);
}

// Service startup and registration
inline void CServiceModule::Start()
{
   SERVICE_TABLE_ENTRY st[] =
   {
      { m_szServiceName, _ServiceMain },
      { NULL, NULL }
   };

   if (m_bService && !::StartServiceCtrlDispatcher(st))
   {
      m_bService = FALSE;
   }

   if (m_bService == FALSE)
      Run();
}

inline void CServiceModule::ServiceMain(DWORD, LPTSTR*)
{
   static const char *funcName = "CServiceModule::ServiceMain()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   // Register the control request handler
   m_status.dwCurrentState = SERVICE_START_PENDING;
   m_hServiceStatus = RegisterServiceCtrlHandler(m_szServiceName, _Handler);
   if (m_hServiceStatus == NULL)
   {
      LogEvent(EVENTLOG_ERROR_TYPE,_T("Handler not installed"));
      return;
   }

   SetServiceStatus(SERVICE_START_PENDING);
   LogEvent(EVENTLOG_INFORMATION_TYPE,
         _T("SORTmanager Map Management service is starting."));

   m_status.dwWin32ExitCode = S_OK;
   m_status.dwCheckPoint = 0;
   m_status.dwWaitHint = 0;

   // When the Run function returns, the service has stopped.
   Run();

   SetServiceStatus(SERVICE_STOPPED);
}

inline void CServiceModule::Handler(DWORD dwOpcode)
{
   static const char *funcName = "CServiceModule::Handler()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   switch (dwOpcode)
   {
      case SERVICE_CONTROL_STOP:
      case SERVICE_CONTROL_SHUTDOWN:
         SetServiceStatus(SERVICE_STOP_PENDING);
         LogEvent(EVENTLOG_INFORMATION_TYPE,
               _T("SORTmanager Map Management service is stopping."));
         PostThreadMessage(dwThreadID, WM_QUIT, 0, 0);
         m_inShutdown = true;
         break;

      case SERVICE_CONTROL_PAUSE:
         break;

      case SERVICE_CONTROL_CONTINUE:
         break;

      case SERVICE_CONTROL_INTERROGATE:
         break;

      default:
         LogEvent(EVENTLOG_ERROR_TYPE,_T("Bad service request"));
   }
}

void WINAPI CServiceModule::_ServiceMain(DWORD dwArgc, LPTSTR* lpszArgv)
{
   _Module.ServiceMain(dwArgc, lpszArgv);
}

void WINAPI CServiceModule::_Handler(DWORD dwOpcode)
{
   _Module.Handler(dwOpcode); 
}

void CServiceModule::SetServiceStatus(DWORD dwState)
{
   m_status.dwCurrentState = dwState;
   ::SetServiceStatus(m_hServiceStatus, &m_status);
}

class RunLocalPointers
{
   public:
      BOOL didCoInit;
      BOOL didRegClassObj;

      RunLocalPointers() : didCoInit(FALSE), didRegClassObj(FALSE) {}
      ~RunLocalPointers() { if (didRegClassObj) _Module.RevokeClassObjects();
                            if (didCoInit) CoUninitialize(); }
};

void CServiceModule::Run()
{
   static const char *funcName = "CServiceModule::Run()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);
   RunLocalPointers LP;

   _Module.dwThreadID = GetCurrentThreadId();

   HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
   if (FAILED(hr))
   {
      logErrorMsg(hr, funcName, "CoInitializeEx() failed.");
      return;
   }
   LP.didCoInit = TRUE;

   // This provides a NULL DACL which will allow access to everyone.
   CSecurityDescriptor sd;
   sd.InitializeFromThreadToken();
   hr = CoInitializeSecurity(sd, -1, NULL, NULL, RPC_C_AUTHN_LEVEL_PKT,
            RPC_C_IMP_LEVEL_IMPERSONATE, NULL, EOAC_NONE, NULL);
   if (FAILED(hr))
   {
      logErrorMsg(hr, funcName, "CoInitializeSecurity() failed.");
      return;
   }

   hr = _Module.RegisterClassObjects(CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER,
             REGCLS_MULTIPLEUSE);
   if (FAILED(hr))
   {
      logErrorMsg(hr, funcName, "RegisterClassObjects() failed.");
      return;
   }
   LP.didRegClassObj = TRUE;

// get registry information

   try
   {
      CRegKey keySW;

      logTraceLevel = logTraceLevelNone;
      logTraceFile[0] = '\0';

      long lRes = keySW.Open(HKEY_LOCAL_MACHINE, _T("Software"), KEY_READ);
      if (lRes != ERROR_SUCCESS)
      {
         logErrorMsg(errSmmRegistryError,funcName,"HKLM\\Software");
         return;
      }

      CRegKey keyCompany;

      lRes = keyCompany.Open(keySW, _T("Electroglas"), KEY_READ);
      if (lRes != ERROR_SUCCESS)         // try EGsoft
      {
         lRes = keyCompany.Open(keySW, _T("EGsoft"), KEY_READ);
         if (lRes != ERROR_SUCCESS)
         {
            logErrorMsg(errSmmRegistryError,funcName,
                  "HKLM\\Software\\Electroglas");
            return;
         }
      }

      CRegKey keySMM;

      lRes = keySMM.Open(keyCompany, _T("SORTMapManager"), KEY_READ);
      if (lRes != ERROR_SUCCESS)
      {
         logErrorMsg(errSmmRegistryError,funcName,
               "HKLM\\Software\\Electroglas\\SORTMapManager");
         return;
      }

      TCHAR szValue[_MAX_PATH+1];
      DWORD dwLen = _MAX_PATH;
      DWORD dwValue;

      lRes = keySMM.QueryValue(dwValue, _T("LogTraceLevel"));
      if (lRes == ERROR_SUCCESS)
      {
         if (dwValue >= 0)
            logTraceLevel = dwValue;
      }

      if (logTraceLevel > 0)
      {
         dwLen = _MAX_PATH;
         lRes = keySMM.QueryValue(szValue, _T("LogTraceFile"), &dwLen);
         if (lRes == ERROR_SUCCESS)
            strcpy(logTraceFile,szValue);
      }

      dwLen = _MAX_PATH;
      lRes = keySMM.QueryValue(szValue, _T("LoadMonitorData"), &dwLen);
      if (lRes != ERROR_SUCCESS)
      {
         logErrorMsg(errSmmRegistryError,funcName,
               "HKLM\\Software\\Electroglas\\SORTMapManager\\LoadMonitorData");
         return;
      }

      strcpy(m_loadMonitorData, szValue);

      if (logTraceLevel >= logTraceLevelDetail)
      {
         char localBuf[512];

         sprintf(localBuf,"LoadMonitorData = %s", m_loadMonitorData);
         logTrace(funcName, localBuf);
      }

      dwLen = _MAX_PATH;
      lRes = keySMM.QueryValue(szValue, _T("Dispatcher"), &dwLen);
      if (lRes == ERROR_SUCCESS)
      {
         strcpy(m_altDispatcher, szValue);
         m_useAltDispatcher = true;
      }

      lRes = keySMM.QueryValue(dwValue, _T("DispatchTimeout"));
      if (lRes == ERROR_SUCCESS)
      {
         if (dwValue > 0)
            m_dispatchTimeout = dwValue;
      }

      if (logTraceLevel >= logTraceLevelDetail)
      {
         char localBuf[128];

         sprintf(localBuf, "Dispatch timeout = %d", m_dispatchTimeout);
         logTrace(funcName, localBuf);
      }

      lRes = keySMM.QueryValue(dwValue, _T("ScavengeSurge"));
      if (lRes == ERROR_SUCCESS)
      {
         if (dwValue >= 0)
            m_scavengeSurgeTrigger = dwValue;
      }

      if (logTraceLevel >= logTraceLevelDetail)
      {
         char localBuf[128];

         sprintf(localBuf, "Scavenge surge trigger = %d",
               m_scavengeSurgeTrigger);
         logTrace(funcName, localBuf);
      }

      lRes = keySMM.QueryValue(dwValue, _T("ScavengeMax"));
      if (lRes == ERROR_SUCCESS)
      {
         if (dwValue > 0)
            m_scavengeMax = dwValue;
      }

      if (logTraceLevel >= logTraceLevelDetail)
      {
         char localBuf[128];

         sprintf(localBuf, "Scavenge maximum = %d", m_scavengeMax);
         logTrace(funcName, localBuf);
      }

      lRes = keySMM.QueryValue(dwValue, _T("UseServerPool"));
      if (lRes == ERROR_SUCCESS)
      {
         if (dwValue == 0)
            m_useServerPool = false;
         else
            m_useServerPool = true;
      }

      if (logTraceLevel >= logTraceLevelDetail)
      {
         if (m_useServerPool)
            logTrace(funcName, "UseServerPool = 1");
         else
            logTrace(funcName, "UseServerPool = 0");
      }

      lRes = keySMM.QueryValue(dwValue, _T("FailureReportInterval"));
      if (lRes == ERROR_SUCCESS)
      {
         if (dwValue > 0)
         {
            m_failureReportInterval = (double) dwValue / 86400.0;

            if (logTraceLevel >= logTraceLevelDetail)
            {
               char localBuf[128];

               sprintf(localBuf,"FailureReportInterval = %g (%ld)",
                     m_failureReportInterval, dwValue);
               logTrace(funcName, localBuf);
            }
         }
      }

      lRes = keySMM.QueryValue(dwValue, _T("LoadMonitorInterval"));
      if (lRes == ERROR_SUCCESS)
      {
         if (dwValue >= 60)
         {
// the interval is in seconds and a the timer is in milliseconds

            m_loadInterval = dwValue * 1000;
         }
         else
         {
            logErrorMsg(errSmmRegistryError, funcName,
               "LoadMonitorInterval must be at least 60 seconds.");
            return;
         }
      }

      keySMM.Close();
      keyCompany.Close();
      keySW.Close();   
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
      return;
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
      return;
   }

   LogEvent(EVENTLOG_INFORMATION_TYPE,
         _T("SORTmanager Map Management service started"));
   if (m_bService)
      SetServiceStatus(SERVICE_RUNNING);

#ifdef _DEBUG
   Sleep(25000);      // attach a debugger here
#endif

   bool lmdFailed = false; 

   try
   {
// load the prober information

      if (loadMonitorData() == Error_Failure)
      {
         logErrorMsg(errSmmWarning, 0,
            "Failed to load the monitor data.");
         lmdFailed = true;
      }

      if (createTimers() == Error_Failure)
         return;

// start the dispatcher

      if (m_dispatcher.setup(m_useAltDispatcher, m_altDispatcher,
            m_dispatchTimeout, m_useServerPool) == Error_Failure)
         return;

      if (m_dispatcher.start() == Error_Failure)
         return;

      bool done = false;

      while (! done)
      {
         DWORD waitResult = MsgWaitForMultipleObjectsEx(0, NULL,
                  MAIN_LOOP_TIMEOUT, QS_ALLINPUT, MWMO_ALERTABLE);
         if (waitResult == WAIT_OBJECT_0)
         {
            MSG msg;
            memset(&msg, 0, sizeof(MSG));

            while ((! done) && PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
            {
               if (msg.message == WM_QUIT)
               {
                  done = true;            // service stop or shutdown
               }
               else if (msg.message == WM_TIMER)
               {
                  if (msg.wParam == m_loadTimer)
                  {
                     if (loadMonitorData() == Error_Success)
                     {
                        if (lmdFailed)
                        {
                           lmdFailed = false;
                           logErrorMsg(errSmmWarning, 0,
                              "The monitor data were successfully loaded.");
                        }
                     }
                  }
               }
               else
               {
                  DispatchMessage(&msg);
               }
            }
         }
         else if (waitResult == WAIT_TIMEOUT)   // check the monitors
         {
            checkMonitors();
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

   if (logTraceLevel >= logTraceLevelProfile)
   {
      logTrace(funcName, "Leaving main message loop");
   }

// redundant, but if we got here by error

   setShutdown(true);

   try
   {
      stopTimers();

// stop the dispatcher

      m_dispatcher.shutdown();

// stop all monitor threads

      m_monitorList.setCapacity(0);
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   if (logTraceLevel >= logTraceLevelProfile)
   {
      logTrace(funcName, "Leaving CServiceModule::Run()");
   }
}

// determine the folders to be monitored

class LoadMonitorLocalPointers
{
   public:
      char **strArray;
      char **extArray;
      char *strValue;
      int *intArray;
      double *dblArray;
      unsigned long nStrings;
      unsigned long nExts;
      HANDLE hDir;

      LoadMonitorLocalPointers() : strArray(0), nStrings(0), intArray(0),
                                   strValue(0), extArray(0), nExts(0),
                                   dblArray(0), hDir(0) {}
      ~LoadMonitorLocalPointers() { delete intArray; delete strValue;
                                    delete dblArray;
                         if (strArray) {
                            for (unsigned long i = 0; i < nStrings; i++)
                               delete strArray[i];
                            delete strArray;
                         }
                         if (extArray) {
                            for (unsigned long i = 0; i < nExts; i++)
                               delete extArray[i];
                            delete extArray;
                         }
                         if (hDir) CloseHandle(hDir);
                      }
};


long CServiceModule::loadMonitorData()
{
   static const char *funcName = "CServiceModule::loadMonitorData()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   LoadMonitorLocalPointers LP;
   unsigned long i, j;

   long retval = Error_Failure;

   if (m_inShutdown)
      return retval;

   try
   {
      eServer statitLib(m_useServerPool);

      if (statitLib.init() == Error_Failure)
         return retval;

// the loadMonitorData buffer contains the proc or macro to execute.  the
// SMM proc, sw_getSMMData, will be used to execute it and validate the data
// global StatitScript variables are used to exchange run information and
// the workspace must contain the dataset Monitor.

      if (statitLib.putGlobalString("%SMM_LoadMonitorData", m_loadMonitorData)
               == Error_Failure)
         return retval;

      if (logTraceLevel >= logTraceLevelDetail)
      {
         char localBuf[512];

         sprintf(localBuf,"Loading monitor data from '%s'.", m_loadMonitorData);
         logTrace(funcName, localBuf);
      }

      if (statitLib.runCmd("smm_getMonitorData") == Error_Failure)
         return retval;

      long numDups = 0;
      if (statitLib.getGlobalInt("$SMM_MonitorDupCount", numDups) ==
               Error_Failure)
         return retval;

      if (numDups > 0)            // issue a warning
      {
         char localBuf[1024];

         if (statitLib.getGlobalString("%SMM_MonitorDups", LP.strValue) ==
                     Error_Failure)
            return retval;

         sprintf(localBuf,"Duplicate monitor id's are ignored: %s",
               LP.strValue);
         delete LP.strValue;
         LP.strValue = 0;      // for dtor

         logErrorMsg(errSmmWarning, 0, localBuf);
      }

// load the monitor.id and monitor.dirpath data

      if (statitLib.getStringVar("Monitor.Id", LP.nStrings, LP.strArray) ==
                  Error_Failure)
         return retval;

      MonitorList newList;
      unsigned long nDirs = 0;

      if (LP.nStrings > 0)
      {
         nDirs = LP.nStrings;

         if (newList.setLength(nDirs) != 0)
         {
            logErrorMsg(errSmmResource, funcName,
               "Failed to allocate memory for the monitor list.");
            return retval;
         }

         for (i = 0; i < nDirs; i++)
         {
            newList[i].monitorId = LP.strArray[i];
            LP.strArray[i] = 0;   // for dtor
         }

         delete LP.strArray;
         LP.strArray = 0;         // for dtor
         LP.nStrings = 0;

         if (statitLib.getStringVar("Monitor.DirPath", LP.nStrings,
                  LP.strArray) == Error_Failure)
            return retval;

         for (i = 0; i < nDirs; i++)
         {
            newList[i].dirPath = LP.strArray[i];
            LP.strArray[i] = 0;   // for dtor
         }

         delete LP.strArray;
         LP.strArray = 0;         // for dtor
         LP.nStrings = 0;

         bool varExists = false;
         if (statitLib.doesVarExist("Monitor.Extensions", varExists) ==
                  Error_Failure)
            return retval;

         if (varExists)
         {
            if (statitLib.getStringVar("Monitor.Extensions", LP.nStrings,
                     LP.strArray) == Error_Failure)
               return retval;

// parse the extensions (comma delimited), may be 0 or more per directory

            for (i = 0; i < nDirs; i++)
            {
               unsigned long numExts = 0;
               unsigned long ls = 0;

               char *pExt = LP.strArray[i];
               if (pExt)
                  ls = strlen(pExt);

               if (ls > 0)
               {
                  numExts = 1;      // at least one
                  long nSqrBracket = 0;
                  long nCurlyBracket = 0;

                  for (j = 0; j < ls; j++)
                  {
                     if (pExt[j] == ',' && nSqrBracket == 0 &&
                              nCurlyBracket == 0)
                     {
                        numExts++;
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
               }

               if (numExts > 0)
               {
                  LP.extArray = new char *[numExts];
                  if (LP.extArray == 0)
                  {
                     logErrorMsg(errSmmResource, funcName,
                        "Failed to allocate memory for the extensions.");
                     return retval;
                  }

                  for (j = 0; j < numExts; j++)
                     LP.extArray[j] = 0;

                  LP.nExts = numExts;
                  unsigned long k = 0;
                  unsigned long l = 0;
                  unsigned long ln = 0;

                  long nSqrBracket = 0;
                  long nCurlyBracket = 0;

                  for (j = 0; j < ls; j++)
                  {
                     if (pExt[j] == ',' && nSqrBracket == 0 &&
                              nCurlyBracket == 0)
                     {
                        ln = j - l;
                        LP.strValue = new char[ln+1];
                        if (LP.strValue == 0)
                        {
                           logErrorMsg(errSmmResource, funcName,
                              "Failed to allocate memory for the extensions.");
                           return retval;
                        }

                        strncpy(LP.strValue,pExt+l,ln);
                        LP.strValue[ln] = '\0';
                        LP.extArray[k++] = LP.strValue;
                        LP.strValue = 0;      // for dtor

                        l = j + 1;
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

                  ln = ls - l;
                  if (ln > 0)
                  {
                     LP.strValue = new char[ln+1];
                     if (LP.strValue == 0)
                     {
                        logErrorMsg(errSmmResource, funcName,
                           "Failed to allocate memory for the extensions.");
                        return retval;
                     }

                     strncpy(LP.strValue,pExt+l,ln);
                     LP.strValue[ln] = '\0';
                     LP.extArray[k++] = LP.strValue;
                     LP.strValue = 0;      // for dtor
                  }

                  newList[i].extensions = LP.extArray;
                  newList[i].numExtensions = LP.nExts;

                  LP.extArray = 0;        // for dtor
                  LP.nExts = 0;
               }

               delete pExt;
               LP.strArray[i] = 0;        // for dtor
            }

            delete LP.strArray;
            LP.strArray = 0;              // for dtor
         }

// lotDir

         varExists = false;
         if (statitLib.doesVarExist("Monitor.LotDir", varExists) ==
                  Error_Failure)
            return retval;

         if (varExists)
         {
            if (statitLib.getStringVar("Monitor.LotDir", LP.nStrings,
                     LP.strArray) == Error_Failure)
               return retval;

            for (i = 0; i < nDirs; i++)
            {
               newList[i].lotDir = LP.strArray[i];
               LP.strArray[i] = 0;   // for dtor
            }

            delete LP.strArray;
            LP.strArray = 0;         // for dtor
            LP.nStrings = 0;
         }

// verifyMap

         varExists = false;
         if (statitLib.doesVarExist("Monitor.VerifyMap", varExists) ==
                  Error_Failure)
            return retval;

         if (varExists)
         {
            if (statitLib.getStringVar("Monitor.VerifyMap", LP.nStrings,
                     LP.strArray) == Error_Failure)
               return retval;

            for (i = 0; i < nDirs; i++)
            {
               newList[i].verifyMap = LP.strArray[i];
               LP.strArray[i] = 0;   // for dtor
            }

            delete LP.strArray;
            LP.strArray = 0;         // for dtor
            LP.nStrings = 0;
         }

// mapAction

         varExists = false;
         if (statitLib.doesVarExist("Monitor.MapAction", varExists) ==
                  Error_Failure)
            return retval;

         if (varExists)
         {
            if (statitLib.getStringVar("Monitor.MapAction", LP.nStrings,
                     LP.strArray) == Error_Failure)
               return retval;

            for (i = 0; i < nDirs; i++)
            {
               newList[i].mapAction = LP.strArray[i];
               LP.strArray[i] = 0;   // for dtor
            }

            delete LP.strArray;
            LP.strArray = 0;         // for dtor
            LP.nStrings = 0;
         }

// verifyLot

         varExists = false;
         if (statitLib.doesVarExist("Monitor.VerifyLot", varExists) ==
                  Error_Failure)
            return retval;

         if (varExists)
         {
            if (statitLib.getStringVar("Monitor.VerifyLot", LP.nStrings,
                     LP.strArray) == Error_Failure)
               return retval;

            for (i = 0; i < nDirs; i++)
            {
               newList[i].verifyLot = LP.strArray[i];
               LP.strArray[i] = 0;   // for dtor
            }

            delete LP.strArray;
            LP.strArray = 0;         // for dtor
            LP.nStrings = 0;
         }

// lotAction

         varExists = false;
         if (statitLib.doesVarExist("Monitor.LotAction", varExists) ==
                  Error_Failure)
            return retval;

         if (varExists)
         {
            if (statitLib.getStringVar("Monitor.LotAction", LP.nStrings,
                     LP.strArray) == Error_Failure)
               return retval;

            for (i = 0; i < nDirs; i++)
            {
               newList[i].lotAction = LP.strArray[i];
               LP.strArray[i] = 0;   // for dtor
            }

            delete LP.strArray;
            LP.strArray = 0;         // for dtor
            LP.nStrings = 0;
         }

// scavenge

         varExists = false;
         if (statitLib.doesVarExist("Monitor.Scavenge", varExists) ==
                  Error_Failure)
            return retval;

         if (varExists)
         {
            if (statitLib.getStringVar("Monitor.Scavenge", LP.nStrings,
                     LP.strArray) == Error_Failure)
               return retval;

            for (i = 0; i < nDirs; i++)
            {
               newList[i].scavenge = LP.strArray[i];
               LP.strArray[i] = 0;   // for dtor
            }

            delete LP.strArray;
            LP.strArray = 0;         // for dtor
            LP.nStrings = 0;
         }

// scavengeTime

         varExists = false;
         if (statitLib.doesVarExist("Monitor.ScavengeTime", varExists) ==
                  Error_Failure)
            return retval;

         if (varExists)
         {
            unsigned long nTimes;

            if (statitLib.getDoubleVar("Monitor.ScavengeTime", nTimes,
                     LP.dblArray) == Error_Failure)
               return retval;

            for (i = 0; i < nTimes; i++)
            {
               newList[i].scavengeTime = (long) LP.dblArray[i];
            }

            delete LP.dblArray;
            LP.dblArray = 0;         // for dtor
         }

// scavengeStart

         varExists = false;
         if (statitLib.doesVarExist("Monitor.ScavengeStart", varExists) ==
                  Error_Failure)
            return retval;

         if (varExists)
         {
            unsigned long nTimes;

            if (statitLib.getDoubleVar("Monitor.ScavengeStart", nTimes,
                     LP.dblArray) == Error_Failure)
               return retval;

            for (i = 0; i < nTimes; i++)
            {
               newList[i].scavengeStart = (long) LP.dblArray[i];
            }

            delete LP.dblArray;
            LP.dblArray = 0;         // for dtor
         }
      }

// check the results against items currently being monitored.  if there
// are new ones, add them, if there are some removed, stop the monitoring
// this is a linear search, but the list should never be more than a
// few hundred items so it should be sufficient

// first, check the existing items and if it is no longer in the list,
// stop monitoring.  if the id exists but the directory path or other
// items have changed then it is considered to be new and the old one
// is released

      if (m_inShutdown)
         return retval;

      if (nDirs > 0)
      {
         LP.intArray = new int[nDirs];
         if (LP.intArray == 0)
         {
            logErrorMsg(errSmmResource, funcName,
               "Failed to allocate memory for the directory path.");
            return retval;
         }

         for (i = 0; i < nDirs; i++)
            LP.intArray[i] = 1;
      }

      for (i = 0; i < m_monitorList.length(); i++)
      {
         if (m_monitorList[i].inUse)
         {
            // check that the monitoring is still going okay, if not, shut
            // it down

            bool clearSlot = false;

            if (m_monitorList[i].monitorShare)
            {
               if (m_monitorList[i].monitorShare->alive() == false)
                  clearSlot = true;
               else if (m_monitorList[i].monitorShare->requestShutdown())
                  clearSlot = true;
            }
            else
            {
               clearSlot = true;
            }

            if (clearSlot)
               clearMonitorSlot(i);
         }

         // test again since it may have been cleared above

         if (m_monitorList[i].inUse)
         {
            char *mName = m_monitorList[i].monitorId;
            char *mDirPath = m_monitorList[i].dirPath;
            bool foundId = false;
            bool dataChanged = false;


            for (j = 0; j < nDirs; j++)
            {
               if (stricmp(mName, newList[j].monitorId) == 0)
               {
                  foundId = true;

                  if (stricmp(mDirPath, newList[j].dirPath) != 0)
                     dataChanged = true;

                  if (! dataChanged)
                  {
                     if (m_monitorList[i].numExtensions !=
                                 newList[j].numExtensions)
                        dataChanged = true;
                  }

                  if (! dataChanged)
                  {
                     if (m_monitorList[i].extensions && newList[j].extensions)
                     {
                        for (unsigned long k = 0;
                                 k < m_monitorList[i].numExtensions; k++)
                        {
                           if (stricmp(m_monitorList[i].extensions[k],
                                 newList[j].extensions[k]) != 0)
                           {
                              dataChanged = true;
                              break;
                           }
                        }
                     }
                     else if (m_monitorList[i].extensions ||
                                 newList[j].extensions)
                     {
                        dataChanged = true;
                     }
                  }

                  if (! dataChanged)
                  {
                     if (m_monitorList[i].lotDir && newList[j].lotDir)
                     {
                        if (stricmp(m_monitorList[i].lotDir,
                               newList[j].lotDir) != 0)
                           dataChanged = true;
                     }
                     else if (m_monitorList[i].lotDir || newList[j].lotDir)
                     {
                        dataChanged = true;
                     }
                  }

                  if (! dataChanged)
                  {
                     if (m_monitorList[i].verifyMap && newList[j].verifyMap)
                     {
                        if (stricmp(m_monitorList[i].verifyMap,
                               newList[j].verifyMap) != 0)
                           dataChanged = true;
                     }
                     else if (m_monitorList[i].verifyMap ||
                                 newList[j].verifyMap)
                     {
                        dataChanged = true;
                     }
                  }

                  if (! dataChanged)
                  {
                     if (m_monitorList[i].mapAction && newList[j].mapAction)
                     {
                        if (stricmp(m_monitorList[i].mapAction,
                               newList[j].mapAction) != 0)
                           dataChanged = true;
                     }
                     else if (m_monitorList[i].mapAction ||
                                 newList[j].mapAction)
                     {
                        dataChanged = true;
                     }
                  }

                  if (! dataChanged)
                  {
                     if (m_monitorList[i].verifyLot && newList[j].verifyLot)
                     {
                        if (stricmp(m_monitorList[i].verifyLot,
                               newList[j].verifyLot) != 0)
                           dataChanged = true;
                     }
                     else if (m_monitorList[i].verifyLot ||
                                 newList[j].verifyLot)
                     {
                        dataChanged = true;
                     }
                  }

                  if (! dataChanged)
                  {
                     if (m_monitorList[i].lotAction && newList[j].lotAction)
                     {
                        if (stricmp(m_monitorList[i].lotAction,
                               newList[j].lotAction) != 0)
                           dataChanged = true;
                     }
                     else if (m_monitorList[i].lotAction ||
                                 newList[j].lotAction)
                     {
                        dataChanged = true;
                     }
                  }

                  if (! dataChanged)
                  {
                     if (m_monitorList[i].scavenge && newList[j].scavenge)
                     {
                        if (stricmp(m_monitorList[i].scavenge,
                               newList[j].scavenge) != 0)
                           dataChanged = true;
                     }
                     else if (m_monitorList[i].scavenge ||
                                 newList[j].scavenge)
                     {
                        dataChanged = true;
                     }
                  }

                  if (! dataChanged)
                  {
                     if (m_monitorList[i].scavengeTime !=
                           newList[j].scavengeTime)
                        dataChanged = true;
                  }

                  if (! dataChanged)
                  {
                     if (m_monitorList[i].scavengeStart !=
                           newList[j].scavengeStart)
                        dataChanged = true;
                  }
               }

               if (foundId == true && dataChanged == false)
               {
                  LP.intArray[j] = 0;
                  break;
               }
            }

            if (foundId == false || dataChanged == true)
            {
               clearMonitorSlot(i);
            }
         }
      }

      if (m_inShutdown)
         return retval;

// at this point, the items in the list that are not being monitored or that
// have changed will have the intArray flag set to -1.  if none are -1 then
// we are done.

      unsigned long newItems = 0;
      for (j = 0; j < nDirs; j++)
      {
         if (LP.intArray[j] > 0)
            newItems++;
      }

      if (newItems > 0)
      {

// first, make sure the directory can be monitored.  to do so make sure that
// it is exists, is a directory, and we have access

         for (j = 0; j < nDirs; j++)
         {
            if (LP.intArray[j] > 0)
            {
               if (m_inShutdown)
                  return retval;

               DWORD eCode = 0;
               bool okayToMonitor = false;
               bool notADir = false;

               DWORD fileAttr = GetFileAttributes(newList[j].dirPath);
               if (fileAttr == -1)
               {
                  eCode = GetLastError();
               }
               else
               {
                  if (fileAttr & FILE_ATTRIBUTE_DIRECTORY)
                  {
                     LP.hDir = CreateFile(newList[j].dirPath,
                        FILE_LIST_DIRECTORY,
                        FILE_SHARE_READ | FILE_SHARE_DELETE | FILE_SHARE_WRITE,
                        0, OPEN_EXISTING,
                        FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OVERLAPPED, 0);
                     if (LP.hDir == INVALID_HANDLE_VALUE)
                     {
                        eCode = GetLastError();
                     }
                     else
                     {
                        CloseHandle(LP.hDir);
                        LP.hDir = 0;            // for dtor

                        okayToMonitor = true;
                     }
                  }
                  else
                  {
                     notADir = true;
                  }
               }

               if (okayToMonitor)
               {
                  removeFromFailureList(newList[j].monitorId,
                        newList[j].dirPath);   // remove (if necessary)
               }
               else                              // add to failure list
               {
                  LP.intArray[j] = 0;
                  addToFailureList(newList[j].monitorId, newList[j].dirPath,
                     eCode, notADir);
               }
            }
         }

// next, report failures (as needed)

         reportMonitorFailures();

// finally, go through the list again and set up the monitoring

         for (j = 0; j < nDirs; j++)
         {
            if (LP.intArray[j] > 0)
            {
               if (m_inShutdown)
                  return retval;

               long slot = findMonitorSlot();
               if (slot < 0)
                  return retval;

               m_monitorList[slot].monitorId = newList[j].monitorId;
               newList[j].monitorId = 0;

               m_monitorList[slot].dirPath = newList[j].dirPath;
               newList[j].dirPath = 0;

               m_monitorList[slot].numExtensions = newList[j].numExtensions;
               m_monitorList[slot].extensions = newList[j].extensions;
               newList[j].extensions = 0;

               m_monitorList[slot].lotDir = newList[j].lotDir;
               newList[j].lotDir = 0;

               m_monitorList[slot].verifyMap = newList[j].verifyMap;
               newList[j].verifyMap = 0;

               m_monitorList[slot].mapAction = newList[j].mapAction;
               newList[j].mapAction = 0;

               m_monitorList[slot].verifyLot = newList[j].verifyLot;
               newList[j].verifyLot = 0;

               m_monitorList[slot].lotAction = newList[j].lotAction;
               newList[j].lotAction = 0;

               m_monitorList[slot].scavenge = newList[j].scavenge;
               newList[j].scavenge = 0;

               m_monitorList[slot].scavengeTime = newList[j].scavengeTime;
               m_monitorList[slot].scavengeStart = newList[j].scavengeStart;

               m_monitorList[slot].monitorShare = new MonitorShare;
               if (m_monitorList[slot].monitorShare == 0)
               {
                  logErrorMsg(errSmmResource, funcName,
                        "MonitorShare allocation");
                  return retval;
               }

               if (m_monitorList[slot].monitorShare->setup(
                     m_monitorList[slot].monitorId,
                     m_monitorList[slot].dirPath,
                     m_monitorList[slot].numExtensions,
                     (const char **) m_monitorList[slot].extensions,
                     m_monitorList[slot].lotDir,
                     m_monitorList[slot].verifyMap,
                     m_monitorList[slot].mapAction,
                     m_monitorList[slot].verifyLot,
                     m_monitorList[slot].lotAction,
                     m_monitorList[slot].scavenge,
                     m_monitorList[slot].scavengeTime,
                     m_monitorList[slot].scavengeStart,
                     m_scavengeSurgeTrigger, m_useServerPool) ==
                        Error_Success)
               {
                  if (m_monitorList[slot].monitorShare->start() ==
                           Error_Success)
                     m_monitorList[slot].inUse = true;
               }

               if (m_monitorList[slot].inUse == false)
                  clearMonitorSlot(slot);
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

long CServiceModule::findMonitorSlot()
{
   static const char *funcName = "CServiceModule::findMonitorSlot()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long openSlot = Error_Failure;

   for (unsigned long i = 0; i < m_monitorList.length(); i++)
   {
      if (m_monitorList[i].inUse == false)
      {
         openSlot = (long) i;
         break;
      }
   }

   if (openSlot < 0)
   {
      openSlot = m_monitorList.length();
      if (m_monitorList.setLength(openSlot+1) < 0)
      {
         logErrorMsg(errSmmResource, funcName, "Monitor list allocation");
         openSlot = Error_Failure;
      }
   }

   return openSlot;
}

long CServiceModule::clearMonitorSlot(unsigned long slot)
{
   static const char *funcName = "CServiceModule::clearMonitorSlot()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;
   if (slot >= 0 && slot < m_monitorList.length())
   {
      if (m_monitorList[slot].inUse)
      {
         if (m_monitorList[slot].monitorShare)
         {
            if (logTraceLevel >= logTraceLevelDetail)
            {
               char localBuf[512];

               sprintf(localBuf,"Stopping monitoring %s '%s'.",
                  m_monitorList[slot].monitorId,
                  m_monitorList[slot].dirPath);
               logTrace(funcName, localBuf);
            }

            m_monitorList[slot].monitorShare->shutdown();
            delete m_monitorList[slot].monitorShare;
            m_monitorList[slot].monitorShare = 0;
         }

         delete m_monitorList[slot].monitorId;
         m_monitorList[slot].monitorId = 0;

         delete m_monitorList[slot].dirPath;
         m_monitorList[slot].dirPath = 0;

         for (unsigned long i = 0; i < m_monitorList[slot].numExtensions; i++)
         {
            delete m_monitorList[slot].extensions[i];
            m_monitorList[slot].extensions[i] = 0;
         }

         delete [] m_monitorList[slot].extensions;
         m_monitorList[slot].extensions = 0;

         delete m_monitorList[slot].lotDir;
         m_monitorList[slot].lotDir = 0;

         delete m_monitorList[slot].verifyMap;
         m_monitorList[slot].verifyMap = 0;

         delete m_monitorList[slot].verifyLot;
         m_monitorList[slot].verifyLot = 0;

         m_monitorList[slot].scavengeTime = defaultScavengeTime;
         m_monitorList[slot].scavengeStart = defaultScavengeStart;
         m_monitorList[slot].inUse = false;
      }

      retval = Error_Success;
   }

   return retval;
}

void CServiceModule::removeFromFailureList(const char *monitorId,
            const char *dirPath)
{
   static const char *funcName = "CServiceModule::removeFromFailureList()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   for (unsigned long i = 0; i < m_failureList.length(); i++)
   {
      if (m_failureList[i].inUse)
      {
         if (stricmp(m_failureList[i].monitorId, monitorId) == 0)
         {
            if (stricmp(m_failureList[i].dirPath, dirPath) == 0)
            {
               delete m_failureList[i].monitorId;
               m_failureList[i].monitorId = 0;

               delete m_failureList[i].dirPath;
               m_failureList[i].dirPath = 0;

               m_failureList[i].eCode = 0;
               m_failureList[i].notADir = false;
               m_failureList[i].inUse = false;
               return;
            }
         }
      }
   }
}

class StringLocalPointers
{
   public:
      char *string;

      StringLocalPointers() : string(0) {}
      ~StringLocalPointers() { delete string; }
};

void CServiceModule::addToFailureList(const char *monitorId,
         const char *dirPath, DWORD eCode, bool notADir)
{
   static const char *funcName = "CServiceModule::addToFailureList()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

// first, check to see if it is already in the failure list (linear, I know)

   for (unsigned long i = 0; i < m_failureList.length(); i++)
   {
      if (m_failureList[i].inUse)
      {
         if (stricmp(m_failureList[i].monitorId, monitorId) == 0)
         {
            if (stricmp(m_failureList[i].dirPath, dirPath) == 0)
            {
               return;         // already there, we are done
            }
         }
      }
   }

   long slot = findFailureSlot();
   if (slot >= 0)
   {
      StringLocalPointers LP;

      unsigned long ls = strlen(monitorId);
      LP.string = new char[ls+1];
      if (LP.string == 0)
      {
         logErrorMsg(errSmmResource, funcName,
            "Failed to allocate memory for the monitor id.");
         return;
      }

      strcpy(LP.string, monitorId);
      m_failureList[slot].monitorId = LP.string;
      LP.string = 0;

      ls = strlen(dirPath);
      LP.string = new char[ls+1];
      if (LP.string == 0)
      {
         logErrorMsg(errSmmResource, funcName,
            "Failed to allocate memory for the directory path.");
         return;
      }

      strcpy(LP.string, dirPath);
      m_failureList[slot].dirPath = LP.string;
      LP.string = 0;

      m_failureList[slot].eCode = eCode;
      m_failureList[slot].notADir = notADir;
      m_failureList[slot].inUse = true;
   }
}

void CServiceModule::reportMonitorFailures()
{
   static const char *funcName = "CServiceModule::reportMonitorFailures()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   BCharVec monitorMsg;
   char newLine[2];
   SYSTEMTIME sTime;
   double vTime;
   bool somethingToReport = false;

   monitorMsg.setLength(0);
   newLine[0] = '\n';
   newLine[1] = '\0';

   GetSystemTime(&sTime);
   SystemTimeToVariantTime(&sTime, &vTime);

   veccat(monitorMsg, "The following could not be monitored:");
   veccat(monitorMsg, newLine);

   for (unsigned long i = 0; i < m_failureList.length(); i++)
   {
      if (m_failureList[i].inUse)
      {
         if (vTime - m_failureList[i].timeStamp >= m_failureReportInterval)
         {
            veccat(monitorMsg, newLine);
            veccat(monitorMsg, m_failureList[i].monitorId);
            veccat(monitorMsg, " '");
            veccat(monitorMsg, m_failureList[i].dirPath);
            veccat(monitorMsg, "' ");

            if (m_failureList[i].notADir)
            {
               veccat(monitorMsg,"(Not a directory.)");
            }
            else
            {
               char errBuf[1024];

               long lnErrBuf = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, 
                        m_failureList[i].eCode, 0, errBuf, 1024, 0);
               if (lnErrBuf == 0)
               {
                  sprintf(errBuf,"Error code: %ld.", m_failureList[i].eCode);
               }
               else   // strip off crlf
               {
                  if (errBuf[lnErrBuf-1] == '\r' || errBuf[lnErrBuf-1] == '\n')
                     lnErrBuf--;
                  if (errBuf[lnErrBuf-1] == '\r' || errBuf[lnErrBuf-1] == '\n')
                     lnErrBuf--;
                  errBuf[lnErrBuf] = '\0';
               }

               veccat(monitorMsg, "(");
               veccat(monitorMsg, errBuf);
               veccat(monitorMsg, ")");
            }

            m_failureList[i].timeStamp = vTime;
            somethingToReport = true;
         }
      }
   }

   if (somethingToReport)
   {
      EGEventLog eventLog("SORTMapManager");

      eventLog.postMessage(EVENTLOG_WARNING_TYPE, cp(monitorMsg));

      if (logTraceLevel)
         logTrace(funcName, cp(monitorMsg));
   }
}


long CServiceModule::findFailureSlot()
{
   static const char *funcName = "CServiceModule::findFailureSlot()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long openSlot = Error_Failure;

   for (unsigned long i = 0; i < m_failureList.length(); i++)
   {
      if (m_failureList[i].inUse == false)
      {
         openSlot = (long) i;
         break;
      }
   }

   if (openSlot < 0)
   {
      openSlot = m_failureList.length();
      if (m_failureList.setLength(openSlot+1) < 0)
      {
         logErrorMsg(errSmmResource, funcName,
               "Monitor failure list allocation");
         openSlot = Error_Failure;
      }
   }

   return openSlot;
}

void CServiceModule::setShutdown(bool b)
{
   static const char *funcName = "CServiceModule::setShutdown()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   m_inShutdown = b;
}


bool CServiceModule::inShutdown()
{
   static const char *funcName = "CServiceModule::inShutdown()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   return m_inShutdown;
}


bool CServiceModule::okayToScavenge(unsigned long& retryTime)
{
   static const char *funcName = "CServiceModule::okayToScavenge()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   bool retval = false;
   retryTime = 0;

   EnterCriticalSection(&m_scavengeLock);

   if (m_scavengeCount < m_scavengeMax)
   {
      retval = true;
      m_scavengeCount++;
   }
   else
   {
      double randU01 = ((double) rand()) / ((double) RAND_MAX);
      retryTime = (unsigned long) ((randU01 * scavengeRetryMax) + 0.5) + 100;
   }
   
   LeaveCriticalSection(&m_scavengeLock);

   return retval;
}

void CServiceModule::scavengeDone()
{
   static const char *funcName = "CServiceModule::scavengeDone()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   EnterCriticalSection(&m_scavengeLock);

   if (m_scavengeCount > 0)
   {
      m_scavengeCount--;
   }
   
   LeaveCriticalSection(&m_scavengeLock);
}



// helper function, could be global but instead will toss it here

int CServiceModule::veccat (BCharVec& v, const char *s)
{
   int i = 0;
   int k = v.length() - 1;
   int ls = strlen(s);

   if (k < 0)
      k = 0;

   int retVal = v.setLength(k + ls + 1);
   if (retVal)
      return (retVal);

   char *p = cp(v);

   while (s[i] != '\0')
   {
      p[k] = s[i];
      k++; i++;
   }
   p[k] = '\0';
   return 0;
}


extern "C" int WINAPI _tWinMain(HINSTANCE hInstance, HINSTANCE,
               LPTSTR lpCmdLine, int)
{
   lpCmdLine = GetCommandLine(); //this line necessary for _ATL_MIN_CRT
   _Module.Init(ObjectMap, hInstance, IDS_SERVICENAME, &LIBID_SMMonitor);
   _Module.m_bService = TRUE;

   TCHAR szTokens[] = _T("-/");

   LPCTSTR lpszToken = FindOneOf(lpCmdLine, szTokens);
   while (lpszToken != NULL)
   {
      if (lstrcmpi(lpszToken, _T("UnregServer"))==0)
         return _Module.UnregisterServer();

      // Register as Local Server
      if (lstrcmpi(lpszToken, _T("RegServer"))==0)
         return _Module.RegisterServer(TRUE, FALSE);
        
      // Register as Service
      if (lstrcmpi(lpszToken, _T("Service"))==0)
         return _Module.RegisterServer(TRUE, TRUE);
        
      lpszToken = FindOneOf(lpszToken, szTokens);
   }

    // Are we Service or Local Server
   CRegKey keyAppID;
   LONG lRes = keyAppID.Open(HKEY_CLASSES_ROOT, _T("AppID"), KEY_READ);
   if (lRes != ERROR_SUCCESS)
      return lRes;

   CRegKey key;
   lRes = key.Open(keyAppID, _T("{53D6D93C-69F1-4D90-B7A0-8A4148F2AC1D}"),
               KEY_READ);
   if (lRes != ERROR_SUCCESS)
      return lRes;

   TCHAR szValue[_MAX_PATH];
   DWORD dwLen = _MAX_PATH;
   lRes = key.QueryValue(szValue, _T("LocalService"), &dwLen);

   _Module.m_bService = FALSE;
   if (lRes == ERROR_SUCCESS)
      _Module.m_bService = TRUE;

   _Module.Start();

   // When we get here, the service has been stopped
   return _Module.m_status.dwWin32ExitCode;
}
long CServiceModule::checkMonitors()
{
   static const char *funcName = "CServiceModule::checkMonitors()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   try
   {
      for (unsigned i = 0; i < m_monitorList.length(); i++)
      {
         if (m_inShutdown)
            return retval;

         if (m_monitorList[i].inUse)
         {
            // check that the monitoring is still going okay, if not, shut
            // it down

            bool clearSlot = false;

            if (m_monitorList[i].monitorShare)
            {
               if (m_monitorList[i].monitorShare->alive() == false)
                  clearSlot = true;
               else if (m_monitorList[i].monitorShare->requestShutdown())
                  clearSlot = true;
            }
            else
            {
               clearSlot = true;
            }

            if (clearSlot)
               clearMonitorSlot(i);
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
