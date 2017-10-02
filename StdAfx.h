// stdafx.h : include file for standard system include files,
//      or project specific include files that are used frequently,
//      but are changed infrequently

#ifndef SMM_STDAFX_H
#define SMM_STDAFX_H

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define STRICT
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0400
#endif
#define _ATL_APARTMENT_THREADED

#include <atlbase.h>
#include "bvec.h"
#include "monitorshare.h"

#import "imports/smdispatcher.dll" no_namespace, raw_interfaces_only

unsigned long strtrimlen(const char *s);

struct MonitorRecord
{
   bool inUse;
   MonitorShare *monitorShare;
   char *monitorId;
   char *dirPath;
   unsigned long numExtensions;
   char **extensions;
   char *lotDir;
   char *verifyMap;
   char *mapAction;
   char *verifyLot;
   char *lotAction;
   char *scavenge;
   unsigned long scavengeTime;
   unsigned long scavengeStart;
};

class MonitorList : public BVector <MonitorRecord>
{
   public:
      MonitorList(size_t a = 0, size_t b = 0) :
         BVector <MonitorRecord> (a, b) {}
      ~MonitorList() { setCapacity(0); }

   protected:
      void deallocateProc(size_t, size_t);
      void allocateProc(size_t, size_t);
};

struct MonitorFailure
{
   bool inUse;
   char *monitorId;
   char *dirPath;
   double timeStamp;
   DWORD eCode;
   bool notADir;
};

class MonitorFailureList : public BVector <MonitorFailure>
{
   public:
      MonitorFailureList(size_t a = 0, size_t b = 0) :
         BVector <MonitorFailure> (a, b) {}
      ~MonitorFailureList() { setCapacity(0); }

   protected:
      void deallocateProc(size_t, size_t);
      void allocateProc(size_t, size_t);
};


class SMMDispatcher  
{
   public:
      SMMDispatcher();
      virtual ~SMMDispatcher();

      long setup(bool useAltDispatcher, const char *altDispatcher,
               unsigned long timeOut, bool useServerPool);
      long start();
      void run();
      void shutdown();
      bool alive();

   protected:
      long runDispatcher(bool& moreRecords);
      long getInstance();
      long releaseInstance();
      bool haveInstance();

   protected:
      char m_altDispatcher[_MAX_PATH+1];
      bool m_useAltDispatcher;
      unsigned long m_timeout;

      HANDLE m_globalDispatchEvent;

      HANDLE m_dispatchThreadHandle;
      HANDLE m_dispatchEvent;
      unsigned long m_dispatchThreadId;

      CRITICAL_SECTION m_lock;

      bool m_initialized;
      bool m_didSetup;
      bool m_useServerPool;

      CLSID m_clsIdDispatcher;
      _clsDispatcher *m_iDispatcher;
};

class CServiceModule : public CComModule
{
   public:
      CServiceModule();
      virtual ~CServiceModule();

   public:
      HRESULT RegisterServer(BOOL bRegTypeLib, BOOL bService);
      HRESULT UnregisterServer();
      void Init(_ATL_OBJMAP_ENTRY* p, HINSTANCE h, UINT nServiceNameID,
               const GUID* plibid = NULL);
      void Start();
      void ServiceMain(DWORD dwArgc, LPTSTR* lpszArgv);
      void Handler(DWORD dwOpcode);
      void Run();
      BOOL IsInstalled();
      BOOL Install();
      BOOL Uninstall();
      LONG Unlock();
      void LogEvent(WORD eventType, LPCTSTR pszFormat, ...);
      void SetServiceStatus(DWORD dwState);
      void SetupAsLocalServer();

      int veccat(BCharVec& v, const char *s);
      bool inShutdown();
      void setShutdown(bool b);

      bool okayToScavenge(unsigned long& retryTime);
      void scavengeDone();

   //Implementation
   private:
      static void WINAPI _ServiceMain(DWORD dwArgc, LPTSTR* lpszArgv);
      static void WINAPI _Handler(DWORD dwOpcode);

   protected:
      long loadMonitorData();
      long clearMonitorSlot(unsigned long slot);
      long checkMonitors();
      long findMonitorSlot();

      void removeFromFailureList(const char *monitorId, const char *dirPath);
      void addToFailureList(const char *monitorId, const char *dirPath,
               DWORD eCode, bool notADir);
      long findFailureSlot();
      void reportMonitorFailures();

      long createTimers();
      void stopTimers();

   // data members
   public:
      char m_szServiceName[32];
      char m_szServiceLabel[128];
      SERVICE_STATUS_HANDLE m_hServiceStatus;
      SERVICE_STATUS m_status;
      DWORD dwThreadID;
      BOOL m_bService;

   protected:
      bool m_useServerPool;
      bool m_inShutdown;
      char m_loadMonitorData[_MAX_PATH+1];

      bool m_useAltDispatcher;
      char m_altDispatcher[_MAX_PATH+1];
      unsigned long m_dispatchTimeout;
      SMMDispatcher m_dispatcher;

      MonitorList m_monitorList;
      MonitorFailureList m_failureList;
      double m_failureReportInterval;

      unsigned long m_scavengeSurgeTrigger;

      UINT_PTR m_loadTimer;
      UINT m_loadInterval;

      CRITICAL_SECTION m_scavengeLock;
      unsigned long m_scavengeMax;
      unsigned long m_scavengeCount;
};

extern CServiceModule _Module;
#include <atlcom.h>

//{{AFX_INSERT_LOCATION}}

#endif
