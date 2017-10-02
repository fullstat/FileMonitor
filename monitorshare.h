#ifndef MONITOR_SHARE_H
#define MONITOR_SHARE_H

#include <comdef.h>

const int defaultScavengeTime = 7200;
const int defaultScavengeStart = 30;

const int ReadChangeBufferLen = 8192;

struct DirEntry
{
   char pathName[_MAX_PATH+1];
   char fileName[_MAX_PATH+1];
   double modifiedTime;
   double creationTime;
   DWORD fileAttrib;
   bool notify;
};

class DirList : public BVector <DirEntry>
{
   public:
      DirList(size_t a = 0, size_t b = 0) : 
            BVector <DirEntry> (a,b) {}
      ~DirList() { setCapacity(0); }

   protected:
      void deallocateProc(size_t, size_t);
      void allocateProc(size_t, size_t);
};

typedef BVectorIterator <DirEntry> DirEntryIterator;

struct FileNotifyInfo
{
   char pathName[_MAX_PATH+1];
   double timeStamp;
};

class FileNotifyHistory : public BVector <FileNotifyInfo>
{
   public:
      FileNotifyHistory(size_t a = 0, size_t b = 0) : 
            BVector <FileNotifyInfo> (a,b) {}
      ~FileNotifyHistory() { setCapacity(0); }

   protected:
      void deallocateProc(size_t, size_t);
      void allocateProc(size_t, size_t);
};

class FileNotifyCheck
{
   public:
      FileNotifyCheck();
      virtual ~FileNotifyCheck();

      bool init(unsigned int size = 50, long ageLimit = 30);
      void addMap(const char *pathName, double timeStamp);
      void addLot(const char *pathName);
      bool isADupMap(const char *pathName, double timeStamp);
      bool isADupLot(const char *pathName);
      void dump(const char *id, const char *dirPath);
      void clear() { m_nItems = 0; m_next = 0; }

   protected:
      FileNotifyHistory m_fileHistory;
      FileNotifyHistory m_lotHistory;
      bool m_didInit;
      unsigned long m_next;
      unsigned long m_nItems;
      unsigned long m_lotNext;
      unsigned long m_lotNItems;
      double m_ageLimit;
};

// simple class for handling extensions. could probably be better done
// with the std library, but it will do.

class SmmString
{
   friend class SmmExtNameVec;

   public:
      SmmString();
      SmmString(const char *s);
      SmmString(const SmmString&);
      ~SmmString();

      int set(const char *s);
      int set(const char *s, long ln);

      char *value() { return m_str; }
      bool regExp() { return m_regexp; }
      bool compare(const char *s);
      bool testRegExp();
      void clearRegExpFlag() { m_regexp = false; }

   protected:
      void clear();
      void setRegExpFlag();

   private:
      char *m_str;
      bool m_regexp;
};

class SmmExtNameVec : public BVector <SmmString>
{
   private:
      SmmExtNameVec(const SmmExtNameVec&);   // not implemented
      SmmExtNameVec& operator=(const SmmExtNameVec&);

   public:
      SmmExtNameVec (size_t a = 0, size_t b = 0) : BVector <SmmString> (a, b) {}
      ~SmmExtNameVec () { setCapacity(0); }

      int append(const char *s);
      int append(const char *s, long ln);

   protected:
      void allocateProc (size_t, size_t);
      void deallocateProc (size_t, size_t);
};

typedef BVectorIterator <SmmString> SmmStringIterator;

class SmmExtension
{
   public:
      typedef enum { ActionNone = 0, ActionOK = 1, ActionMissingAssoc = 2,
             ActionAssociation = 3, ActionExclusion = 4 } MonitorAction;

      SmmExtension();
      ~SmmExtension();

      int set(const char *s, long ln=0);
      int addAssoc(const char *s, long ln=0);
      int addExclude(const char *s, long ln=0);
      bool testRegExp();

      MonitorAction find(const char *pathName);

   protected:
      SmmString m_ext;
      SmmExtNameVec m_assocExt;
      SmmExtNameVec m_excludeExt;
};

class SmmExtensionVec : public BVector <SmmExtension *>
{
   private:
      SmmExtensionVec(const SmmExtensionVec&);   // not implemented
      SmmExtensionVec& operator=(const SmmExtensionVec&);

   public:
      SmmExtensionVec (size_t a = 0, size_t b = 0) : BVector <SmmExtension *> (a, b) {}
      ~SmmExtensionVec () { setCapacity(0); }

   protected:
      void allocateProc (size_t, size_t);
      void deallocateProc (size_t, size_t);
};

typedef BVectorIterator <SmmExtension *> SmmExtensionIterator;

// maintain the paths that were valid but missing one or more associations

struct HoldForAssocPath
{
   bool inUse;
   char *pathName;
   double timeStamp;
};

class HoldForAssoc : public BVector <HoldForAssocPath>
{
   public:
      HoldForAssoc(size_t a = 0, size_t b = 0) :
         BVector <HoldForAssocPath> (a, b) {}
      ~HoldForAssoc() { setCapacity(0); }

   protected:
      void deallocateProc(size_t, size_t);
      void allocateProc(size_t, size_t);
};

class QueueEntry
{
   friend class MonitorShare;

   public:
      QueueEntry();
      QueueEntry(const char *pathName, DWORD action, bool fromScavenge);

      char *pathName() { return m_pathName; }
      double timeStamp() { return m_timeStamp; }
      DWORD action() { return m_action; }
      DWORD fromScavenge() { return m_fromScavenge; }

   protected:
      char m_pathName[_MAX_PATH+1];
      double m_timeStamp;
      DWORD m_action;
      bool m_fromScavenge;
      QueueEntry *m_next;
};

class MonitorShare  
{
   public:
      MonitorShare();
      virtual ~MonitorShare();

      long setup(const char *monitorId, const char *dirPath,
               unsigned long numExtensions, const char **extensions,
               const char *lotDir, const char *verifyMap,
               const char *mapAction, const char *verifyLot,
               const char *lotAction, const char *scavenge,
               unsigned long scavengeTime, unsigned long scavengeStart,
               unsigned long scavengeSurgeTrigger, bool useServerPool);
      long start();
      void run();
      void shutdown();
      bool alive();
      bool requestShutdown();

      void runQueue();
      void runScavenge();

      bool checkMap(const char *pathName, double timeStamp);
      bool checkLot(const char *pathName, double timeStamp);

      unsigned char *readBuffer() { return m_readChangeBuffer; }
      OVERLAPPED *readOverlapped() { return &m_readChangeOverlapped; }
      HANDLE dirHandle() { return m_dirHandle; }

   protected:
      bool parseExtensions(const char **extensions,
               unsigned long numExtensions);

      long findInHold(const char *pathName);
      long findBaseInHold(const char *pathName);
      long addToHold(const char *pathName, double timeStamp);
      long removeFromHold(const char *pathName);
      long removeFromHold(long slot);

      long processMap(const char *pathName, double timeStamp);
      long processLot(const char *pathName);

      long addToQueue(const char *pathName, DWORD action,
               bool fromScavenge = false);

      long processScavenge();
      long runScavengeScript(DirList& dList);
      long buildFileDirList(DirList& dList, const char *searchString,
               bool recurseSubDirs = true);

      void setRequestShutdown(bool b);

   private:
      char *m_monitorId;
      char *m_dirPath;
      char *m_verifyMap;
      char *m_mapAction;
      char *m_verifyLot;
      char *m_lotAction;
      char *m_scavenge;

      SmmString m_lotDir;

      unsigned long m_lenVerifyMap;
      unsigned long m_lenMapAction;

      unsigned long m_lenVerifyLot;
      unsigned long m_lenLotAction;

      unsigned long m_lenScavenge;

      SmmExtensionVec m_extInfo;

      HoldForAssoc m_mapHold;
      unsigned long m_numberOnHold;

      FileNotifyCheck m_notifyCheck;

      unsigned long m_scavengeTime;
      unsigned long m_scavengeStart;
      unsigned long m_numExtensions;

      unsigned char *m_readChangeBuffer;
      OVERLAPPED m_readChangeOverlapped;

      HANDLE m_compPort;

      HANDLE m_threadHandle;
      HANDLE m_threadEvent;
      HANDLE m_dirHandle;

      HANDLE m_queueSemaphore;
      HANDLE m_queueThreadHandle;
      HANDLE m_queueEvent;

      HANDLE m_scavengeThreadHandle;
      HANDLE m_scavengeEvent;

      CRITICAL_SECTION m_sLock;
      CRITICAL_SECTION m_qLock;

      unsigned long m_scavengeThreadId;
      unsigned long m_queueThreadId;
      unsigned long m_threadId;

      bool m_initialized;
      bool m_didSetup;
      bool m_useServerPool;
      bool m_inShutdown;

      // request a shutdown when a serious error occurs so that the parent
      // thread can issue a shutdown and restart, if necessary.

      bool m_requestShutdown;

      // map action object CLSID

      CLSID m_clsIdMapAction;

      // primitive queue

      QueueEntry *m_queueHead;
      QueueEntry *m_queueTail;

      // queue depth for non-scavenged entries

      unsigned long m_queueDepth;
      unsigned long m_maxQueueDepth;

      unsigned long m_scavengeSurgeTrigger;
      bool m_scavengeSurgeExceeded;
      bool m_triggerScavenge;
};
#endif
