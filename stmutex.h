#ifndef STMUTEX
#define STMUTEX

// simple mutex object for synchronizing log tracing

class STMutex
{
   protected:
      HANDLE hMutex0;
      BOOL didInit0;

   public:
      STMutex(BOOL doInit);
      STMutex();
      virtual ~STMutex();

      BOOL didInit() { return didInit0; }
      BOOL init();
      BOOL get();
      BOOL release();
};

#ifdef STMUTEX_MAIN
STMutex stLogMutex(TRUE);
#else
extern STMutex stLogMutex;
#endif


#endif
