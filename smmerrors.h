#ifndef EG_SMMERRORS_H
#define EG_SMMERRORS_H

////////////////////////////////////////////////////////////////////////////
// SMMErrors.h
//
// This file defines functions and data for handling errors and reporting
// them to the NT/Win2K application event log.
//
////////////////////////////////////////////////////////////////////////////
#include "stdafx.h"

#define SMM_MAX_ERROR_MSG 512

const int Error_Failure = -1;
const int Error_Success = 0;

typedef enum
{
   errSmmInternalError = 0x80004f00,
   errSmmSetupNotCalled,
   errSmmRegistryError,
   errSmmWarning,
   errSmmUnknownException,
   errSmmResource,
   errSmmSetTimerFailed,
   errSmmPostThread,
   errSmmDatabaseServer,
   errSmmMaxErrorShutdown,
   errSmmLicense,
   errSmmConfiguration,
   errSmmBadParameter,
   errSmmIO,
   errSmmEServer,
   errSmmWaitTimeout,
   errSmmRegexpInit,
   errSmmDispatcher,
   errSmmMapAction,
   errSmmScavenge,
   errSmmBadRegExp,
   errSmmLastError
} SORTMapManagerErrors;

// global flag used to indicate the logging level (file based)

const unsigned int logTraceLevelNone = 0;
const unsigned int logTraceLevelErrors = 1;
const unsigned int logTraceLevelDetail = 2;
const unsigned int logTraceLevelFiles = 3;
const unsigned int logTraceLevelProfile = 4;
const unsigned int logTraceLevelFunction = 5;


// class used for function level tracing

class SMMFcnTrace
{
   public:
      SMMFcnTrace(const char *fcnName, unsigned int traceLevel);
      ~SMMFcnTrace();

   protected:
      char m_fcnNameBuffer[256];
      unsigned int m_traceLevel;
      LARGE_INTEGER m_tickCount;
};


#ifdef SMMERROR_MAIN
unsigned int logTraceLevel = logTraceLevelNone;
char logTraceFile[_MAX_PATH+1];
#else
extern unsigned int logTraceLevel;
extern char logTraceFile[_MAX_PATH+1];
#endif

// global function for writing a trace log (based on logInfoLevel) to
// the path as defined in the registry.

void logTrace(const char *fcn, const char *text);

// global function for writing errors to the application event log

void logErrorMsg(HRESULT hr, const char *fcn, const char *text);

// global function for getting a message associated with an internal
// error code (HRESULT)

const char *getSMMErrorMessage(long errorNum);

#endif
