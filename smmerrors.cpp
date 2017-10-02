#include "stdafx.h"
#include <stdio.h>
#include <comdef.h>

#define SMMERROR_MAIN

#include "smmerrors.h"
#include "EGEventLog.h"
#include "stmutex.h"


static char * errorStrings[errSmmLastError-errSmmInternalError] =
{
   "Internal error.",
   "The setup() method has not been called.",
   "Failed to open or read a SORTMapManager registry entry.",
   "SORTMapManager Warning.",
   "An unknown exception was thrown.",
   "Resource allocation failed (memory, handle, thread).",
   "SetTimer() creation failed.",
   "Failed to post a quit to a worker thread.",
   "Could not establish a connection to the database server.",
   "Exceeded maximum error count, shutting down SORTMapManager.",
   "SORTMapManager license problem.",
   "Installation or configuration problem.",
   "Invalid parameter.",
   "I/O error.",
   "e-Server error.",
   "Timeout waiting for thread shutdown.",
   "Regular expression init() method not called.",
   "SMDispatcher object error.",
   "SMMapActions object error.",
   "Scavenge error.",
   "Invalid regular expression."
};


const char *getSMMErrorMessage(long errorNum)
{
   const char *result = 0;

   if ((errorNum >= errSmmInternalError) && (errorNum < errSmmLastError))
   {
      result = errorStrings[errorNum-errSmmInternalError];
   }

   return result;
}

class LocalPointersLogError
{
   public:
      char *newText;
      char *msgBuffer;

      LocalPointersLogError() : newText(0), msgBuffer(0)  {}
      ~LocalPointersLogError() { delete newText; delete msgBuffer; }
};

void logErrorMsg(HRESULT hr, const char *fcn, const char* text)
{
   LocalPointersLogError LP;
   EGEventLog eventLog("SORTMapManager");

   stLogMutex.get();

   try
   {
      IErrorInfo *errInfo = 0;

// check for an IErrorInfo active object.  if there is one append information
// from it to the text of the message.

      if (GetErrorInfo(0,&errInfo) == S_OK)
      {
         BSTR bStr;

         HRESULT hr2 = errInfo->GetDescription(&bStr);
         if (hr2 == S_OK)
         {
            _bstr_t errDesc(bStr,FALSE);

            hr2 = errInfo->GetSource(&bStr);
            if (hr2 == S_OK)
            {
               _bstr_t errSource(bStr,FALSE);

               int ls = 0;
               if (text)
                  ls = strlen(text) + 5;

               int ln = (ls + errDesc.length() + errSource.length() + 1) * 2;
               LP.newText = new char[ln+1];
               if (LP.newText)
               {
                  if (text)
                     sprintf(LP.newText,"%s '%s (%s)'",text,
                           (const char *) errDesc, (const char *) errSource);
                  else
                     sprintf(LP.newText,"%s (%s)",(const char *) errDesc,
                           (const char *) errSource);
               }
            }
         }
         errInfo->Release();
      }

      const char *smmMsg = 0;
      unsigned int lnSmmMsg = 0;

      if ((hr >= errSmmInternalError) && (hr < errSmmLastError))
      {
         smmMsg = errorStrings[hr-errSmmInternalError];
         if (smmMsg)
            lnSmmMsg = strlen(smmMsg);
      }

      unsigned int lnMsgBuf = 0;
      if (smmMsg)
         lnMsgBuf += (lnSmmMsg + 5);

      if (fcn)
         lnMsgBuf += (strlen(fcn) + 5);

      if (LP.newText)
         lnMsgBuf += (strlen(LP.newText) + 5);
      else if (text)
         lnMsgBuf += (strlen(text) + 5);

      LP.msgBuffer = new char[SMM_MAX_ERROR_MSG+lnMsgBuf];
      if (LP.msgBuffer)
      {
         char localBuf[20];      // used for encoding
         char crLf[2];

         crLf[0] = '\r';
         crLf[1] = '\n';
         crLf[2] = '\0';

         if (hr == errSmmWarning)
            strcpy(LP.msgBuffer,"Warning: ");
         else
            strcpy(LP.msgBuffer,"Error: ");
         if (smmMsg)
         {
            strcat(LP.msgBuffer,"'");
            strcat(LP.msgBuffer,smmMsg);
            strcat(LP.msgBuffer,"' ");
         }
         else
         {
            strcat(LP.msgBuffer,"'Unknown' ");
         }

         if (fcn)
         {
            strcat(LP.msgBuffer," occurred in: ");
            strcat(LP.msgBuffer,fcn);
         }

         if (hr != errSmmWarning)
         {
            strcat(LP.msgBuffer,", Error Code: ");
            sprintf(localBuf,"%x",hr);
            strcat(LP.msgBuffer,localBuf);
         }

         if (text || LP.newText)
            strcat(LP.msgBuffer,crLf);

         if (LP.newText)
            strcat(LP.msgBuffer,LP.newText);
         else if (text)
            strcat(LP.msgBuffer,text);


         if (hr == errSmmWarning)
            eventLog.postMessage(EVENTLOG_WARNING_TYPE, LP.msgBuffer);
         else
            eventLog.postMessage(EVENTLOG_ERROR_TYPE, LP.msgBuffer);
      }
   }
   catch (_com_error&)
   {
      eventLog.postMessage(EVENTLOG_ERROR_TYPE, "COM exception in logErrorMsg()");
   }
   catch (...)
   {
      eventLog.postMessage(EVENTLOG_ERROR_TYPE, "Exception in logErrorMsg()");
   }

   stLogMutex.release();

   if (LP.msgBuffer && logTraceLevel)
   {
      logTrace(fcn,LP.msgBuffer);
   }
}


class LocalPointersLogTrace
{
   public:
      FILE *fp;

      LocalPointersLogTrace() : fp(0) {}
      ~LocalPointersLogTrace() { if (fp) fclose(fp); }
};

void logTrace(const char *fcn, const char *text)
{
   LocalPointersLogTrace LP;
   SYSTEMTIME uTime;

   if (logTraceLevel <= 0)      // no logging
      return;

   stLogMutex.get();

   try
   {
      GetSystemTime(&uTime);
      DWORD threadId = GetCurrentThreadId();

      LP.fp = fopen(logTraceFile,"a");
      if (LP.fp)
      {
         fprintf(LP.fp,"%.4d/%-.2d/%-.2d %.2d:%.2d:%.2d.%.3d",
               uTime.wYear,uTime.wMonth,uTime.wDay,uTime.wHour,uTime.wMinute,
               uTime.wSecond,uTime.wMilliseconds);
         if (fcn)
            fprintf(LP.fp," (%.4x)  %s\n",threadId,fcn);
         else
            fprintf(LP.fp," (%.4x)\n",threadId);

         if (text)
            fprintf(LP.fp,"   %s\n",text);

         fclose(LP.fp);
         LP.fp = 0;      // for dtor
      }
   }
   catch (...)      // fire-wall so the mutex is guaranteed to be released
   {
   }

   stLogMutex.release();
}

SMMFcnTrace::SMMFcnTrace(const char *fcnName, unsigned int traceLevel)
{
   m_fcnNameBuffer[0] = '\0';
   m_traceLevel = traceLevel;
   m_tickCount.QuadPart = 0;

   if (traceLevel >= logTraceLevelProfile)
   {
      QueryPerformanceCounter(&m_tickCount);
      strcpy(m_fcnNameBuffer,fcnName);

      if (traceLevel >= logTraceLevelFunction)
      {
         char localBuf[256];

         strcpy(localBuf,"---> ");
         strcat(localBuf,m_fcnNameBuffer);
         logTrace(localBuf,0);
      }
   }
}

SMMFcnTrace::~SMMFcnTrace()
{
   if (m_traceLevel >= logTraceLevelProfile)
   {
      char localBuf[256];
      LARGE_INTEGER tickEnd;
      LARGE_INTEGER freq;
      double nSecs = 0.0;

      QueryPerformanceCounter(&tickEnd);
      m_tickCount.QuadPart = tickEnd.QuadPart - m_tickCount.QuadPart;

      if (QueryPerformanceFrequency(&freq))
         nSecs = ((double) m_tickCount.QuadPart)/((double) freq.QuadPart);

      sprintf(localBuf,"%s duration = %f seconds.", m_fcnNameBuffer, nSecs);
      logTrace(localBuf,0);

      if (m_traceLevel >= logTraceLevelFunction)
      {
         strcpy(localBuf,"<--- ");
         strcat(localBuf,m_fcnNameBuffer);
         logTrace(localBuf,0);
      }
   }
}
