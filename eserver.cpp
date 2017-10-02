

#include "stdafx.h"
#include <stdio.h>
#include <process.h>
#include "eserver.h"
#include "smmerrors.h"
#include <eh.h>

const unsigned long maxRetryCount = 3;
const unsigned long retryDelay = 50;   // milliseconds


eServer::eServer(bool useServerPool)
{
   static const char *funcName = "eServer::eServer()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   m_useServerPool = useServerPool;

   m_initialized = false;
   m_errBuf = 0;
   m_iStatit = 0;

   strcpy(m_serverProgId,"StatitLib.Statit");
}

eServer::~eServer()
{
   static const char *funcName = "eServer::~eServer()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   if (m_iStatit)
   {
      m_iStatit->Release();
      m_iStatit = 0;
   }

   delete m_errBuf;
}

long eServer::init()
{
   static const char *funcName = "eServer::init()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   CLSID clsIdStatit;
   HRESULT hr = S_OK;
   bool gotStatit = false;

   try
   {
      m_initialized = false;

// this may get run again if a pooled e-Server is requested and it
// fails.  it will get retried without a pooled e-Server.  the code
// is rather convoluted since it tests for the loop counter in several
// places.  this is so that the appropriate error message is generated
// when everything fails.  the logic is so any failure causes a complete
// retry, not just the item that failed.

      unsigned long tryCount = 0;

      do
      {
         if (m_iStatit)
         {
            m_iStatit->Release();
            m_iStatit = 0;
         }

         if (tryCount > 0)            // very brief delay
            ::Sleep(retryDelay);

         if (m_useServerPool)
         {
            strcpy(m_serverProgId,"StatitPool.Statit");
            hr = CLSIDFromProgID(OLESTR("StatitPool.Statit"), &clsIdStatit);
            if (FAILED(hr))
            {
               m_useServerPool = false;
               strcpy(m_serverProgId,"StatitLib.Statit");
               hr = CLSIDFromProgID(OLESTR("StatitLib.Statit"), &clsIdStatit);
            }
         }
         else
         {
            strcpy(m_serverProgId,"StatitLib.Statit");
            hr = CLSIDFromProgID(OLESTR("StatitLib.Statit"), &clsIdStatit);
         }

         if (FAILED(hr))
         {
            if (tryCount >= maxRetryCount)
            {
               char errBuf[eServerSysErrBufLen];
               char localBuf[eServerSysErrBufLen+256];

               if (! FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, hr, 0,
                     errBuf, eServerSysErrBufLen, 0))
               {
                  sprintf(localBuf,
                     "Failed to lookup the %s CLSID. Error code=%x.",
                     m_serverProgId, hr);
               }
               else
               {
                  sprintf(localBuf,
                     "Failed to lookup the %s CLSID (%s). Error code=%x.",
                     m_serverProgId, errBuf, hr);
               }
               logErrorMsg(errSmmEServer, funcName, localBuf);
            }
         }
         else
         {
            if (m_useServerPool)
               hr = CoCreateInstance(clsIdStatit, 0, CLSCTX_SERVER,
                        __uuidof(IPooledStatit), (void **) &m_iStatit);
            else
               hr = CoCreateInstance(clsIdStatit, 0, CLSCTX_SERVER,
                  __uuidof(IStatit), (void **) &m_iStatit);

            if (FAILED(hr))
            {
               if (tryCount >= maxRetryCount)
               {
                  char errBuf[eServerSysErrBufLen];
                  char localBuf[eServerSysErrBufLen+256];

                  if (! FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, hr, 0,
                        errBuf, eServerSysErrBufLen, 0))
                  {
                     sprintf(localBuf,
                        "Failed to create an instance of %s. Error code=%x.",
                        m_serverProgId, hr);
                  }
                  else
                  {
                     sprintf(localBuf,
                        "Failed to create an instance of %s (%s). Error code=%x.",
                        m_serverProgId, errBuf, hr);
                  }
                  logErrorMsg(errSmmEServer, funcName, localBuf);
               }
               else
               {
                  m_useServerPool = false;
               }
            }
            else
            {
               hr = m_iStatit->Init();
               if (FAILED(hr))
               {
                  if (tryCount >= maxRetryCount)
                  {
                     if (getErrors(hr))
                        logErrorMsg(errSmmEServer, funcName, m_errBuf);
                     else
                        logErrorMsg(errSmmEServer, funcName,
                              "e-Server Init() failed.");
                  }
                  else
                  {
                     m_useServerPool = false;
                  }
               }
            }
         }
         tryCount++;
      }
      while (FAILED(hr) && tryCount <= maxRetryCount);
      if (FAILED(hr))
         return Error_Failure;

      m_initialized = true;
   }
   catch (_com_error &e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
      return Error_Failure;
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName,
            "Exception occurred during e-Server initialization");
      return Error_Failure;
   }

   return Error_Success;
}


class GetCOMErrorLocalPointers
{
   public:
      char *errMsg;
      unsigned long lnErrBuf;
      unsigned long errBufSize;
      IError *pErr;
      IErrors *pErrors;
      IErrorInfo *errInfo;

      GetCOMErrorLocalPointers() : errMsg(0), lnErrBuf(0), pErr(0), pErrors(0),
                                   errInfo(0), errBufSize(eServerErrBufLen) {}
      ~GetCOMErrorLocalPointers() { delete errMsg;
                                    if (pErr) pErr->Release();
                                    if (pErrors) pErrors->Release();
                                    if (errInfo) errInfo->Release(); }

      void errMsgCat(const char *strValue);
};


void GetCOMErrorLocalPointers::errMsgCat(const char *strValue)
{
   if (errMsg == 0)
   {
      errMsg = new char[errBufSize+1];
      if (errMsg == 0)
         return;

      errMsg[0] = '\0';    // needed for strcat()
   }

   unsigned long ln = strlen(strValue);
   while (lnErrBuf + ln + 1 >= errBufSize)
   {
      errBufSize += eServerErrBufLen;
      char *newBuffer = new char[errBufSize+1];
      if (newBuffer == 0)
         return;

      strcpy(newBuffer, errMsg);

      delete errMsg;         // take ownership
      errMsg = newBuffer;
   }

   strcat(errMsg, strValue);
   lnErrBuf += ln;
}


// the pointer returned references memory owned by the class

char *eServer::getErrors(HRESULT hr)
{
   static const char *funcName = "eServer::getErrors()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   char localBuf[10];   // for encoding error numbers

   try
   {
      if (m_errBuf)
      {
         delete m_errBuf;
         m_errBuf = 0;
      }

      long eCode = HRESULT_CODE(hr);
      if (m_iStatit && (eCode > swErrNone && eCode <= swErrFatal))
      {
         GetCOMErrorLocalPointers LP;

         HRESULT hr2 = m_iStatit->get_Errors(&LP.pErrors);
         if (SUCCEEDED(hr2))
         {
            long eCount = 0;

            hr2 = LP.pErrors->get_Count(&eCount);
            if (SUCCEEDED(hr2) && eCount > 0)
            {
               char crLf[2];

               crLf[0] = '\r';
               crLf[1] = '\n';
               crLf[2] = '\0';

               for (long i = 0; i < eCount; i++)
               {

                  LP.errMsgCat(crLf);

                  if (i > 0)   // one more for subsequent error records
                     LP.errMsgCat(crLf);

                  _variant_t tv = (long) i;

                  hr2 = LP.pErrors->get_Item(tv,&LP.pErr);
                  if (SUCCEEDED(hr2))
                  {
                     long errNum;
                     hr2 = LP.pErr->get_Number(&errNum);
                     if (SUCCEEDED(hr2) && errNum != swErrWarning)
                     {
                        BSTR bStr;
                        long errNative = 0;

                        hr2 = LP.pErr->get_Description(&bStr);
                        if (SUCCEEDED(hr2))
                           hr2 = LP.pErr->get_Native(&errNative);

                        if (SUCCEEDED(hr2))
                        {
                           _bstr_t errDesc(bStr,false);   // take ownership

                           sprintf(localBuf,"%ld",errNum);
                           LP.errMsgCat("Error Number: ");
                           LP.errMsgCat(localBuf);

                           if (errNative != 0)
                           {
                              sprintf(localBuf,"%ld",errNative);
                              LP.errMsgCat(" (Native error: ");
                              LP.errMsgCat(localBuf);
                              LP.errMsgCat(")");
                           }

                           if (errDesc.length() > 0)
                           {
                              LP.errMsgCat(crLf);
                              LP.errMsgCat("Error Description:");
                              LP.errMsgCat(crLf);
                              LP.errMsgCat((const char *) errDesc);
                           }
                        }

                        hr2 = LP.pErr->get_Cmd(&bStr);
                        if (SUCCEEDED(hr2))
                        {
                           _bstr_t errCmd(bStr,false);   // take ownership

                           if (errCmd.length() > 0)
                           {
                              LP.errMsgCat(crLf);
                              LP.errMsgCat("Command:");
                              LP.errMsgCat(crLf);
                              LP.errMsgCat((const char *) errCmd);
                           }
                        }
                     }
                     LP.pErr->Release();
                     LP.pErr = 0;
                  }
               }

               m_errBuf = LP.errMsg;
               LP.errMsg = 0;
            }
            LP.pErrors->Release();
            LP.pErrors = 0;
         }
      }
      else
      {
         char errBuf[eServerSysErrBufLen+1];
         char localBuf[eServerSysErrBufLen+256];

         if (! FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, hr, 0,
                  errBuf, eServerSysErrBufLen, 0))
         {
            sprintf(localBuf,"%s, Error code=%x.", m_serverProgId, hr);
         }
         else
         {
            sprintf(localBuf,"%s (%s). Error code=%x.", m_serverProgId,
                  errBuf, hr);
         }

         m_errBuf = new char[strlen(localBuf)+1];
         if (m_errBuf)
         {
            strcpy(m_errBuf, localBuf);
         }
      }
   }
   catch (...)
   {
      if (m_errBuf)
      {
         delete m_errBuf;
         m_errBuf = 0;
      }
   }

   return m_errBuf;
}


// run a command using e-Server

long eServer::runCmd(const char *cmd)
{
   static const char *funcName = "eServer::runCmd()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmEServer, funcName, "init() has not been called.");
      return retval;
   }

   try
   {
      _bstr_t bCmd(cmd);

      HRESULT hr = m_iStatit->ExecCmd(bCmd);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName, "e-Server ExecCmd() failed.");
      }
      else
      {
         retval = Error_Success;
      }
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}

// helper function to set a global int in e-Server

long eServer::putGlobalInt(const char *var, long ival)
{
   static const char *funcName = "eServer::putGlobalInt()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmEServer, funcName, "init() has not been called.");
      return retval;
   }

   _variant_t value((long) ival,VT_I4);
   VARIANT vn = value.Detach();

   try
   {
      _bstr_t varName(var);

      HRESULT hr = m_iStatit->put_GlobalVar(varName,vn);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server put_GlobalVar() failed.");
      }
      else
      {
         retval = Error_Success;
      }
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   VariantClear(&vn);
   return retval;
}

// helper function to set a global string in e-Server

long eServer::putGlobalString(const char *var, const char *stringVal)
{
   static const char *funcName = "eServer::putGlobalString()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmEServer, funcName, "init() has not been called.");
      return retval;
   }

   _variant_t value(stringVal);
   VARIANT vn = value.Detach();

   try
   {
      _bstr_t varName(var);

      HRESULT hr = m_iStatit->put_GlobalVar(varName,vn);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server put_GlobalVar() failed.");
      }
      else
      {
         retval = Error_Success;
      }
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   VariantClear(&vn);
   return retval;
}

// helper function to set a global double in e-Server

long eServer::putGlobalDouble(const char *var, double dval)
{
   static const char *funcName = "eServer::putGlobalDouble()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmEServer, funcName, "init() has not been called.");
      return retval;
   }

   _variant_t value(dval);
   VARIANT vn = value.Detach();

   try
   {
      _bstr_t varName(var);

      HRESULT hr = m_iStatit->put_GlobalVar(varName,vn);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server put_GlobalVar() failed.");
      }
      else
      {
         retval = Error_Success;
      }
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   VariantClear(&vn);
   return retval;
}

// helper function to get a global int from e-Server

long eServer::getGlobalInt(const char *var, long& ival)
{
   static const char *funcName = "eServer::getGlobalInt()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmEServer, funcName, "init() has not been called.");
      return retval;
   }

   VARIANT vn;
   ival = 0;

   try
   {
      _bstr_t varName(var);

      HRESULT hr = m_iStatit->get_GlobalVar(varName,&vn);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server get_GlobalVar() failed.");
      }
      else
      {
         retval = Error_Success;
      }
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   if (retval == Error_Failure)
      return retval;

   retval = Error_Failure;

   try
   {
      _variant_t vt;

      vt.Attach(vn);
      ival = (long) vt;
      retval = Error_Success;
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}

// helper function to get a global string from e-Server
// the caller is responsible for freeing the memory for the returned string

long eServer::getGlobalString(const char *var, char *& strValue)
{
   static const char *funcName = "eServer::getGlobalString()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmEServer, funcName, "init() has not been called.");
      return retval;
   }

   VARIANT vn;
   strValue = 0;

   try
   {
      _bstr_t varName(var);

      HRESULT hr = m_iStatit->get_GlobalVar(varName,&vn);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server get_GlobalVar() failed.");
      }
      else
      {
         retval = Error_Success;
      }
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   if (retval == Error_Failure)
      return retval;

   retval = Error_Failure;

   try
   {
      _variant_t vt;

      vt.Attach(vn);
      _bstr_t bt = (_bstr_t) vt;
      strValue = new char[bt.length() + 1];
      if (strValue)
      {
         strcpy(strValue, (const char *) bt);
         retval = Error_Success;
      }
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}

// helper function to get a global double from e-Server

long eServer::getGlobalDouble(const char *var, double& dval)
{
   static const char *funcName = "eServer::getGlobalDouble()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmEServer, funcName, "init() has not been called.");
      return retval;
   }

   VARIANT vn;
   dval = 0.0;

   try
   {
      _bstr_t varName(var);

      HRESULT hr = m_iStatit->get_GlobalVar(varName,&vn);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server get_GlobalVar() failed.");
      }
      else
      {
         retval = Error_Success;
      }
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   if (retval == Error_Failure)
      return retval;

   retval = Error_Failure;

   try
   {
      _variant_t vt;

      vt.Attach(vn);
      dval = (double) vt;
      retval = Error_Success;
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}

// the caller is responsible for freeing the returned pointer

class GetVarLocalPointers
{
   public:
      IData *pData;
      char **strArray;
      char *strValue;
      double *dblArray;
      unsigned long nCases;

      GetVarLocalPointers() : pData(0), strArray(0), dblArray(0), nCases(0),
                              strValue(0) {}
      ~GetVarLocalPointers() { if (pData) pData->Release(); 
                               delete strValue;
                               delete dblArray;
                               if (strArray) {
                                  for (unsigned long i = 0; i < nCases; i++)
                                     delete strArray[i];
                                  delete strArray;
                               } }
};


long eServer::getDoubleVar(const char *var, unsigned long& nCases,
            double *& dblArray)
{
   static const char *funcName = "eServer::getDoubleVar()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;
   nCases = 0;
   dblArray = 0;

   if (! m_initialized)
   {
      logErrorMsg(errSmmEServer, funcName, "init() has not been called.");
      return retval;
   }

// first, determine the content of the variable.  if the variable doesn't
// exist this will be caught here as well

   VARIANT vn;
   long lContent = 0;

   try
   {
      char localBuf[256];

      if (strchr(var,'.'))
         sprintf(localBuf,"content(%s)", var);
      else
         sprintf(localBuf,"content(.%s)", var);

      _bstr_t varContent(localBuf);

      HRESULT hr = m_iStatit->Eval(varContent,&vn);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server Eval() failed.");
      }
      else
      {
         _variant_t vt;

         vt.Attach(vn);
         lContent = (long) vt;

         if (lContent == 1)         // string not allowed
         {
            sprintf(localBuf,
               "The variable '%s' contains string data.", var);
            logErrorMsg(errSmmEServer, funcName, localBuf);
         }
         else
         {
            retval = Error_Success;
         }
      }
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   if (retval == Error_Failure)
      return retval;

   retval = Error_Failure;

// next, get the number of cases

   try
   {
      char localBuf[256];

      if (strchr(var,'.'))
         sprintf(localBuf,"case(%s)", var);
      else
         sprintf(localBuf,"case(.%s)", var);

      _bstr_t varCases(localBuf);

      HRESULT hr = m_iStatit->Eval(varCases,&vn);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server Eval() failed.");
      }
      else
      {
         _variant_t vt;

         vt.Attach(vn);
         nCases = (long) vt;
         retval = Error_Success;
      }
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

// if there was an error or zero cases, we are done

   if (retval == Error_Failure || nCases == 0)
      return retval;

   retval = Error_Failure;

// now, allocate the buffer and get the data.  the get_Value() method
// is used instead of GetArray().  it is less efficient but avoids the
// pain of safe arrays

   try
   {
      GetVarLocalPointers LP;
      char localBuf[256];

      HRESULT hr = m_iStatit->get_Data(&LP.pData);
      if (FAILED(hr))
      {
         sprintf(localBuf,"e-Server get_Data() returned %x.", hr);
         logErrorMsg(errSmmEServer, funcName, localBuf);
         return retval;
      }

      LP.dblArray = new double[nCases];
      if (LP.dblArray == 0)
      {
         sprintf(localBuf,"double(%ld)", nCases);
         logErrorMsg(errSmmResource, funcName, localBuf);
         return retval;
      }

      _bstr_t varName(var);

      for (unsigned long i = 0; i < nCases; i++)
      {
         hr = LP.pData->get_Value(varName, i+1, &vn);
         if (FAILED(hr))
         {
            if (getErrors(hr))
               logErrorMsg(errSmmEServer, funcName, m_errBuf);
            else
               logErrorMsg(errSmmEServer, funcName,
                        "e-Server get_Value() failed.");
            return retval;
         }

         _variant_t vt;

         vt.Attach(vn);
         LP.dblArray[i] = (double) vt;
      }

      dblArray = LP.dblArray;
      LP.dblArray = 0;      // for dtor

      retval = Error_Success;
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}

long eServer::getStringVar(const char *var, unsigned long& nCases,
            char **& strArray)
{
   static const char *funcName = "eServer::getStringVar()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;
   nCases = 0;
   strArray = 0;

   if (! m_initialized)
   {
      logErrorMsg(errSmmEServer, funcName, "init() has not been called.");
      return retval;
   }

// first, determine the content of the variable.  if the variable doesn't
// exist this will be caught here as well

   VARIANT vn;
   long lContent = 0;

   try
   {
      char localBuf[256];

      if (strchr(var,'.'))
         sprintf(localBuf,"content(%s)", var);
      else
         sprintf(localBuf,"content(.%s)", var);

      _bstr_t varContent(localBuf);

      HRESULT hr = m_iStatit->Eval(varContent,&vn);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server Eval() failed.");
      }
      else
      {
         _variant_t vt;

         vt.Attach(vn);
         lContent = (long) vt;

         if (lContent != 1)         // numeric not allowed
         {
            sprintf(localBuf,
               "The variable '%s' contains numeric data.", var);
            logErrorMsg(errSmmEServer, funcName, localBuf);
         }
         else
         {
            retval = Error_Success;
         }
      }
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   if (retval == Error_Failure)
      return retval;

   retval = Error_Failure;

// next, get the number of cases

   try
   {
      char localBuf[256];

      if (strchr(var,'.'))
         sprintf(localBuf,"case(%s)", var);
      else
         sprintf(localBuf,"case(.%s)", var);

      _bstr_t varCases(localBuf);

      HRESULT hr = m_iStatit->Eval(varCases,&vn);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server Eval() failed.");
      }
      else
      {
         _variant_t vt;

         vt.Attach(vn);
         nCases = (long) vt;
         retval = Error_Success;
      }
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

// if there was an error or zero cases, we are done

   if (retval == Error_Failure || nCases == 0)
      return retval;

   retval = Error_Failure;

// now, allocate the buffer and get the data.  the get_Value() method
// is used instead of GetArray().  it is less efficient but avoids the
// pain of safe arrays

   try
   {
      GetVarLocalPointers LP;
      char localBuf[256];
      unsigned long i;

      HRESULT hr = m_iStatit->get_Data(&LP.pData);
      if (FAILED(hr))
      {
         sprintf(localBuf,"e-Server get_Data() returned %x.", hr);
         logErrorMsg(errSmmEServer, funcName, localBuf);
         return retval;
      }

      LP.strArray = new char *[nCases];
      if (LP.strArray == 0)
      {
         sprintf(localBuf,"char *(%ld)", nCases);
         logErrorMsg(errSmmResource, funcName, localBuf);
         return retval;
      }

      for (i = 0; i < nCases; i++)
         LP.strArray[i] = 0;

      LP.nCases = nCases;   // for dtor

      _bstr_t varName(var);

      for (i = 0; i < nCases; i++)
      {
         hr = LP.pData->get_Value(varName, i+1, &vn);
         if (FAILED(hr))
         {
            if (getErrors(hr))
               logErrorMsg(errSmmEServer, funcName, m_errBuf);
            else
               logErrorMsg(errSmmEServer, funcName,
                        "e-Server get_Value() failed.");
            return retval;
         }

         _variant_t vt;

         vt.Attach(vn);
         _bstr_t bt = (_bstr_t) vt;
         LP.strValue = new char[bt.length() + 1];
         if (LP.strValue == 0)
         {
            sprintf(localBuf,"char(%ld)", bt.length()+1);
            logErrorMsg(errSmmResource, funcName, localBuf);
            return retval;
         }

         strcpy(LP.strValue, (const char *) bt);
         LP.strArray[i] = LP.strValue;
         LP.strValue = 0;      // for dtor
      }

      strArray = LP.strArray;
      LP.strArray = 0;      // for dtor
      LP.nCases = 0;

      retval = Error_Success;
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}


long eServer::doesVarExist(const char *var, bool& itExists)
{
   static const char *funcName = "eServer::doesVarExist()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;
   itExists = false;

   if (! m_initialized)
   {
      logErrorMsg(errSmmEServer, funcName, "init() has not been called.");
      return retval;
   }

   VARIANT vn;

   try
   {
      char localBuf[256];

      sprintf(localBuf,"ifexist(\"%s\")", var);

      _bstr_t varExists(localBuf);

      HRESULT hr = m_iStatit->Eval(varExists,&vn);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server Eval() failed.");
      }
      else
      {
         _variant_t vt;

         vt.Attach(vn);
         long vNum = (long) vt;

         if (vNum != 0)         // it exists
            itExists = true;

         retval = Error_Success;
      }
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}


long eServer::getValue(const char *var, unsigned long nRow, double& dblValue)
{
   static const char *funcName = "eServer::getValue()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;
   dblValue = 0.0;

   if (! m_initialized)
   {
      logErrorMsg(errSmmEServer, funcName, "init() has not been called.");
      return retval;
   }

// first, determine the content of the variable.  if the variable doesn't
// exist this will be caught here as well

   VARIANT vn;
   long lContent = 0;

   try
   {
      char localBuf[256];

      if (strchr(var,'.'))
         sprintf(localBuf,"content(%s)", var);
      else
         sprintf(localBuf,"content(.%s)", var);

      _bstr_t varContent(localBuf);

      HRESULT hr = m_iStatit->Eval(varContent,&vn);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server Eval() failed.");
      }
      else
      {
         _variant_t vt;

         vt.Attach(vn);
         lContent = (long) vt;

         if (lContent == 1)         // string not allowed
         {
            sprintf(localBuf,
               "The variable '%s' contains string data.", var);
            logErrorMsg(errSmmEServer, funcName, localBuf);
         }
         else
         {
            retval = Error_Success;
         }
      }
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   if (retval == Error_Failure)
      return retval;

   retval = Error_Failure;

// now, get the value

   try
   {
      GetVarLocalPointers LP;
      char localBuf[256];

      HRESULT hr = m_iStatit->get_Data(&LP.pData);
      if (FAILED(hr))
      {
         sprintf(localBuf,"e-Server get_Data() returned %x.", hr);
         logErrorMsg(errSmmEServer, funcName, localBuf);
         return retval;
      }

      _bstr_t varName(var);

      hr = LP.pData->get_Value(varName, nRow, &vn);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server get_Value() failed.");
         return retval;
      }

      _variant_t vt;

      vt.Attach(vn);
      dblValue = (double) vt;
      retval = Error_Success;
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}

long eServer::getString(const char *var, unsigned long nRow, char *& str)
{
   static const char *funcName = "eServer::getString()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;
   str = 0;

   if (! m_initialized)
   {
      logErrorMsg(errSmmEServer, funcName, "init() has not been called.");
      return retval;
   }

// first, determine the content of the variable.  if the variable doesn't
// exist this will be caught here as well

   VARIANT vn;
   long lContent = 0;

   try
   {
      char localBuf[256];

      if (strchr(var,'.'))
         sprintf(localBuf,"content(%s)", var);
      else
         sprintf(localBuf,"content(.%s)", var);

      _bstr_t varContent(localBuf);

      HRESULT hr = m_iStatit->Eval(varContent,&vn);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server Eval() failed.");
      }
      else
      {
         _variant_t vt;

         vt.Attach(vn);
         lContent = (long) vt;

         if (lContent != 1)         // numeric not allowed
         {
            sprintf(localBuf,
               "The variable '%s' contains numeric data.", var);
            logErrorMsg(errSmmEServer, funcName, localBuf);
         }
         else
         {
            retval = Error_Success;
         }
      }
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   if (retval == Error_Failure)
      return retval;

   retval = Error_Failure;

// now, get the data

   try
   {
      GetVarLocalPointers LP;
      char localBuf[256];

      HRESULT hr = m_iStatit->get_Data(&LP.pData);
      if (FAILED(hr))
      {
         sprintf(localBuf,"e-Server get_Data() returned %x.", hr);
         logErrorMsg(errSmmEServer, funcName, localBuf);
         return retval;
      }

      _bstr_t varName(var);

      hr = LP.pData->get_Value(varName, nRow, &vn);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server get_Value() failed.");
         return retval;
      }

      _variant_t vt;

      vt.Attach(vn);
      _bstr_t bt = (_bstr_t) vt;
      LP.strValue = new char[bt.length() + 1];
      if (LP.strValue == 0)
      {
         sprintf(localBuf,"char(%ld)", bt.length()+1);
         logErrorMsg(errSmmResource, funcName, localBuf);
         return retval;
      }

      strcpy(LP.strValue, (const char *) bt);
      str = LP.strValue;
      LP.strValue = 0;      // for dtor
      retval = Error_Success;
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}


long eServer::putValue(const char *var, long nRow, double dblValue)
{
   static const char *funcName = "eServer::putValue()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmEServer, funcName, "init() has not been called.");
      return retval;
   }

// first, determine the content of the variable.  if the variable doesn't
// exist this will be caught here as well

   VARIANT vn;
   long lContent = 0;

   try
   {
      char localBuf[256];

      if (strchr(var,'.'))
         sprintf(localBuf,"content(%s)", var);
      else
         sprintf(localBuf,"content(.%s)", var);

      _bstr_t varContent(localBuf);

      HRESULT hr = m_iStatit->Eval(varContent,&vn);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server Eval() failed.");
      }
      else
      {
         _variant_t vt;

         vt.Attach(vn);
         lContent = (long) vt;

         if (lContent == 1)         // string not allowed
         {
            sprintf(localBuf,
               "The variable '%s' contains string data.", var);
            logErrorMsg(errSmmEServer, funcName, localBuf);
         }
         else
         {
            retval = Error_Success;
         }
      }
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   if (retval == Error_Failure)
      return retval;

   retval = Error_Failure;

// now, put the value

   try
   {
      GetVarLocalPointers LP;
      char localBuf[256];

      HRESULT hr = m_iStatit->get_Data(&LP.pData);
      if (FAILED(hr))
      {
         sprintf(localBuf,"e-Server get_Data() returned %x.", hr);
         logErrorMsg(errSmmEServer, funcName, localBuf);
         return retval;
      }

      _bstr_t varName(var);
      _variant_t vValue(dblValue);

      hr = LP.pData->put_Value(varName, nRow, vValue);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server put_Value() failed.");
         return retval;
      }

      retval = Error_Success;
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}


long eServer::putString(const char *var, long nRow, const char *str)
{
   static const char *funcName = "eServer::putString()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmEServer, funcName, "init() has not been called.");
      return retval;
   }

// first, determine the content of the variable.  if the variable doesn't
// exist this will be caught here as well

   VARIANT vn;
   long lContent = 0;

   try
   {
      char localBuf[256];

      if (strchr(var,'.'))
         sprintf(localBuf,"content(%s)", var);
      else
         sprintf(localBuf,"content(.%s)", var);

      _bstr_t varContent(localBuf);

      HRESULT hr = m_iStatit->Eval(varContent,&vn);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server Eval() failed.");
      }
      else
      {
         _variant_t vt;

         vt.Attach(vn);
         lContent = (long) vt;

         if (lContent != 1)         // numeric not allowed
         {
            sprintf(localBuf,
               "The variable '%s' contains numeric data.", var);
            logErrorMsg(errSmmEServer, funcName, localBuf);
         }
         else
         {
            retval = Error_Success;
         }
      }
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   if (retval == Error_Failure)
      return retval;

   retval = Error_Failure;

// now, get the data

   try
   {
      GetVarLocalPointers LP;
      char localBuf[256];

      HRESULT hr = m_iStatit->get_Data(&LP.pData);
      if (FAILED(hr))
      {
         sprintf(localBuf,"e-Server get_Data() returned %x.", hr);
         logErrorMsg(errSmmEServer, funcName, localBuf);
         return retval;
      }

      _bstr_t varName(var);
      _variant_t vValue(str);

      hr = LP.pData->put_Value(varName, nRow, vValue);
      if (FAILED(hr))
      {
         if (getErrors(hr))
            logErrorMsg(errSmmEServer, funcName, m_errBuf);
         else
            logErrorMsg(errSmmEServer, funcName,
                     "e-Server put_Value() failed.");
         return retval;
      }

      retval = Error_Success;
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}


long eServer::createDoubleVar(const char *var)
{
   static const char *funcName = "eServer::createDoubleVar()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmEServer, funcName, "init() has not been called.");
      return retval;
   }

// create the variable (overwriting is okay) there will be one case of 0

   try
   {
      char localBuf[256];

      sprintf(localBuf,"assign %s 0 /nolist", var);
      retval = runCmd(localBuf);
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}

long eServer::createStringVar(const char *var, unsigned long nChars)
{
   static const char *funcName = "eServer::createStringVar()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmEServer, funcName, "init() has not been called.");
      return retval;
   }

// create the variable (overwriting is okay) there will be one case of the
// requested length

   try
   {
      char localBuf[256];

      if (nChars > 0)
         sprintf(localBuf,"let %s = repeat(\" \",%d)", var, nChars-1);
      else
         sprintf(localBuf,"let %s = \"\"", var);

      retval = runCmd(localBuf);
   }
   catch (_com_error& e)
   {
      logErrorMsg(e.Error(), funcName, e.ErrorMessage());
   }
   catch (...)
   {
      logErrorMsg(errSmmUnknownException, funcName, 0);
   }

   return retval;
}
