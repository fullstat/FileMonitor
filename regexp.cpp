#include "stdafx.h"
#include <stdio.h>
#include <comdef.h>
#include "smmerrors.h"
#include "regexp.h"

/*
 * this class is a simple wrapper around the VbScript.RegExp object.
 * the underlying support of the RegExp object is limited since all
 * I want to determine is if there is a pattern match.  other C++
 * implmentations of regular expressions that I checked were not
 * thread safe.
 */

const long badRegularExpression = 0x800a139a;

SMRegExp::SMRegExp()
{
   static const char *funcName = "SMRegExp::SMRegExp()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   m_iRegExp = 0;
   m_input = 0;
   m_value = 0;

   m_index = 0;
   m_lastIndex = 0;

   m_found = false;
   m_global = false;
   m_ignoreCase = true;
}

SMRegExp::~SMRegExp()
{
   static const char *funcName = "SMRegExp::~SMRegExp()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   if (m_iRegExp)
   {
      m_iRegExp->Release();
      m_iRegExp = 0;
   }

   delete m_input;
   m_input = 0;

   delete m_value;
   m_value = 0;
}


long SMRegExp::init()
{
   static const char *funcName = "SMRegExp::setGlobal()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   m_initialized = false;

   if (m_iRegExp)
   {
      m_iRegExp->Release();
      m_iRegExp = 0;
   }

   delete m_input;
   m_input = 0;

   delete m_value;
   m_value = 0;

   m_index = 0;
   m_lastIndex = 0;

   m_found = false;
   m_global = false;
   m_ignoreCase = true;

   try
   {
      CLSID clsIdRegExp;

      HRESULT hr = CLSIDFromProgID(_bstr_t("VBScript.RegExp"), &clsIdRegExp);
      if (FAILED(hr))
      {
         logErrorMsg(hr, funcName, "CLSIDFromProgID()");
      }
      else
      {
         hr = CoCreateInstance(clsIdRegExp, 0, CLSCTX_SERVER,
               __uuidof(IRegExp2), reinterpret_cast<void **>(&m_iRegExp));
         if (FAILED(hr))
         {
            logErrorMsg(hr, funcName, "CoCreateInstance()");
         }
         else
         {
            VARIANT_BOOL vb = VARIANT_TRUE;
            retval = Error_Success;

            HRESULT hr = m_iRegExp->put_IgnoreCase(vb);
            if (FAILED(hr))
            {
               logErrorMsg(hr, funcName, 0);
               retval = Error_Failure;
            }

            vb = VARIANT_FALSE;
            hr = m_iRegExp->put_Global(vb);
            if (FAILED(hr))
            {
               logErrorMsg(hr, funcName, 0);
               retval = Error_Failure;
            }

            if (retval == Error_Success)
               m_initialized = true;
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

   return retval;
}

long SMRegExp::setGlobal(bool b)
{
   static const char *funcName = "SMRegExp::setGlobal()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmRegexpInit, funcName, 0);
   }
   else
   {
      try
      {
         VARIANT_BOOL vb;

         if (b)
            vb = VARIANT_TRUE;
         else
            vb = VARIANT_FALSE;

         HRESULT hr = m_iRegExp->put_Global(vb);
         if (FAILED(hr))
         {
            logErrorMsg(hr, funcName, 0);
         }
         else
         {
            m_global = b;
            retval = Error_Success;
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
   }

   return retval;
}

long SMRegExp::setIgnoreCase(bool b)
{
   static const char *funcName = "SMRegExp::setIgnoreCase()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmRegexpInit, funcName, 0);
   }
   else
   {
      try
      {
         VARIANT_BOOL vb;

         if (b)
            vb = VARIANT_TRUE;
         else
            vb = VARIANT_FALSE;

         HRESULT hr = m_iRegExp->put_IgnoreCase(vb);
         if (FAILED(hr))
         {
            logErrorMsg(hr, funcName, 0);
         }
         else
         {
            m_ignoreCase = b;
            retval = Error_Success;
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
   }

   return retval;
}

long SMRegExp::setPattern(const char *p)
{
   static const char *funcName = "SMRegExp::setPattern()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmRegexpInit, funcName, 0);
   }
   else
   {
      try
      {
         HRESULT hr = m_iRegExp->put_Pattern(_bstr_t(p));
         if (FAILED(hr))
         {
            logErrorMsg(hr, funcName, 0);
         }
         else
         {
            retval = Error_Success;
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
   }
   return retval;
}

long SMRegExp::exec(const char *s)
{
   static const char *funcName = "SMRegExp::exec(const char *)";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmRegexpInit, funcName, 0);
   }
   else
   {
      m_index = 0;
      m_lastIndex = 0;
      m_found = true;

      delete m_input;
      m_input = 0;

      delete m_value;
      m_value = 0;

      unsigned long ls = strlen(s);
      m_input = new char[ls+1];
      if (m_input == 0)
      {
         logErrorMsg(errSmmResource, funcName, "Input string allocation.");
      }
      else
      {
         if (ls > 0)
            strcpy(m_input, s);
         else
            m_input[0] = '\0';

         retval =  exec();
      }
   }
   return retval;
}

class LocalPointersExec
{
   public:
      IDispatch *pDisp;
      IMatch2 *pMatch;
      IMatchCollection2 *pMatches;

      LocalPointersExec() : pDisp(0), pMatch(0), pMatches(0) {}
      ~LocalPointersExec() { if (pMatch) pMatch->Release();
                             if (pMatches) pMatches->Release();
                             if (pDisp) pDisp->Release(); }
};

long SMRegExp::exec()
{
   static const char *funcName = "SMRegExp::exec(void)";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmRegexpInit, funcName, 0);
   }
   else
   {
      LocalPointersExec LP;

      try
      {
         const char *ss = (const char *) (m_input + m_lastIndex);
         _bstr_t bss(ss);

         HRESULT hr = m_iRegExp->Execute(bss, &LP.pDisp);
         if (FAILED(hr))
         {
            if (hr != badRegularExpression)
               logErrorMsg(hr, funcName, "RegExp->Execute()");
            return retval;
         }

         hr = LP.pDisp->QueryInterface(__uuidof(IMatchCollection2),
                           reinterpret_cast<void **>(&LP.pMatches));
         if (FAILED(hr))
         {
            logErrorMsg(hr, funcName,
               "RegExp->QueryInterface(IMatchCollection2)");
            return retval;
         }

         LP.pDisp->Release();
         LP.pDisp = 0;

         long fCount = 0;

         hr = LP.pMatches->get_Count(&fCount);
         if (FAILED(hr))
         {
            logErrorMsg(hr, funcName, "RegExp->MatchCollection->get_Count()");
            return retval;
         }

         if (fCount > 0)
         {
            m_found = true;
            hr = LP.pMatches->get_Item(0, &LP.pDisp);
            if (FAILED(hr))
            {
               logErrorMsg(hr, funcName,
                     "RegExp->MatchCollection->get_Item(0)");
               return retval;
            }

            hr = LP.pDisp->QueryInterface(__uuidof(IMatch2),
                        reinterpret_cast<void **>(&LP.pMatch));
            if (FAILED(hr))
            {
               logErrorMsg(hr, funcName,
                  "RegExp->MatchCollection2->QueryInterface(IMatch2)");
               return retval;
            }

            LP.pDisp->Release();
            LP.pDisp = 0;

            long fIndex, fLength;

            hr = LP.pMatch->get_FirstIndex(&fIndex);
            if (FAILED(hr))
            {
               logErrorMsg(hr, funcName,
                     "RegExp->MatchCollection2->Match2->get_FirstIndex()");
               return retval;
            }

            hr = LP.pMatch->get_Length(&fLength);
            if (FAILED(hr))
            {
               logErrorMsg(hr, funcName,
                     "RegExp->MatchCollection2->Match2->get_Length()");
               return retval;
            }

            BSTR fValue;

            hr = LP.pMatch->get_Value(&fValue);
            if (FAILED(hr))
            {
               logErrorMsg(hr, funcName,
                     "RegExp->MatchCollection2->Match2->get_Value()");
               return retval;
            }

            if (m_value)
            {
               delete m_value;
               m_value = 0;
            }

            _bstr_t bv(fValue, false);
            unsigned long ln = bv.length();
            m_value = new char[ln+1];
            if (m_value)
            {
               strcpy(m_value, (const char *) bv);
            }

            m_index = fIndex + m_lastIndex;
            m_lastIndex = m_index + fLength;
            retval = fCount;         // number of matches
         }
         else
         {
            m_index = 0;
            m_lastIndex = 0;

            retval = Error_Success;
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
   }
   return retval;
}

// this function is for when you only want to know if there is a match
// of the pattern in the string and nothing else

long SMRegExp::match(const char *s)
{
   static const char *funcName = "SMRegExp::match(const char *)";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmRegexpInit, funcName, 0);
   }
   else
   {
      m_index = 0;
      m_lastIndex = 0;
      m_found = true;

      delete m_input;
      m_input = 0;

      unsigned long ls = strlen(s);
      m_input = new char[ls+1];
      if (m_input == 0)
      {
         logErrorMsg(errSmmResource, funcName, "Input string allocation.");
      }
      else
      {
         if (ls > 0)
            strcpy(m_input, s);
         else
            m_input[0] = '\0';

         long fStatus = exec();
         if (fStatus > 0)
         {
            if (m_index == 0 && m_value)
               retval = (stricmp(s, m_value) == 0);
         }
         else
         {
            retval = fStatus;
         }
      }
   }
   return retval;
}

long SMRegExp::test(const char *s)
{
   static const char *funcName = "SMRegExp::test()";
   SMMFcnTrace smmFcnTrace(funcName,logTraceLevel);

   long retval = Error_Failure;

   if (! m_initialized)
   {
      logErrorMsg(errSmmRegexpInit, funcName, 0);
   }
   else
   {
      try
      {
         _bstr_t bss(s);
         VARIANT_BOOL vb;

         HRESULT hr = m_iRegExp->Test(bss, &vb);
         if (FAILED(hr))
         {
            logErrorMsg(hr, funcName, "RegExp->Test()");
            return retval;
         }

         if (vb == VARIANT_TRUE)
            retval = 1;
         else
            retval = 0;
      }
      catch (_com_error &e)
      {
         logErrorMsg(e.Error(), funcName, e.ErrorMessage());
      }
      catch (...)
      {
         logErrorMsg(errSmmUnknownException, funcName, 0);
      }
   }
   return retval;
}
