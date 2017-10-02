#ifndef ESERVER_H
#define ESERVER_H

#include <comdef.h>

#import "statitpool.tlb" no_namespace, raw_interfaces_only

const unsigned int eServerProgIdLen = 32;
const unsigned int eServerErrBufLen = 4096;
const unsigned int eServerSysErrBufLen = 1024;

class eServer
{
   public:
      eServer(bool useServerPool);
      virtual ~eServer();

   public:
      long init();

      long doesVarExist(const char *var, bool& itExists);
      long runCmd(const char *cmd);
      long putGlobalInt(const char *var, long ival);
      long putGlobalString(const char *var, const char *stringVal);
      long putGlobalDouble(const char *var, double dval);
      long getGlobalInt(const char *var, long& ival);
      long getGlobalDouble(const char *var, double& dval);
      long getGlobalString(const char *var, char *& strValue);
      long getStringVar(const char *var, unsigned long& nCases,
                  char **& strArray);
      long getDoubleVar(const char *var, unsigned long& nCases,
                  double *& dblArray);
      long getString(const char *var, unsigned long nRow, char *& strValue);
      long getValue(const char *var, unsigned long nRow, double& dblValue);

      long putValue(const char *var, long nRow, double dblValue);
      long putString(const char *var, long nRow, const char *strValue);

      long createDoubleVar(const char *var);
      long createStringVar(const char *var, unsigned long nChars);

   protected:
      bool m_useServerPool;
      bool m_initialized;

      char m_serverProgId[eServerProgIdLen+1]; // for e-Server or the pool
      char *m_errBuf;

      IStatit *m_iStatit;

   protected:
      char *getErrors(HRESULT hr);
};


#endif
