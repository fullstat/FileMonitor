#ifndef BVEC_H
#define BVEC_H

#include "bvector.h"

typedef BVector <char> BCharVec;
typedef BVector <int> BIntVec;
typedef BVector <double> BDoubleVec;

typedef BVector <char *> BCharPtrVec;
typedef BVector <int *> BIntPtrVec;
typedef BVector <double *> BDoublePtrVec;
typedef BVector <char **> BCharPtrPtrVec;

typedef BVector <BCharPtrVec *> BCharPtrVecArr;
typedef BVector <BIntVec *> BIntVecArr;
typedef BVector <BDoubleVec *> BDoubleVecArr;

inline char   *cp(BCharVec& c, int offset=0) { return (char *) &c[offset]; }
inline int    *ip(BIntVec& i, int offset=0) { return (int *) &i[offset]; }
inline double *dp(BDoubleVec& d, int offset=0) { return (double *) &d[offset]; }

#endif
