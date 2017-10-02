#include <limits.h>    // for INT_MAX
#include <string.h>    // for memset

#include "stdafx.h"
#include "bVector.h"

#ifndef NDEBUG

// Sanity check.
void BVectorBase :: invariant() const
{
  assert (length() <= capacity());
  assert (capacity() <= maxLength());
  assert (nodeSize() > 0);
}


// Called at member function entry.
void BVectorBase :: precondition() const
{
  invariant();
}


// Called prior to member function exit.
void BVectorBase :: postcondition() const
{
  invariant();
}

#endif



BVectorBase :: BVectorBase (size_t ns, size_t length, size_t maxSize)
             : data0 (0), alSize0 (0)
{
  // Default is an unbounded vector.
  if (maxSize == 0)
    maxSize = INT_MAX;

  // Sanity check and correction.
  if (length > maxSize)
    length = maxSize;

  assert (length <= maxSize);

  maxSize0  = maxSize;
  nodeSize0 = ns;
  length0   = 0;

  if (setCapacity (length) == 0)
    length0 = length;

  postcondition();
}



BVectorBase :: ~BVectorBase()
{
}



void BVectorBase :: allocateProc (size_t first, size_t last)
{
  assert (first <= last);
  precondition();

  // Zero out fresh memory.
  memset (data() + first * nodeSize(), 0, (last - first + 1) * nodeSize());
}



void BVectorBase :: deallocateProc (size_t, size_t)
{}


const int GrowthStep1 = 32;
const int GrowthStep2 = 256;
const int GrowthStep3 = 1024;
const int GrowthStep4 = 4096;

inline int GrowthRate1(size_t c) { return (c * 3) / 2; }
inline int GrowthRate2(size_t c) { return (c * 4) / 3; }
inline int GrowthRate3(size_t c) { return (c * 5) / 4; }
inline int GrowthRate4(size_t c) { return (c * 6) / 5; }
inline int GrowthRate5(size_t c) { return (c * 11) / 10; }

int BVectorBase :: setLength (size_t newLength)
{
  precondition();
  if (newLength > maxLength())
    return 1;

  register size_t retVal;

  if (newLength > capacity())
  {
    size_t testCap;

    if (capacity() <= GrowthStep1)
      testCap = GrowthRate1(capacity());
    else if (capacity() <= GrowthStep2)
      testCap = GrowthRate2(capacity());
    else if (capacity() <= GrowthStep3)
      testCap = GrowthRate3(capacity());
    else if (capacity() <= GrowthStep4)
      testCap = GrowthRate4(capacity());
    else
      testCap = GrowthRate5(capacity());
    if (testCap > newLength)
      retVal = setCapacity (testCap);
    else
      retVal = setCapacity (newLength);
    if (retVal != 0)
      goto END;
  }

  length0 = newLength;
  retVal  = 0;

END:
  postcondition();
  return retVal;
}



int BVectorBase :: setCapacity (register size_t newCap)
{
  assert (newCap <= maxLength());
  precondition();

  register int retVal;
  register size_t oldCap = capacity();

  if (newCap < oldCap)
  {
    if (newCap < length())
    {
      retVal = setLength (newCap);
      if (retVal)
        goto END;
    }

    deallocateProc (newCap, oldCap - 1);

    Byte * newData;
    if (newCap == 0)
    {
      free (data0);
      newData = 0;
    }
    else
    {
      if (data0 == 0)
        newData = (Byte *) malloc (newCap * nodeSize());
      else
        newData = (Byte *) realloc (data0, newCap * nodeSize());

      if (newData == 0)
      {
        retVal = -1;
        goto END;
      }
    }

    data0 = newData;
    alSize0 = newCap;
  }
  else
    if (newCap > oldCap)
    {
      Byte * newData;

      if (data0 == 0)
        newData = (Byte *) malloc (newCap * nodeSize());
      else
        newData = (Byte *) realloc (data0, newCap * nodeSize());

      if (newData == 0)
      {
        retVal = -1;
        goto END;
      }

      data0 = newData;
      alSize0 = newCap;

      allocateProc (oldCap, capacity() - 1);
    }

  // Success.
  retVal = 0;

END:
  postcondition();
  return retVal;
}



int BVectorBase :: setMaxLength (register size_t newMaxLength)
{
  precondition();

  if (newMaxLength < length())
    return 1;

  maxSize0 = newMaxLength;

  postcondition();
  return 0;
}



Byte * BVectorBase :: nodeSafe (register size_t i)
{
  precondition();

  if (i < length())
    RET1:
    return nodeFast (i);

  if (i >= maxLength())
    return 0;

  if (i >= capacity())
  {
    register size_t newCap = capacity();
    if (newCap == 0)
      newCap = 1;
    while (newCap <= i)
      newCap *= 2;

    if (newCap > maxLength())
      newCap = maxLength();

    register int retVal = setCapacity (newCap);
    if (retVal)
      return 0;
  }

  length0 = i+1;
  goto RET1;
}


//// BVectorIterator definitions: /////////////////////////////////////////////



Byte * BVIBase :: scanForward()
{
  register Byte * node = node0;

  if (node == 0)
  {
    register size_t len = vector0->length();
    if (len > 0)
    {
      node      = vector0->data();
      stopNode0 = node + (len - 1) * vector0->nodeSize();
    }
  }
  else
    node = 0;

  return node0 = node;
}



Byte * BVIBase :: scanBackward()
{
  register Byte * node = node0;

  if (node == 0)
  {
    register size_t len = vector0->length();
    if (len > 0)
    {
      stopNode0 = node = vector0->data();
      node     += (len - 1) * vector0->nodeSize();
    }
  }
  else
    node = 0;

  return node0 = node;
}
