#ifndef BVECTOR_H
#define BVECTOR_H
#include <stdlib.h>  // for size_t
#include <assert.h>

#ifdef VERBOSE
#include <iostream.h>
#endif

typedef unsigned char Byte;


// Working base.  Not directly instantiatable.
class BVectorBase
{
 public:
  ~BVectorBase();

  // The logical magnitude of the vector.
  size_t length() const { return length0; }
  int setLength (size_t);

  // The actual allocated size; allocateProc & deallocateProc wired into this.
  size_t capacity() const { return alSize0; }
  int setCapacity (size_t);

  // The vector bounds.  length may not exceed this.
  size_t maxLength() const { return maxSize0; }
  int setMaxLength (size_t);

  // The size of each element.
  size_t nodeSize() const { return nodeSize0; }

  // The first element of the vector.
  Byte * data() const { return data0; }

  Byte * nodeSafe (size_t i);
  Byte * nodeFast (size_t i) { return data() + i * nodeSize(); }

#ifdef NDEBUG
  void invariant() const {}
  void precondition() const {}   void postcondition() const {}
#else
  void invariant() const;
  //// User error //////////     //// BVector error ////////
  void precondition() const;     void postcondition() const;
#endif

 protected:
  // The node size, initial length, and a length limit (limit can be INT_MAX).
  BVectorBase (size_t nodeSize, size_t initialLength, size_t maxLength);

  virtual void allocateProc (size_t, size_t);
  virtual void deallocateProc (size_t, size_t);

 private:
  Byte * data0;      // The heap pointer.
  size_t length0,    // Extent of the users abstract vector.
         alSize0,    // Extent of heap allocation.
	 maxSize0;   // Maximum extent of vector.
  size_t nodeSize0;  // A vector node's size.

  friend class SVectorBase;
};


template <class T>
class BVector : public BVectorBase
{
 public:
  BVector (size_t len = 0, size_t s = 0) : BVectorBase (sizeof (T), len, s) {}
  ~BVector() { setCapacity (0); }

#ifdef COMPILE_INSTANTIATE
  int append(const T& v)
  {
     register int retval = setLength(length() + 1);
     if (retval == 0)
        (* this)[length() - 1] = v;
     return (retval);
  }
#else
  int append (const T & v);
#endif

  // Bounds checking reference indexer implemented as a function call.
  T & operator() (size_t i)
  {
    assert (i <= maxLength());
    precondition();
    return * (T *) nodeSafe (i);
  }

  // Inline value indexer without bounds checking (fast).
  T operator[] (size_t i) const
  {
    assert (i <= maxLength());
    precondition();
    return * ((T *) data() + i);
  }

  // Inline reference indexer without bounds checking (fast).
  T & operator[] (size_t i)
  {
    assert (i <= maxLength());
    precondition();
    return * ((T *) data() + i);
  }

};

//// BVector iterators: ///////////////////////////////////////////////////////


class BVIBase              // Not directly instantiatable.
{
 protected:
  Byte        * node0,     // Current position.
              * stopNode0; // Last valid position.
  BVectorBase * vector0;   // Object of iteration.

  BVIBase (BVectorBase * v) : vector0 (v), node0 (0), stopNode0 (0) {}

  Byte * scanForward();
  Byte * scanBackward();

 public:
  // Object of iteration.
  BVectorBase * vector() { return vector0; }

  // Reset the iterator to its initial state (nil).
  void reset() { node0 = stopNode0 = 0; }
};


template <class T>
class BVectorIterator : public BVIBase
{
 public:
  BVectorIterator (BVector <T> * v) : BVIBase (v) {}
#ifdef COMPILE_INSTANTIATE
  BVectorIterator(BVector <T> *v, size_t start) : BVIBase(v)
  {
     if (start == 0) {
        node0 = 0;
        stopNode0 = 0;
     } else {
        node0 = v->nodeFast(start-1);
        stopNode0 = v->nodeFast(v->length() - 1);
     }
  }
#else
  BVectorIterator (BVector <T> * v, size_t start);
#endif

  // The iterator function.  Advance; if valid, return node else return nil.
  T * operator()()
  {
    if (node0 < stopNode0)
      return (T *) (node0 += sizeof (T));
    else
      return (T *) scanForward();
  }

  T * currentNode() { return (T *) node0; }
};


template <class T>
class BVectorBackIterator : public BVIBase
{
 public:
  BVectorBackIterator (BVector <T> * v) : BVIBase (v) {}

  // The iterator function.
  T * operator()()
  {
    if (node0 > stopNode0)
      return (T *) (node0 -= sizeof (T));
    else
      return (T *) scanBackward();
  }

  T * currentNode() { return (T *) node0; }
};

#endif
