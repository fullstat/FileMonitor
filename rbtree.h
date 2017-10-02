#ifndef RBTREE_H
#define RBTREE_H

/*   (red-black tree class)
 *- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 *
 *   this is a stripped down version of the class to avoid the I/O, and Tools.h++
 *   dependency.
 *
 *- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 */

class RBitem
{
   public:
      RBitem();
      virtual ~RBitem();

      virtual int compareTo(const RBitem *) const;
      virtual bool isEqual(const RBitem *) const;
};

struct node
{
   RBitem *item;
   node *left;
   node *right;
   int link;
};

struct qnode
{
   node *n;
   qnode *next;
};


class Queue
{
   protected:
      qnode *front, *back, *freefront, *freeback;

   public:
      Queue () : front(0), back(0), freefront(0), freeback(0) {}
      ~Queue();

      int push(node *);
      node * pop();
      int isEmpty() 
      {
         if (front == 0) 
            return 1;
         else
            return 0;
      }
};

class RBTree
{
   protected:  
      node *head;
      node *z;
      int entries0;
     
   protected:
      node * rotate(RBitem *, node *);
      void split(RBitem *, node *, node *, node *, node *);
      int destroyTree (node*, node*);
      void clearAll1 (node*, node*);
      void clearAll2 (node*, node*);
     
   public:
      RBTree() : head(0), z(0), entries0(0) {}
      ~RBTree();
     
      RBitem * insert (RBitem *);
      RBitem * insertNode (node *n) { return insertInto (n); }
      RBitem * insertInto (node *);
      RBitem * find (const RBitem *) const;
      RBitem * remove (const RBitem *);
      RBitem * removeAndBalance (const RBitem *);
      int             balance ();
      size_t entries() const {return entries0;}
      bool isEmpty() const { if (head == 0 || head->right == z)
                                return true;  else return false; }
      void clearAndDestroy();
      void clear();
     
      void inorder();
      enum { RBT_BLACK = 0, RBT_RED = 1 };
};

#endif
