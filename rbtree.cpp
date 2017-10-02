#include "rbtree.h"

/*   (red-black tree class)
 *- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 *
 *   this is a stripped down version of the class to avoid the I/O, and Tools.h++
 *   dependency.
 *
 *- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
 */

// RBitem

RBitem::RBitem()
{
}

RBitem::~RBitem()
{
}

int RBitem::compareTo(const RBitem *c) const
{
   return this == c ? 0 : (this > c ? 1 : -1);
}

bool RBitem::isEqual(const RBitem *c) const
{
   return this == c;
}


// Queue definitions

int Queue::push(node *n)
{
   qnode *push;

   if (freefront == 0)
   {
      if ((push = new qnode) == 0)
         return (-1);
   }
   else
   {
      push = freefront;
      freefront = freefront->next;
      if (freeback == push)
         freeback = 0;
   }

   push->n = n;
   push->next = 0;

   if (front == 0)
   {
      front = push;
      back = push;
   }
   else
   {
      back->next = push;
      back = push;
   }

   return 0;
}

node * Queue::pop()
{
   qnode *pop;

   if (front == 0)
      return 0;

   pop = front;
   front = front->next;
   if (back == pop)
      back = 0;

   // put freed node onto free list.
   pop->next = 0;
   if (freefront == 0)
   {
      freefront = pop;
      freeback = pop;
   }
   else
   {
      freeback->next = pop;
      freeback = pop;
   }

   return pop->n;
}

Queue::~Queue()
{
   qnode *t = freefront;
   while (t != 0)
   {
      t = t->next;
      delete freefront;
      freefront = t;
   }

   t = front;
   while (t != 0)
   {
      t = t->next;
      delete front;
      front = t;
   }
}


// RBTree definitions

node * RBTree::rotate(RBitem *i, node *y)
{
   node *c, *gc;
   if (y == head)
      c = y->right;
   else if (y->item->compareTo(i) > 0)
      c = y->left;
   else
      c = y->right;

   if (c->item->compareTo(i) > 0)
   {
      gc = c->left;
      c->left = gc->right;
      gc->right = c;
   }
   else
   {
      gc = c->right;
      c->right = gc->left;
      gc->left = c;
   }

   if (y == head)
      y->right = gc;
   else if (y->item->compareTo(i) > 0)
      y->left = gc;
   else
      y->right = gc;
   return gc;
}

void RBTree::split(RBitem *i, node *x, node *g, node *p, node *gg)
{
   x->link = RBT_RED;
   x->left->link = RBT_BLACK;
   x->right->link = RBT_BLACK;

   if (p->link)
   {
      int gcomp = g->item->compareTo (i);
      int pcomp = p->item->compareTo (i);

      g->link = RBT_RED;
      if (gcomp > 0 && pcomp < 0)
         rotate (i,g);
      else if (gcomp < 0 && pcomp > 0)
         rotate (i,g);
      x = rotate(i,gg);
      x->link = RBT_BLACK;
   }
}


RBitem * RBTree::insertInto(node *n)
{
   RBitem *i = n->item;

   if (z == 0 || head == 0)
   {
      z = new node;
      if (z == 0)
         return 0;
      head = new node;
      if (head == 0)
         return 0;

      z->item = 0;
      z->left = z;
      z->right = z;
      z->link = RBT_BLACK;

      head->item = 0;
      head->left = 0;
      head->right = z;
      head->link = RBT_BLACK;
   }

   node *gg = head;
   node *x = head;
   node *p = head;
   node *g = head;

   // find place to insert i
   while (x != z)
   {
      if (x == head)
      {
         x = x->right;
      }
      else
      {
         gg = g;
         g = p;
         p = x;

         int compare = x->item->compareTo (i);

        // return error if item is already in the tree.
         if ( !compare)
            return 0;

         if (compare > 0)
            x = x->left;
         else if (compare < 0)
            x = x->right;
      }

      // if a 4 node is encountered split it
      if (x->left->link && x->right->link)
      {
         split(i,x,g,p,gg);
         head->right->link = RBT_BLACK;
      }
   }

   x = n;
   x->left = z;
   x->right = z;
   x->link = RBT_RED;

   // if the tree is not empty
   if (head->right != z)
   {
      if (p->item->compareTo(i) > 0)
         p->left = x;
      else
         p->right = x;
      split(i,x,g,p,gg);
      head->right->link = RBT_BLACK;
   }
   else
   {
      head->right = x;
      head->right->link = RBT_BLACK;
   }

   entries0++;
   return i;
}

RBitem * RBTree::insert (RBitem *i)
{
   node *n;
   if ((n = new node) == 0)
      return 0;
   n->item = i;
   n->link = RBT_BLACK;
   n->left = 0;
   n->right = 0;
   RBitem *ret = insertInto (n);
   return (ret);
}


RBitem * RBTree::find(const RBitem *i) const
{
   if (head == 0)
      return 0;
   node *x = head->right;

   while (x != z)
   {
      int compare = x->item->compareTo (i);
      if (compare == 0)
         return x->item;
      if (compare > 0)
         x = x->left;
      else
         x = x->right;
   }

   if (x == z)
      return 0;

   return x->item;
}


RBitem * RBTree::remove (const RBitem *i)
{
  if (head == 0)
    return 0;

  if (entries0 == 0 || head->right == z)
    return 0;

  // p=parent, g=grandparent, gg=great grandparent, t = node to be removed
  // c= index for searching, x= index for searching.
  node *c, *p, *x, *t, *g, *gg, *parent;

  p = head; x = head->right; g = head; gg = head; parent = head;

  // if the first item is a 4-node split it.
  if (x->right->link && x->left->link)
  {
    split (x->left->item, x, g, p, gg);
    head->right->link = RBT_BLACK;
    return remove (i);
  }

  // search for i
  while (x != z && x->item->compareTo(i) !=0)
  {
    gg = g; g = p; p = x; parent = x;

    if (x->item->compareTo(i) > 0)
      x = x->left;
    else
      x = x->right;

    // if a 4-node is encountered split it.
    if (x->right->link && x->left->link)
    {
      split (x->left->item, x, g, p, gg);
      head->right->link = RBT_BLACK;
      return remove (i);
    }
  }

  // return 0 if i not found
  if (x == z)
    return 0;

  t = x;

  // if t is the right item of a 3 node.
  if ( (t->link && p->item->compareTo(t->item) < 0) || t->left->link)
  {
    // if t is a terminal node.
    if ( t->link && t->left == z && t->right == z)
      x = z;
    else if (t->link == RBT_BLACK && t->right == z &&
             t->left->left == z && t->left->right == z)
    {  x = x->left; x->link = RBT_BLACK; }

    // if there is more than one item in right node
    // search for node furthest left down right side.
    // to replace node being removed.
    else if (t->right->left != z || t->right->right != z)
    {
      if (t->right->left == z)
      {
        x = x->right; x->left = t->left; x->link = t->link;
        x->right->link = RBT_BLACK;
      }
      else
      {
        c = x->right;
        while (c->left->left != z)
          c = c->left;
        x = c->left;
        x->right->link = RBT_BLACK;
        c->left = x->right;
        x->left = t->left;
        x->right = t->right;
        x->link = t->link;
      }
    }

    // if there is more than one item in left node.
    // search for node furthest right down left side.
    else if ((t->link && (t->left->left != z || t->left->right != z)) ||
              ((t->link == RBT_BLACK) &&
               (t->left->right->left !=z || t->left->right->right != z)))
    {
      if (t->link && t->left->right == z)
      {
        x = x->left; x->right = t->right; x->link = t->link;
        x->left->link = RBT_BLACK;
      }
      else if (t->link == RBT_BLACK && t->left->right->right == z)
      {
        x = t->left->right; x->right = t->right; x->link = t->link;
        x->left->link = RBT_BLACK;
      }
      else
      {
        if (t->link)
          c = x->left;
        else
         c = x->left->right;
        while (c->right->right != z)
          c = c->right;
        x = c->right;
        x->left->link = RBT_BLACK;
        c->right = x->left;
        x->left = t->left;
        x->right = t->right;
        x->link = t->link;
      }
    }
    // if there is only one item in left and right reduce the 3 node
    // to a 2 node.
    else
    {
      x = x->left;
      if (t->link)
      {
        //x = x->left;
        x->right = t->right;
        x->right->link = RBT_RED;
        //x->link = RBT_BLACK;
      }
      else
      {
        //x = x->left;
        x->right->right = t->right;
        x->right->right->link = RBT_RED;
      }
      x->link = RBT_BLACK;
    }
    if (parent == head)
      parent->right = x;
    else
    {
      if (parent->item->compareTo(i) >0)
        parent->left = x;
      else
        parent->right = x;
    }
  }

  // if item to be removed is the left end of a 3 node.
  else if (t->right->link || (t->link && p->item->compareTo(t->item) > 0) )
  {
    // if item is a terminal node.
    if (t->link && t->left == z && t->right == z)
      x = z;
    else if (t->link == RBT_BLACK && t->left == z &&
             t->right->left == z && t->right->right == z)
    {
      x = x->right; x->link = RBT_BLACK;
    }

    // if there is more than one item on left search for the
    // item furthest right down left side.
    else if ( t->left->left != z || t->left->right != z)
    {
      if (t->left->right == z)
      {
        x = x->left; x->right = t->right; x->link = t->link;
        x->left->link = RBT_BLACK;
      }
      else
      {
        c = x->left;
        while (c->right->right != z)
          c = c->right;
        x = c->right;
        x->left->link = RBT_BLACK;
        c->right = x->left;
        x->left = t->left;
        x->right = t->right;
        x->link = t->link;
      }
    }

    // if there is more than one item on right search for the
    // item furthest left down right side.
    else if ( (t->link && (t->right->left != z || t->right->right != z)) ||
            (t->link == RBT_BLACK && (t->right->left->left !=z ||
                                   t->right->left->right != z)))
    {
      if (t->link && t->right->left == z)
      {
        x = x->right; x->left = t->left; x->link = t->link;
        x->right->link = RBT_BLACK;
      }
      else if (t->link == RBT_BLACK && t->right->right->left == z)
      {
        x = x->right->left; x->left = t->left; x->link = t->link;
        x->right->link = RBT_BLACK;
      }
      else
      {
        if (t->link)
          c = x->right;
        else
          c = x->right->left;
        while (c->left->left != z)
          c = c->left;
        x = c->left;
        x->right->link = RBT_BLACK;
        c->left = x->right;
        x->left = t->left;
        x->right = t->right;
        x->link = t->link;
      }
    }

    // If there is only one item in left and right nodes reduce
    // the 3-node to a 2-node.
    else
    {
      if (t->link)
      {
        x = x->left;
        x->right = t->right;
        x->right->link = RBT_RED;
        x->link = RBT_BLACK;
      }
      else if (t->right->link)
      {
        x = x->right;
        x->left->left = t->left;
        x->left->left->link = RBT_RED;
        x->link = RBT_BLACK;
      }
    }
    if (parent == head)
      parent->right = x;
    else
    {
      if (parent->item->compareTo(i) >0)
        parent->left = x;
      else
        parent->right = x;
    }
  }

  // if item to be removed is a 2 node.
  else
  {
    // if parent is a 3 node && t is a terminal node.
    if (p != head && (p->link || p->right->link || p->left->link) &&
                         t->right == z && t->left == z )
    {
      if (p->right->link)
      {
        if (g->right == p)
          g->right = p->right;
        else if (g->left == p)
          g->left = p->right;
        p->right->link = RBT_BLACK;
        insert (p->item);
        entries0--;
        delete p;
      }
      else if (p->left->link)
      {
        if (g->right == p)
          g->right = p->left;
        else if (g->left == p)
          g->left = p->left;
        p->left->link = RBT_BLACK;
        insert (p->item);
        entries0--;
        delete p;
      }
      else if (p->link)
      {
        if (g->right == p)
        {
          if (p->left == t)
            g->right = p->right;
          else
            g->right = p->left;
        }
        else if (g->left == p)
        {
          if (p->left == t)
            g->left = p->right;
          else
            g->left = p->left;
        }
        insert (p->item);
        entries0--;
        delete p;
      }
    }


    // If t is not a terminal node with a 3 node parent remove to and
    // replace with least item bigger than t.
    else
    {
      if (t->right == z)
      {
        x = x->left;
        x->link = RBT_BLACK;
      }
      else if (t->right->left == z)
      {
        x = x->right;
        x->left = t->left;
        x->link = t->link;
        x->right->link = RBT_BLACK;
      }
      else
      {
         c = x->right; gg = g; g = p; p = x;
         while (c->left->left != z)
           {c = c->left; gg = g; g = p; p = c;}

         x = c->left; gg = g; g = p; p = c;
         x->right->link = RBT_BLACK;

         c->left = x->right;
         //c->left->link = x->link;
         x->left = t->left;
         x->right = t->right;
         x->link = t->link;
       }
       if (parent == head)
         parent->right = x;
       else
       {
         if (parent->item->compareTo(i) >0)
           parent->left = x;
         else
           parent->right = x;
       }
    }

  }  // end else if t is a 2 node.

  RBitem *ret =  t->item;
  delete t;

  head->right->link = RBT_BLACK;

  entries0--;
  return ret;
}

RBitem * RBTree::removeAndBalance (const RBitem *i)
{
  if (head == 0 || head == z || head->right == z)
    return 0;

  entries0 = 0;

  Queue q;
  q.push (head->right);

  head->right = z;

  RBitem *ret = 0;
  node *t = q.pop();

  while (t != 0)
  {
    if (t->left != z)
      q.push (t->left);
    if (t->right != z)
      q.push (t->right);

    if (i->compareTo (t->item) == 0)
      {ret = t->item; delete t;}
    else
      insertNode (t);
      //insert (t->item);
    //delete t;
    t = q.pop();
  }
  return ret;
}


int RBTree::destroyTree(node *head, node *z)
{
  if (head == 0)
  {
    delete head;
    delete z;
    return 0;
  }

  Queue q;

  if (head->right != z)
  {
    q.push(head->right);

    node *t = q.pop();
    while (t != 0)
    {
      if (t->left != z)
        q.push(t->left);
      if (t->right != z)
        q.push(t->right);
      delete t->item;
      delete t;
      t = q.pop();
    }
  }
  delete head;
  delete z;
  return 0;
}


int RBTree::balance ()
{
  if (head == 0 || head == z || head->right == z)
    return 1;

  entries0 = 0;
  Queue q;
  q.push (head->right);

  head->right = z;

  node *t = q.pop();

  while (t != 0)
  {
    if (t->left != z)
      q.push (t->left);
    if (t->right != z)
      q.push (t->right);
    insertNode (t);
    t = q.pop();
  }
  return 0;
}


void RBTree::clearAll1(node *t, node *z)
{
  if (t != z)
  {
    clearAll1 (t->left, z);
    node *temp = t->right;
    delete t->item;
    delete t;
    clearAll1 (temp, z);
  }
}

void RBTree::clearAll2 (node *t, node *z)
{
  if (t!=z)
  {
    clearAll2 (t->left,z);
    node *temp = t->right;
    delete t;
    clearAll2 (temp, z);
  }
}

void RBTree::clearAndDestroy ()
{
  if (head != 0)
  {
    clearAll1(head->right,z);
    head->right = z;
  }
}

void RBTree::clear()
{
  if (head != 0)
  {
    clearAll2 (head->right, z);
    head->right = z;
  }
}


RBTree::~RBTree()
{
  int err = destroyTree (head, z);
}

