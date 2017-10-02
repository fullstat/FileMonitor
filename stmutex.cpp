#define STMUTEX_MAIN
#include "stdafx.h"
#include "stmutex.h"

#define MUTEX_WAIT 10000L

// helper mutex class

STMutex::STMutex(BOOL doInit)
{
   didInit0 = FALSE;
   hMutex0 = 0;

   if (doInit)
   {
      hMutex0 = CreateMutex(NULL,FALSE,NULL);
      if (hMutex0)
         didInit0 = TRUE;
   }
}

STMutex::STMutex()
{
   didInit0 = FALSE;
   hMutex0 = 0;
}

STMutex::~STMutex()
{
   if (didInit0)
      CloseHandle(hMutex0);
}

BOOL STMutex::init()
{
   if (! didInit0)
   {
      hMutex0 = CreateMutex(NULL,FALSE,NULL);
      if (hMutex0)
         didInit0 = TRUE;
   }

   return didInit0;
}

BOOL STMutex::get()
{
   BOOL retval = FALSE;

   if (! didInit0)
      return retval;

   bool done = false;

   while (! done)
   {
      DWORD waitResult = MsgWaitForMultipleObjectsEx(1, &hMutex0,
                           MUTEX_WAIT, QS_ALLINPUT, MWMO_ALERTABLE);
      if (waitResult == WAIT_OBJECT_0)
      {
         done = true;
         retval = TRUE;
      }
      else if (waitResult = WAIT_OBJECT_0+1)
      {
         MSG msg;
         memset(&msg, 0, sizeof(MSG));

         while ((! done) && PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
         {
            // If it's a quit message, we're out of here.
            if (msg.message == WM_QUIT)
               done = true;
            else    // Otherwise, dispatch the message.
               DispatchMessage(&msg);
         }
      }
      else if (waitResult == WAIT_TIMEOUT || waitResult == WAIT_ABANDONED)
      {
         retval = FALSE;
         done = true;
      }
      else
      {
         retval = TRUE;
         done = true;
      }
   }

   return retval;
}

BOOL STMutex::release()
{
   if (! didInit0)
      return FALSE;

   if (! ReleaseMutex(hMutex0))
      return FALSE;

   return TRUE;
}
