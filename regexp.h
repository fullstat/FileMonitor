
#ifndef _REGEXP_H_
#define _REGEXP_H_

/*
 * this class is a simple wrapper around the VbScript.RegExp object.
 * the underlying support of the RegExp object is limited since all
 * I want to determine is if there is a pattern match.  other C++
 * implmentations of regular expressions that I checked were not
 * thread safe.  this one also allows case to be ignored which is
 * quite useful when working with file names on Windows.
 *
 * the type library for regular expressions is the 3rd in the vbscript.dll
 * resource block.  the compiler #import directive does not support the
 * syntax to load it (trailing \3). therefore, I used the visual studio
 * object viewer which does allow you to view the RegExp type information.
 * I saved this to a local regexp.idl file and used midl to generate the
 * regexp.tlb type library.
 *
 * if you want to test a string for an exact match, use the match()
 * method.  it does the test and verifies that the matched string is
 * the same as the input string.
 *
 * to use the class do the following:
 *
 *  SMRegExp re;     // create an instance
 *  re.init();       // initialize
 *  re.setPattern("LOT[0-9][0-9]*");   // set the pattern
 *  if (re.match("LOT1234") > 0)       // test a string
 *
 * there is limited support for finding information about the match.
 * if you want this, use exec() instead of find and reference the index()
 * and lastIndex() methods.  or even better, add more support to this class.
 */

#import "regexp.tlb" raw_interfaces_only, no_namespace

class SMRegExp
{
   public:
      SMRegExp();
      virtual ~SMRegExp();

      bool global() { return m_global; }
      bool ignoreCase() { return m_ignoreCase; }

      long index() { return m_index; }
      long lastIndex() { return m_lastIndex; }

      long setGlobal(bool b);
      long setIgnoreCase(bool b);
      long setPattern(const char *p);

      long init();
      long exec(const char *s);
      long exec();

      long match(const char *s);
      long test(const char *s);

   protected:
      IRegExp2 *m_iRegExp;
      char *m_input;
      char *m_value;

      bool m_initialized;
      long m_index;
      long m_lastIndex;
      bool m_found;
      bool m_global;
      bool m_ignoreCase;
};

#endif
