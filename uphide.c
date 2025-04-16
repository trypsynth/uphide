// *************************************************
//      Made with BCX BASIC To C/C++ Translator
//            Version 8.2.5 (02/25/2025)
// *************************************************
//    Translated for compiling with a C Compiler
// *************************************************
#ifdef _MSC_VER
  #ifndef _CRT_SECURE_NO_WARNINGS
    #define _CRT_SECURE_NO_WARNINGS
  #endif
#endif

#ifndef _WIN32_DCOM
  #define _WIN32_DCOM
#endif

// Windows API headers
#include <windows.h>    // Core Windows API
#include <windowsx.h>   // Windows API macros and functions
#include <commctrl.h>   // Common Controls
#include <commdlg.h>    // Common Dialogs
#include <direct.h>     // Directory handling
#include <mmsystem.h>   // Multimedia
#include <oaidl.h>      // OLE Automation
#include <objbase.h>    // Component Object Model (COM)
#include <ocidl.h>      // OLE Control interfaces
#include <ole2.h>       // OLE 2.0 functionality
#include <oleauto.h>    // OLE Automation
#include <olectl.h>     // OLE Controls
#include <richedit.h>   // Rich Edit Control
#include <shellapi.h>   // Shell API
#include <shlobj.h>     // Shell Objects
#include <urlmon.h>     // URL Monikers
#include <wchar.h>      // Wide character support (also in ISO C)
#include <wctype.h>     // Wide character classification (also in ISO C)
#include <tchar.h>      // Unicode/ANSI character mapping
#include <unknwn.h>     // COM base interfaces
#include <wingdi.h>     // GDI
#include <wininet.h>    // Internet
#include <winsock.h>    // Windows Sockets
#include <winuser.h>    // User Interface

// ISO C Standard Library headers
#include <ctype.h>      // Character classification
#include <math.h>       // Mathematical functions
#include <setjmp.h>     // Non-local jumps
#include <stdarg.h>     // Variable arguments
#include <stddef.h>     // Common definitions
#include <stdio.h>      // Input/output
#include <stdlib.h>     // General utilities
#include <string.h>     // String handling
#include <time.h>       // Date and time
#include <errno.h>      // Error numbers (also POSIX)

// C99 Standard headers
#include <stdbool.h>    // Boolean type
#include <inttypes.h>   // Integer types

// POSIX headers
#include <fcntl.h>      // File control options

// Windows-specific headers
#include <process.h>    // Process control functions
#include <io.h>         // Low-level I/O (Windows POSIX subset)
#include <conio.h>      // Console I/O (Windows-specific)

// *************************************************
//            System Defined Macros
// *************************************************

#define BCXSTRSIZE 2048
#include <wuapi.h>

#ifndef    VK_0
   #define VK_0  0x30
   #define VK_1  0x31
   #define VK_2  0x32
   #define VK_3  0x33
   #define VK_4  0x34
   #define VK_5  0x35
   #define VK_6  0x36
   #define VK_7  0x37
   #define VK_8  0x38
   #define VK_9  0x39
   #define VK_A  0x41
   #define VK_B  0x42
   #define VK_C  0x43
   #define VK_D  0x44
   #define VK_E  0x45
   #define VK_F  0x46
   #define VK_G  0x47
   #define VK_H  0x48
   #define VK_I  0x49
   #define VK_J  0x4A
   #define VK_K  0x4B
   #define VK_L  0x4C
   #define VK_M  0x4D
   #define VK_N  0x4E
   #define VK_O  0x4F
   #define VK_P  0x50
   #define VK_Q  0x51
   #define VK_R  0x52
   #define VK_S  0x53
   #define VK_T  0x54
   #define VK_U  0x55
   #define VK_V  0x56
   #define VK_W  0x57
   #define VK_X  0x58
   #define VK_Y  0x59
   #define VK_Z  0x5A
#endif 

// *************************************************
//                   Microsoft VC++
// *************************************************

#ifndef DECLSPEC_UUID
  #if(_MSC_VER >= 1100) && defined(__cplusplus)
    #define DECLSPEC_UUID(x)  __declspec(uuid(x))
  #else
    #define DECLSPEC_UUID(x)
  #endif
#endif
#if (_MSC_VER >= 1900)            // earlier versions untested
   #include <intrin.h>
      #ifndef _rdtsc
         #define _rdtsc __rdtsc   // MSVC uses 2 underscores
      #endif
   #pragma warning(disable: 4018) // signed/unsigned mismatch warnings
   #pragma warning(disable: 4100) // unreferenced argument warnings
   #pragma warning(disable: 4244) // conversion from type1 to type2 warnings
   #pragma warning(disable: 4459) // local hides global declaration
   #pragma warning(disable: 4267) // conversion from type1 to type2 warnings
   #pragma warning(disable: 4305) // truncation from double to float warnings
   #pragma warning(disable: 4459) // hides global declaration warnings
   #pragma warning(disable: 4800) // forcing value to bool warnings
   #pragma warning(disable: 4838) // conversion from type1 to type2 warnings
   #pragma warning(disable: 4244) // conversion from type1 to type2 warnings
   #pragma warning(disable: 4245) // conversion from type1 to type2 warnings
   #pragma warning(disable: 4554) // use parentheses to clarify precedence
   #pragma warning(disable: 4389) // conversion from signed to unsigned warnings
#endif

// *************************************************
//                  GCC and CLANG
// *************************************************

#if defined(__GNUC__) || defined(__clang__)
   #ifndef __BCPLUSPLUS__
      #include <x86intrin.h>
   #endif
   #pragma GCC diagnostic ignored "-Wwrite-strings"
   #pragma GCC diagnostic ignored "-Wunused-parameter"
   #pragma GCC diagnostic ignored "-Wunknown-pragmas"
   #pragma GCC diagnostic ignored "-Wdangling-else"
   #pragma GCC diagnostic ignored "-Wdeprecated"
   #pragma GCC diagnostic ignored "-Wsizeof-pointer-div"
#endif

// *************************************************
//                  Embarcadero C++
// *************************************************

#if defined(__BCPLUSPLUS__)
      #include<malloc.h>
      #if defined (_clang__)
          #include <mmintrin.h>
      #endif
      #define _kbhit kbhit
      #ifndef _rdtsc
         #define _rdtsc __rdtsc  // Uses 2 underscores
      #endif
#endif

// *************************************************
//                     Lcc-Win32
// *************************************************

#if defined(__LCC__)
    #define _strdup    strdup
    #define _fseeki64  fseek
    #define _ftelli64  ftell
    #define _stricmp   stricmp
    #define _strnicmp  strnicmp
    #define _itoa      itoa
    #define _ltoa      ltoa
    #include <intrinsics.h>
    #include <malloc.h>  // for _msize
    #if defined(__windows_h__)
       #define COMPILE_MULTIMON_STUBS
       #include <multimon.h>
       #include <iehelper.h>
       #include <exdisp.h>
    #endif
#endif

// *************************************************
//                     Pelles C
// *************************************************

#if defined(__POCC__)
    #include <intrin.h>
    #pragma pack_stack(off)
    #pragma warn(disable: 2006)    // Non-portable conversion int to const char*
    #pragma warn(disable: 2007)    // Non-portable inline code
    #pragma warn(disable: 2073)    // Loss of precision float to int
    #pragma warn(disable: 2115)    // Initialized but not used warnings
    #pragma warn(disable: 2118)    // Unreferenced argument warnings
    #pragma warn(disable: 2134)    // Possible infinite loop
    #pragma warn(disable: 2154)    // Buggy unreachable code warning using sizeof
    #pragma warn(disable: 2197)    // Unsigned long int not a std bit-field type
    #pragma warn(disable: 2215)    // Conversion from type1 to type2 warnings
    #pragma warn(disable: 2218)    // Unreferenced parameter
    #pragma warn(disable: 2230)    // Incomplete struct declarations (vbs support)
    #pragma warn(disable: 2235)    // Not all control paths return a value
    #pragma warn(disable: 2241)    // Function marked for deprecation
    #pragma warn(disable: 2243)    // Use parentheses to clarify
    #pragma warn(disable: 2248)    // Non-portable use of extension
    #pragma warn(disable: 2251)    // Types with different signedness
    #pragma warn(disable: 2804)    // Consider changing type to size_t warnings
    #pragma warn(disable: 2805)    // Possible anti-aliasing violation warnings
    #pragma warn(disable: 2808)    // Unsequenced assignment warnings
    #pragma warn(disable: 2810)    // Potential realloc warnings
    #pragma warn(disable: 2812)    // Attempt to read from non-readable location
    #pragma warn(disable: 2813)    // Possible out-of-bounds warning (caused by CAST)
#endif

// *************************************************
// Instruct Linker to Search Object/Import Libraries
// *************************************************

#if !defined(__GNUC__) && !defined(__TINYC__)
   #if !(defined(__BCPLUSPLUS__) && defined(_WIN64))
      #if !defined(__LCC__)
    #pragma comment(lib,"kernel32.lib")
    #pragma comment(lib,"user32.lib")
    #pragma comment(lib,"gdi32.lib")
    #pragma comment(lib,"comctl32.lib")
    #pragma comment(lib,"advapi32.lib")
    #pragma comment(lib,"winspool.lib")
    #pragma comment(lib,"shell32.lib")
    #pragma comment(lib,"msimg32.lib")
    #pragma comment(lib,"ole32.lib")
    #pragma comment(lib,"oleaut32.lib")
    #pragma comment(lib,"uuid.lib")
    #pragma comment(lib,"odbc32.lib")
    #pragma comment(lib,"odbccp32.lib")
    #pragma comment(lib,"winmm.lib")
    #pragma comment(lib,"comdlg32.lib")
    #pragma comment(lib,"imagehlp.lib")
    #pragma comment(lib,"version.lib")
    #pragma comment(lib,"wininet.lib")
    #pragma comment(lib,"urlmon.lib")
  #else
    #pragma lib <winspool.lib>
    #pragma lib <shell32.lib>
    #pragma lib <msimg32.lib>
    #pragma lib <ole32.lib>
    #pragma lib <oleaut32.lib>
    #pragma lib <uuid.lib>
    #pragma lib <odbc32.lib>
    #pragma lib <odbccp32.lib>
    #pragma lib <winmm.lib>
    #pragma lib <imagehlp.lib>
    #pragma lib <version.lib>
    #pragma lib <wininet.lib>
    #pragma lib <urlmon.lib>
     #endif
  #endif
#endif

// *************************************************
//               Standard Prototypes 
// *************************************************

char*   BCX_TempStr (size_t);
char*   mid (LPCTSTR, int, int=-1);
char*   left (LPCTSTR, int);
char*   replace (LPCTSTR, LPCTSTR, LPCTSTR);
char*   str (double, int=0);
char*   RemoveStr (LPCTSTR, LPCTSTR);
char*   join(int, LPCTSTR, ...);
char*   _stristr_(LPCTSTR, LPCTSTR);
int     instr (LPCTSTR, LPCTSTR, int=1, int=0);
void    Pause (void);
int     scan (LPCTSTR input, LPCTSTR format, ... );
void    KBinput (void);
int     Split (char[][BCXSTRSIZE], LPCTSTR, LPCTSTR, int = 0);
unsigned char*  MakeUCaseTbl (void);

// *************************************************
//          Late binding COM support section
//  (c) Ljubisa Knezevic 2004-2009, ljube@blic.net
// *************************************************

#define COM_MSGBOX(a,b,c)MessageBox(GetActiveWindow(),(a),(b),(c))
#define COM_PROPS DISPATCH_PROPERTYGET|DISPATCH_METHOD
#define COM_PARAMS COM_param_list.pParams

//<---UNICODE AWARE
char*   WideToAnsi (BSTR, UINT = CP_ACP, ULONG = 0);
//>---UNICODE AWARE

//<---UNICODE AWARE
LPOLESTR AnsiToWide (LPCSTR, UINT = CP_ACP, ULONG = MB_PRECOMPOSED);
//>---UNICODE AWARE

HRESULT  COM_AS2WS (LPCSTR  ansi_string, UINT code_page = CP_ACP);
HRESULT  COM_WS2AS (LPCWSTR wide_string, UINT code_page = CP_ACP);

#define COM_STACK_SIZE  64
#ifndef CON_VARBOOL2BOOL
  #define CON_VARBOOL2BOOL(b) ((BOOL)(b ? TRUE : FALSE))
#endif

typedef struct _OBJECT {
  IUnknown*  p_unknown;
  VARIANT    pObjects[COM_STACK_SIZE];
  BOOL       pStatus;
  int        ipointer;
}OBJECT, *LPOBJECT;

typedef struct _PARAM_VARARRAY {
  VARIANT  pParams[COM_STACK_SIZE];
}PARAM_VARARRAY, *LPPARAM_VARARRAY;

// *************************************************
//       COM functions used internally by BCX
// *************************************************

void     COM_ole_initialize            (void);
void     COM_ole_uninitialize          (void);
void     COM_HR_ErrMsg                 (HRESULT hr, TCHAR *extra_info);
DISPID   COM_get_dispatch_ID           (IDispatch *lpDispatch, LPOLESTR comsegment);
void     COM_get_next_dispatch         (OBJECT *object, LPOLESTR comsegment);
void     COM_invoke                    (OBJECT *object, LPOLESTR comsegment, WORD wFlags, VARIANT *pvResult);
void     COM_build_except_info         (HRESULT hr, EXCEPINFO *pexcep = NULL, UINT uiArgErr = 0);
void     COM_clean_plist               (void);
void     COM_reset_disp_chain          (OBJECT *object);
void     COM_create_safearray          (void);
void     COM_trace_dump_DISPPARAMS     (DISPPARAMS* dp);
void     COM_trace_add_line            (LPCTSTR dp);
void     COM_trace_dump_indicators     (OBJECT *object);
void     COM_trace_dump_flags          (WORD wFlags);
void     COM_FREE_TEMP_ANSI_STRING     (void);
void     COM_FREE_TEMP_WIDE_STRING     (void);

// *************************************************
//       Public COM support functions
// *************************************************

OBJECT   COM                    (LPCTSTR);
void     UNCOM                  (OBJECT);
void     BCX_SHOW_COM_ERRORS    (BOOL);
void     BCX_SetNothing         (OBJECT*);
BOOL     BCX_GET_COM_SUCCESS    (void);
BOOL     BCX_GET_COM_STATUS     (OBJECT*);
char*    BCX_GET_COM_ERROR_DESC (void);
int      ISOBJECT               (OBJECT);
HRESULT  BCX_GET_COM_ERROR_CODE (void);
void     BCX_CreateObject       (TCHAR *objname, OBJECT *obj);
void     BCX_GetObject          (TCHAR *objname, OBJECT *obj);
void     BCX_GetObjectMon       (LPCOLESTR objname, OBJECT *obj);
void     BCX_DispatchObject     (IUnknown *iobj, OBJECT *obj, BOOL b_release = TRUE);

static int     ScanError;          // holds last error from scan function
static char    InputBuffer[1048577];
static unsigned char*  UprCase;

// *************************************************
//    Global vars used by late binding COM support
// *************************************************

static PARAM_VARARRAY  COM_param_list;
static _TCHAR          COM_last_ErrMsg[4096];
static _TCHAR          COM_ErrMsg[64];
static int             COM_plist_index = 0;
static VARIANT         COM_vt_result;
static HRESULT         COM_last_HR;
static BOOL            COM_ole_initd = FALSE;
static int             COM_objects_cnt = 0;
static int             COM_reset_chain = 0;
static BOOL            COM_BCX_ERROR = FALSE;
static BOOL            COM_bSHOW_ERROR = FALSE;
static SAFEARRAY *     COM_PTR_safearray = NULL;
static LPOLESTR        COM_LPWSTR_temp = NULL;
static char*           COM_psz_tmp = NULL;
static ULONG           COM_wstr_size = 0;
static ULONG           COM_zstr_size = 0;
static BOOL            COM_GetEnum_iface = FALSE;
static int             COM_dispatch_storage[ 32];
static int             COM_dispatch_storage_index = 0;
static int             COM_dispatch_at_offset = 0;
static VARIANT         COM_ack_var;
static IEnumVARIANT*   COM_enum_var = NULL;
static ULONG           COM_long_coll = 0;

// *************************************************
//            User's Global Variables 
// *************************************************

static OBJECT  Session;
static OBJECT  Searcher;
static OBJECT  SearchResult;
static OBJECT  Update;
static OBJECT  Updates;
static int     NumUpdates;
static int     I;
static int     NumSelections;
static int     Index=1;
static char    Input[BCXSTRSIZE];
static char    ToHide[100][BCXSTRSIZE];

// *************************************************
//               Standard Macros 
// *************************************************

#define VAL(a)(double)atof((a))

char *BCX_TempStr(size_t iBytes)
{
  static int StrCnt;
  static char *StrFunc[2048];
  static CRITICAL_SECTION cs;
  static int initialized = 0;
  // Initialize the critical section once
  if (!initialized) {
    InitializeCriticalSection(&cs);
    initialized = 1;
  }
  EnterCriticalSection(&cs);
  StrCnt = (StrCnt + 1) & 2047;
  if (StrFunc[StrCnt]) {
    free(StrFunc[StrCnt]);
    StrFunc[StrCnt] = NULL;
  }
  StrFunc[StrCnt] = (char*)calloc(iBytes + 1, sizeof(char));
  if (!StrFunc[StrCnt]) {
    LeaveCriticalSection(&cs);
    return NULL; // Allocation failed
  }
  LeaveCriticalSection(&cs);
  return StrFunc[StrCnt];
}

char *left (LPCTSTR szSrc, int length)
{
  size_t tmplen = strlen(szSrc);
  if (length < 1) return BCX_TempStr(1);
  if (length < (int)tmplen) tmplen = length;
  char *strtmp = BCX_TempStr(tmplen);
  return (char*)memcpy(strtmp, szSrc, tmplen);
}

char *mid (LPCTSTR szSrc, int start, int length)
{
  char *strtmp;
  int tmplen = (int)strlen(szSrc);
  if (start > tmplen||start < 1) return BCX_TempStr(1);
  if (length < 0 || length > (tmplen-start) + 1)
    length = (tmplen-start) + 1;
  strtmp = BCX_TempStr(length);
  return (char*)memcpy(strtmp, &szSrc[start - 1], length);
}

char *replace (LPCTSTR szMain, LPCTSTR szFind, LPCTSTR szRepl)
{
  size_t LenPat, LenRep, iTmp, iChange;
  char *strtmp, *cp_p, *cp_q, *cp_r;
  if (!szFind || !*szFind) return (char*) szMain;
  LenRep = strlen(szRepl);
  LenPat = strlen(szFind);
  for (iTmp = 0, cp_p = (char*)szMain;
       (cp_q = strstr(cp_p, (char*)szFind))!=0;
       cp_p = cp_q + LenPat)
    iTmp += (int)(cp_q - cp_p) + LenRep;
  iTmp += strlen(cp_p);
  strtmp = BCX_TempStr(iTmp);
  if (!strtmp) return NULL;
  for (cp_r = strtmp, cp_p=(char*)szMain;
       (cp_q=strstr(cp_p,(char*)szFind))!=0;
       cp_p = cp_q + LenPat)
  {
    iChange = (int)(cp_q-cp_p);
    memcpy(cp_r, cp_p, iChange); cp_r += iChange;
    strcpy(cp_r, szRepl); cp_r += LenRep;
  }
  strcpy(cp_r, cp_p);
  return strtmp;
}

char *RemoveStr (LPCTSTR szSrc, LPCTSTR RemoveMe)
{
  char *strtmp, *pp, *dd;
  int  tmplen;
  strtmp = dd = BCX_TempStr(strlen(szSrc));
  if ((!RemoveMe) || (!*RemoveMe)) return strcpy(strtmp, szSrc);
  pp = strstr((char*)szSrc, (char*)RemoveMe); tmplen = (int)strlen(RemoveMe);
  while (pp)
  {
    memcpy (dd, szSrc, pp-szSrc);
    dd+= (pp - szSrc);
    szSrc = pp + tmplen;
    pp = strstr((char*)szSrc,(char*)RemoveMe);
  }
  strcpy (dd, (char*)szSrc);
  return strtmp;
}

char *str(double d_, int nospc)
{
  char *strtmp = BCX_TempStr(64);
  if (nospc)
    snprintf(strtmp, 64, "%.15G", d_);
  else
    snprintf(strtmp, 64, "% .15G", d_);
  return strtmp;
}

char *join(int ArgCount, LPCTSTR Str1, ...)
{
  va_list  ap={0};
  PSTR BCX_RetStr = NULL;
  PSTR currentStr = NULL;
  int i;
  size_t totalLen;
  totalLen = strlen(Str1);
                                      //  Calc memory reqmnt
  va_start(ap,Str1);
  for(i=1; i<=ArgCount-1; i++)
    {
      currentStr = va_arg(ap,PSTR);
      if(strlen(currentStr)>0 ){
          totalLen += (int)strlen(currentStr);
        }
    }
  va_end(ap);
                                      //  Allocate memory
  BCX_RetStr = BCX_TempStr(totalLen);
  if(BCX_RetStr == 0 ){
      return NULL;                    //   Not a good sign :-(
    }
                                      //   Init result with Str1
  strcpy(BCX_RetStr,Str1);
                                      //   Concat remainiing strings
  va_start(ap,Str1);
  for(i=1; i<=ArgCount-1; i++)
    {
      currentStr = va_arg(ap,PSTR);
      if((int)strlen(currentStr)>0 ){
          strcat(BCX_RetStr,currentStr);
        }
    }
  va_end(ap);
  return BCX_RetStr;
}

char *_stristr_(LPCTSTR Haystack, LPCTSTR Needle)
{
  int mi=(-1);
  while(Needle[++mi])
  {
    if (Haystack[mi]==0) return 0;
    if (UprCase[(unsigned char)Haystack[mi]]!=UprCase[(unsigned char)Needle[mi]])
    {
      Haystack++;
      mi=(-1);
    }
  }
  return (char*)Haystack;
}

//<---UNICODE AWARE
char* WideToAnsi (BSTR WideStr, UINT CodePage, ULONG dwFlags)
{
  char *BCX_RetStr;
  UINT uLen;
  uLen = WideCharToMultiByte(CodePage, dwFlags, WideStr,-1,0,0,0,0);
  BCX_RetStr = (char*)BCX_TempStr(uLen);
  WideCharToMultiByte(CodePage, dwFlags, WideStr, -1, BCX_RetStr,uLen,0,0);
  return BCX_RetStr;
}
//>---UNICODE AWARE

void Pause(void)
{
  printf("\n%s\n","Press any key to continue . . .");
  _getch();
}

int instr(LPCTSTR Haystack, LPCTSTR Needle, int offset, int sensflag)
{
  if(offset == 0) {return 0;}
  LPCTSTR s;
  if (!Haystack||!*Haystack||!Needle||!*Needle||offset>(int)strlen(Haystack)) return 0;
  if (sensflag)
    s = _stristr_(offset > 0 ? Haystack + offset - 1 : Haystack, Needle);
  else
    s = strstr(offset > 0 ? Haystack + offset - 1 : Haystack, Needle);
  return s ? (int)(s-Haystack)+1 : 0;
}

unsigned char* MakeUCaseTbl (void)
{
  static unsigned char tbl[257];
  int ii; for (ii = 0; ii < 256; ii++)
    tbl[ii] = (unsigned char) ii;
  tbl[0] = 1;
  tbl[256] = 0;
  CharUpperA((char *)tbl);
  tbl[0] = 0;
  return tbl;
}

int scan(LPCTSTR input, LPCTSTR format, ... )
{
  int       L_c ;
  int       L_d ;
  char     *s_;
  int      *intptr;
  float    *floatptr;
  double   *doubleptr;
  char      L_ScanBuffer[50][BCXSTRSIZE];
  va_list   marker;
  L_c = 0;
  L_d = Split(L_ScanBuffer,input,",");
  va_start(marker, format); //Initialize arguments
  while(L_d && *format)
  {
    if (*format == '%') format++;
    if (*format == 's')
    {
      s_ = va_arg(marker, char *);
      strcpy(s_, L_ScanBuffer[L_c]);
      L_c++;
      L_d--;
    }
    if (*format == 'd')
    {
      intptr = va_arg(marker, int *);
      *intptr = atoi(L_ScanBuffer[L_c]);
      L_c++;
      L_d--;
    }
    if (*format == 'g')
    {
      floatptr = va_arg(marker, float *);
      *floatptr = strtof(L_ScanBuffer[L_c], NULL);
      L_c++;
      L_d--;
    }
    if (*format == 'l')
    {
      format++;
      doubleptr = va_arg(marker, double *);
      *doubleptr = atof(L_ScanBuffer[L_c]);
      L_c++;
      L_d--;
    }
    format++;
  }
  va_end(marker);              // Reset variable arguments
  if (L_d) return(1);          // More data than variables
  if (0 == *format) return(0); // OK
  return(-1);                  // More variables than data
}

void KBinput (void)
{
  *InputBuffer = 0;
  fgets(InputBuffer, 0x100000, stdin);
  InputBuffer[(int)strlen(InputBuffer)-1]=0;
}

int Split (char Buf[][BCXSTRSIZE], LPCTSTR T, LPCTSTR Delim, int Flg)
{
  int  Begin = 0;
  int  Count = 0;
  int  Quote = 0;
  int  Index = 0;
  int  ii    = 0;
  int  lenT  = (int)strlen(T);
  if (Flg<0||Flg>4) Flg=0;
  const char DDQt [3] = {34,34,0}; // Double Quotation Mark
  const char US   [2] = {30,0};    // Information Separator One
  const char DQt  [2] = {34,0};    // Quotation Mark
  for(Index = 1;Index<=lenT;Index++)
  {
    if (instr(Delim, mid(T, Index, 1)) && !Quote)
    {
      strcpy(Buf[Count], (char*)mid(T, Begin, Index-Begin));
      if (0 == (Flg & 2))
        Count++;
      else
      if (Buf[Count][0] != 0) Count++;
      Begin = 0;
      if ((Flg & 1) == 1)   // 1 if true
        strcpy(Buf[Count++], (char*)mid(T, Index, 1));
    }
    else
    {
      if (strcmp(mid(T, Index, 1), DQt)==0) Quote=!Quote;
      if (Begin==0) Begin = Index;
    }
  }
  if (Begin)
    strcpy(Buf[Count++], (char*)mid(T, Begin, Index-Begin));
  if ((0 == (Flg & 1)) && (Flg < 4)) // 0 if false
    for(ii = 0;ii<Count;ii++) strcpy(Buf[ii], (char*)RemoveStr(Buf[ii], DQt));
  if(Flg == 4)
  {
    for(Index = 0;Index<Count;Index++)
    {
      if(instr(Buf[Index],DDQt))
      {
        strcpy(Buf[Index],replace(Buf[Index],DDQt,US));
        strcpy(Buf[Index],RemoveStr(Buf[Index],DQt));
        strcpy(Buf[Index],replace(Buf[Index],US,DQt));
      }
      else
      {
        strcpy(Buf[Index],RemoveStr(Buf[Index],DQt));
      }
    }
  }
  return Count;
}

//<---UNICODE AWARE
LPOLESTR AnsiToWide (LPCSTR AnsiStr, UINT CodePage, ULONG dwFlags)
{
  UINT uLen;
  BSTR WideStr;
  uLen = MultiByteToWideChar(CodePage,dwFlags,AnsiStr,-1,0,0);
  if (uLen<=1) return (BSTR) BCX_TempStr(2);
  WideStr = (BSTR) BCX_TempStr(2*uLen);
  MultiByteToWideChar(CodePage,dwFlags,AnsiStr,uLen,WideStr,uLen);
  return WideStr;
}
//>---UNICODE AWARE

OBJECT COM (LPCTSTR szProgID_) {
    static OBJECT oProxy;
    BCX_CreateObject ((char*)szProgID_, &oProxy);
    return oProxy;
}

void UNCOM (OBJECT oObj_) {
    static OBJECT oProxy;
    oProxy = oObj_;
    BCX_SetNothing (&oProxy);
}

int ISOBJECT (OBJECT Obj) {
    if(Obj.pStatus)
        return TRUE;
   else
       return FALSE;
}

HRESULT BCX_GET_COM_ERROR_CODE (void) {
    return COM_last_HR;
}

_TCHAR* BCX_GET_COM_ERROR_DESC (void) {
    return COM_last_ErrMsg;
}

BOOL BCX_GET_COM_SUCCESS (void) {
    return (!COM_BCX_ERROR);
}

void BCX_SHOW_COM_ERRORS (BOOL Show_err) {
    COM_bSHOW_ERROR = Show_err;
}

BOOL BCX_GET_COM_STATUS(OBJECT *object) {
    if(object)
        return object->pStatus;
    return 0;
}

void  COM_clean_plist (void) {
int total_parms = COM_plist_index;
if (COM_plist_index > 0)
   {
      do
   {
        COM_last_HR = VariantClear (&COM_PARAMS[COM_plist_index-1]);
        if (FAILED(COM_last_HR))
        {
        wsprintf (COM_ErrMsg, _T("\nVariant type = %d, at index %d, total params = %d."),
                  COM_PARAMS[COM_plist_index-1].vt, COM_plist_index-1, total_parms);
        COM_HR_ErrMsg (COM_last_HR, _T("Error while cleaning parameter list."));
      }
        COM_plist_index--;
    }
      while (COM_plist_index > 0);
  }
}

void  COM_reset_disp_chain (OBJECT *object) {
  if (object->ipointer>COM_dispatch_at_offset)
  {
    do
    {
      VariantClear (&object->pObjects[object->ipointer]);
      object->ipointer--;
    }
    while (object->ipointer > COM_dispatch_at_offset);
  }
}

void  COM_ole_uninitialize (void) {
  if (COM_objects_cnt)
  {
    _TCHAR ermm[BCXSTRSIZE];
    wsprintf (ermm,_T("Check SET Nothing statement!\nUnreleased objects: %d"),COM_objects_cnt);
    COM_MSGBOX (ermm, _T("Memory leaks detected!"),MB_OK|MB_ICONWARNING|MB_TASKMODAL);
  }
  COM_FREE_TEMP_WIDE_STRING();
  COM_FREE_TEMP_ANSI_STRING();
  CoUninitialize();
}

void  COM_get_next_dispatch (OBJECT *object, LPOLESTR comsegment) {
  if (!object->pStatus) return;
  if (0 == object->ipointer) COM_BCX_ERROR = FALSE;
  if (COM_BCX_ERROR) return;
  COM_reset_chain++;
  COM_invoke (object, comsegment,COM_PROPS,&object->pObjects[object->ipointer+1]);
  COM_reset_chain--;
  if (!COM_BCX_ERROR)
     object->ipointer++;
}

void COM_HR_ErrMsg (HRESULT hr, _TCHAR *extra_info) {
  COM_BCX_ERROR = TRUE;
  PVOID pMsgBuf;

  FormatMessage (FORMAT_MESSAGE_ALLOCATE_BUFFER|FORMAT_MESSAGE_FROM_SYSTEM,
  NULL, hr, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPTSTR)&pMsgBuf, 0, NULL);

  wsprintf(COM_last_ErrMsg, _T("COM error code %d (0x%X)\n%s\n%s\nmember: %s"),
  hr, hr, extra_info, pMsgBuf, COM_ErrMsg);

  memset(&COM_ErrMsg,0,sizeof(COM_ErrMsg));
  if (COM_bSHOW_ERROR)
      COM_MSGBOX(COM_last_ErrMsg,_T("COM Parser Error Report:"),
      MB_OK|MB_ICONERROR|MB_SYSTEMMODAL);
  LocalFree(pMsgBuf);
}

void BCX_SetNothing (OBJECT *obj) {
  if (!obj->pStatus) return;
  COM_objects_cnt--;
  #ifdef __cplusplus
    if (obj->p_unknown) obj->p_unknown->Release();
  #else
    if (obj->p_unknown) obj->p_unknown->lpVtbl->Release(obj->p_unknown);
  #endif
  if (obj->ipointer)
  {
    COM_reset_disp_chain (obj);
    obj->ipointer = 0;
  }
  COM_last_HR = VariantClear (&obj->pObjects[0]);
  if (FAILED(COM_last_HR))
  {
    lstrcpy(COM_ErrMsg, _T("BCX_SetNothing Failed!"));
    COM_HR_ErrMsg (COM_last_HR, _T("Release of IDispatch objs failed!"));
  }
  obj->pStatus = FALSE;
  Sleep(100);
}

void COM_ole_initialize (void) {
  if (COM_ole_initd) return;
  #ifdef __BCX_MULTITHREADED__
    COM_last_HR = CoInitializeEx (NULL, COINIT_MULTITHREADED);
  #else
    COM_last_HR = CoInitializeEx (NULL, COINIT_APARTMENTTHREADED);
  #endif
  if (SUCCEEDED (COM_last_HR))
  {
    COM_ole_initd = TRUE;
    VariantInit (&COM_vt_result);
    atexit (COM_ole_uninitialize);

    COM_dispatch_storage[0] = 0;
    COM_dispatch_storage_index = 0;

    if (COM_LPWSTR_temp) CoTaskMemFree((void*)COM_LPWSTR_temp);
    COM_LPWSTR_temp = (LPOLESTR)CoTaskMemAlloc (BCXSTRSIZE);
    if (NULL == COM_LPWSTR_temp)
    {
      COM_wstr_size = 0;
      COM_last_HR =  E_OUTOFMEMORY;
      COM_HR_ErrMsg (COM_last_HR,_T(
     "CoInitializeEx: Memory allocation failure!"));
      return;
    }
    COM_wstr_size = BCXSTRSIZE;
    if (COM_psz_tmp) free(COM_psz_tmp);
    COM_psz_tmp = (char*)calloc(BCXSTRSIZE, sizeof(char));
    if (NULL == COM_psz_tmp)
    {
      COM_last_HR =  E_OUTOFMEMORY;
      COM_HR_ErrMsg (COM_last_HR,_T(
     "CoInitializeEx: Memory allocation failure!"));
    return;
    }
    COM_zstr_size = BCXSTRSIZE;
    return;
  }
  else
    COM_HR_ErrMsg (COM_last_HR,_T("CoInitializeEx Failed!"));
}

void COM_build_except_info (HRESULT hr, EXCEPINFO *pexcep, UINT uiArgErr) {
  SCODE oleSCODE;
  static _TCHAR lv_message[BCXSTRSIZE];
  oleSCODE = GetScode(hr);
  for (;;)
  {
    if (oleSCODE==DISP_E_PARAMNOTFOUND)
    {
      wsprintf (lv_message, _T("\nArgument not found, argument %d."), uiArgErr);
      _tcscat(COM_last_ErrMsg, lv_message);
      break;
    }
    if (oleSCODE==DISP_E_TYPEMISMATCH)
    {
      wsprintf (lv_message, _T("\nType mismatch, argument %d."), uiArgErr);
      _tcscat (COM_last_ErrMsg, lv_message);
      break;
    }
    if (oleSCODE==DISP_E_BADVARTYPE)
    {
      _tcscat(COM_last_ErrMsg, _T("\nOne or more invalid VARIANT arguments."));
      break;
    }
    if (oleSCODE==E_INVALIDARG)
    {
      _tcscat(COM_last_ErrMsg, _T("\nOne of the arguments is invalid."));
      break;
    }
    break;
  }
  if (pexcep)
  {
    wsprintf(lv_message, _T("\nCOM Error %X:"), pexcep->wCode);
    if (pexcep->bstrDescription)
    {
      _TCHAR err_desc[512];
      #ifndef UNICODE
        COM_last_HR = COM_WS2AS (pexcep->bstrDescription);
        if (FAILED(COM_last_HR))
          COM_HR_ErrMsg (COM_last_HR,_T("Get Error Desc: W2A failure!"));
        else
          wsprintf(err_desc, _T("\nException desc: %s"), COM_psz_tmp);
      #else
          wsprintf(err_desc, _T("\nException desc: %s"), pexcep->bstrDescription);
      #endif
      _tcscat(lv_message, err_desc);
      SysFreeString(pexcep->bstrDescription);
      pexcep->bstrDescription = NULL;
    }
    if (pexcep->bstrSource)
    {
      _TCHAR err_desc[512];
      #ifndef UNICODE
        COM_last_HR = COM_WS2AS (pexcep->bstrSource);
        if (FAILED(COM_last_HR))
            COM_HR_ErrMsg (COM_last_HR,_T("Get Error Source Failed! W2A failure!"));
        else
            wsprintf(err_desc, _T("\nException source: %s"), COM_psz_tmp);
      #else
            wsprintf(err_desc, _T("\nException source: %s"), pexcep->bstrSource);
      #endif
      _tcscat (lv_message, err_desc);
      SysFreeString (pexcep->bstrSource);
      pexcep->bstrSource = NULL;
    }
    if (pexcep->bstrHelpFile)
    {
      _TCHAR err_desc[512];
      #ifndef UNICODE
        COM_last_HR = COM_WS2AS (pexcep->bstrHelpFile);
        if (FAILED(COM_last_HR))
        COM_HR_ErrMsg (COM_last_HR,_T("Help File failed: W2A failure!"));
        else
        wsprintf(err_desc, _T(
        "\nHelp file: %s | topic: %lu"), COM_psz_tmp, pexcep->dwHelpContext);
      #else
        wsprintf(err_desc, _T(
        "\nHelp file: %s | topic: %lu"), pexcep->bstrHelpFile, pexcep->dwHelpContext);
      #endif
      _tcscat(lv_message, err_desc);
      SysFreeString (pexcep->bstrHelpFile);
      pexcep->bstrHelpFile = NULL;
    }
    _tcscat (COM_last_ErrMsg ,lv_message);
  }
  if (COM_bSHOW_ERROR)
  {
    COM_MSGBOX(COM_last_ErrMsg, _T("COM Parser Exception Info:"),
    MB_OK|MB_ICONERROR|MB_SYSTEMMODAL);
  }
}

DISPID COM_get_dispatch_ID (IDispatch* lpDispatch, LPOLESTR comsegment) {
  static DISPID D_ID;
  D_ID = 0;
  if (!lpDispatch) return -1;
  #ifdef __cplusplus
    COM_last_HR = lpDispatch->GetIDsOfNames
                 (IID_NULL, &comsegment, 1, LOCALE_SYSTEM_DEFAULT, &D_ID);
  #else
    COM_last_HR = lpDispatch->lpVtbl->GetIDsOfNames
                 (lpDispatch, &IID_NULL, &comsegment, 1, LOCALE_SYSTEM_DEFAULT, &D_ID);
  #endif
  if (FAILED(COM_last_HR))
  {
    COM_HR_ErrMsg (COM_last_HR,_T("Unrecognized member name"));
    return -1;
  }
  return D_ID;
}

void COM_invoke (OBJECT *object, LPOLESTR comsegment, WORD wFlags, VARIANT *pvResult) {
  if (!object->pStatus) return;
  if (COM_BCX_ERROR) return;
  DISPID lv_DID;
  EXCEPINFO exception;
  UINT argerr = 0;
  DISPPARAMS dp = {NULL, NULL, 0, 0 };
  WORD invoke_flags = 0;
  DISPID setdispid = DISPID_PROPERTYPUT;
  if (COM_GetEnum_iface)
    lv_DID = DISPID_NEWENUM;
  else
  {
    lv_DID = COM_get_dispatch_ID (object->pObjects[object->ipointer].pdispVal, comsegment);
    if (-1 == lv_DID) goto cleanInProp;
  }
  memset (&exception, 0, sizeof(EXCEPINFO));
  if (COM_plist_index>0)
  {
    setdispid = DISPID_PROPERTYPUT;
    dp.rgvarg = (VARIANTARG*)COM_PARAMS;
  }
  if (wFlags & DISPATCH_PROPERTYPUT)
  {
    dp.rgdispidNamedArgs = &setdispid;
    dp.cNamedArgs = 1;
  }
  dp.cArgs = COM_plist_index;
  invoke_flags = wFlags;
  if (wFlags & DISPATCH_PROPERTYGET)
    if (pvResult) VariantInit(pvResult);
  if (VT_DISPATCH!=object->pObjects[object->ipointer].vt||NULL==object->pObjects[object->ipointer].pdispVal)
  {
     COM_last_HR = E_NOINTERFACE;
     COM_HR_ErrMsg (COM_last_HR,_T("COM_invoke::Invalid IDispatch interface."));
     goto cleanInProp;
  }
    #ifdef __cplusplus
    COM_last_HR = object->pObjects[object->ipointer].pdispVal
    ->Invoke (lv_DID, IID_NULL, LOCALE_SYSTEM_DEFAULT, 
        invoke_flags, &dp, pvResult, &exception, &argerr);
    #else
    COM_last_HR = object->pObjects[object->ipointer].pdispVal->lpVtbl
    ->Invoke(object->pObjects[object->ipointer].pdispVal, lv_DID, &IID_NULL, 
        LOCALE_SYSTEM_DEFAULT, invoke_flags , &dp, pvResult, &exception, &argerr);
    #endif
  if (FAILED(COM_last_HR))
  {
    COM_HR_ErrMsg (COM_last_HR,_T("COM_invoke::Invoke failed."));
    COM_build_except_info (COM_last_HR, &exception, argerr);
  }
cleanInProp:
  if (V_ISARRAY(&COM_PARAMS[0]))
  {
    if ((VT_ARRAY|VT_VARIANT)==COM_PARAMS[0].vt)
    {
      COM_last_HR = SafeArrayDestroy(V_ARRAY(&COM_PARAMS[0]));
      if (FAILED (COM_last_HR))
      {
        lstrcpy(COM_ErrMsg, _T("SafeArrayDestroy failed"));
        COM_HR_ErrMsg (COM_last_HR,_T("Error wiping param list."));
      }
      ZeroMemory((PVOID)&COM_PARAMS[0],sizeof(VARIANT));
      COM_PTR_safearray = NULL;
      COM_plist_index = 0;
    }
  }
  if (COM_plist_index)  COM_clean_plist();
  if (0 == COM_reset_chain)  COM_reset_disp_chain(object);
}

void BCX_CreateObject (_TCHAR *objname, OBJECT *obj) {
  if (!COM_ole_initd) COM_ole_initialize();
  CLSID clsid;
  #ifndef UNICODE
    COM_last_HR = COM_AS2WS(objname);
    if (FAILED(COM_last_HR))
    {
      COM_HR_ErrMsg (COM_last_HR,_T("CreateObject Failed! A2W failure!"));
      return;
    }
    COM_last_HR = COM_WS2AS(COM_LPWSTR_temp);
    if (FAILED (COM_last_HR))
    {
       COM_HR_ErrMsg (COM_last_HR,_T("CreateObject Failed! W2A failure!"));
      return;
    }
    sprintf(COM_ErrMsg,"%s, WideName(%s)", objname, COM_psz_tmp);
    COM_last_HR = CLSIDFromProgID ((LPCOLESTR)COM_LPWSTR_temp, (LPCLSID)&clsid);
  #else
    lstrcpy(COM_ErrMsg, objname);
    COM_last_HR = CLSIDFromProgID ((LPCOLESTR)objname, (LPCLSID)&clsid);
  #endif
  if (FAILED (COM_last_HR))
  {
    COM_HR_ErrMsg (COM_last_HR,_T("CLSIDFromProgID failed!"));
    return;
  }
  #ifdef __cplusplus
    COM_last_HR = CoCreateInstance((REFCLSID)clsid, NULL,  
    CLSCTX_INPROC_SERVER|CLSCTX_LOCAL_SERVER, IID_IUnknown, (void **)&obj->p_unknown);
  #else
    COM_last_HR = CoCreateInstance ((REFCLSID)&clsid, NULL, 
    CLSCTX_INPROC_SERVER|CLSCTX_LOCAL_SERVER,&IID_IUnknown, (void **)&obj->p_unknown);
  #endif
  if (FAILED (COM_last_HR))
     COM_HR_ErrMsg (COM_last_HR,_T("CoCreateInstance failed!"));
  else
  {
    VariantInit(&obj->pObjects[0]);
    obj->pObjects[0].vt = VT_DISPATCH;
  #ifdef __cplusplus
    COM_last_HR = obj->p_unknown
  ->QueryInterface(IID_IDispatch, (void **)&obj->pObjects[0].pdispVal);
  #else
    COM_last_HR = obj->p_unknown->lpVtbl
  ->QueryInterface(obj->p_unknown, &IID_IDispatch, (void**)&obj->pObjects[0].pdispVal);
  #endif
  }
  if (FAILED(COM_last_HR))
  {
    COM_HR_ErrMsg (COM_last_HR,_T("QueryInterface::IID_IDispatch  failed!"));
    obj->pObjects[0].vt = VT_NULL;
    VariantClear(&obj->pObjects[0]);
    #ifdef __cplusplus
       obj->p_unknown->Release();
    #else
       obj->p_unknown->lpVtbl->Release(obj->p_unknown);
    #endif
  }
  else
  {
    obj->pStatus = TRUE;
    obj->ipointer = 0;
    COM_objects_cnt++;
  }
}

void BCX_GetObject (_TCHAR *objname, OBJECT *obj) {
  if (!COM_ole_initd) COM_ole_initialize();
  CLSID clsid;
  #ifndef UNICODE
    COM_last_HR = COM_AS2WS (objname);
    if (FAILED(COM_last_HR))
    {
      COM_HR_ErrMsg (COM_last_HR,_T("CreateObject Failed! A2W failure!"));
      return;
    }
    COM_last_HR = COM_WS2AS (COM_LPWSTR_temp);
    if (FAILED (COM_last_HR))
    {
      COM_HR_ErrMsg (COM_last_HR,_T("CreateObject Failed! W2A failure!"));
      return;
    }
    sprintf (COM_ErrMsg,"%s, WideName(%s)", objname, COM_psz_tmp);
    COM_last_HR = CLSIDFromProgID((LPCOLESTR)COM_LPWSTR_temp, (LPCLSID)&clsid);
  #else
    COM_last_HR = COM_WS2AS ((LPCWSTR)objname);
    lstrcpy (COM_ErrMsg, COM_psz_tmp);
    COM_last_HR = CLSIDFromProgID ((LPCOLESTR)objname, (LPCLSID)&clsid);
  #endif
  if (FAILED (COM_last_HR))
  {
  #ifndef UNICODE
    BCX_GetObjectMon ((LPCOLESTR)COM_LPWSTR_temp, obj);
  #else
    BCX_GetObjectMon ((LPCOLESTR)objname, obj);
  #endif
    return;
  }
  #ifdef __cplusplus
    COM_last_HR = GetActiveObject((REFCLSID)clsid, NULL, (IUnknown**)&obj->p_unknown);
  #else
    COM_last_HR = GetActiveObject((REFCLSID)&clsid, NULL, (IUnknown**)&obj->p_unknown);
  #endif
  if (FAILED (COM_last_HR))
  {
    COM_HR_ErrMsg (COM_last_HR,_T("GetActiveObject failed!"));
    return;
  }
  else
  {
    VariantInit (&obj->pObjects[0]);
    #ifdef __cplusplus
      COM_last_HR = obj->p_unknown
    ->QueryInterface(IID_IDispatch, (void**)&obj->pObjects[0].pdispVal);
    #else
      COM_last_HR = obj->p_unknown->lpVtbl
    ->QueryInterface(obj->p_unknown, &IID_IDispatch, (void **)&obj->pObjects[0].pdispVal);
    #endif
  }
  if (FAILED (COM_last_HR))
  {
    COM_HR_ErrMsg (COM_last_HR,_T("QueryInterface::IID_IDispatch  failed!"));
    if (obj->p_unknown)
    {
    #ifdef __cplusplus
      obj->p_unknown->Release();
    #else
      obj->p_unknown->lpVtbl->Release(obj->p_unknown);
    #endif
    }
  }
  else
  {
    obj->pObjects[0].vt = VT_DISPATCH;
    obj->pStatus = TRUE;
    obj->ipointer = 0;
    COM_objects_cnt++;
  }
}

void BCX_GetObjectMon (LPCOLESTR objname, OBJECT *obj) {
  IBindCtx* vBindCtx = NULL;
  IMoniker* vMoniker = NULL;
  ULONG vChEaten = 0;
  COM_last_HR = CreateBindCtx(0, &vBindCtx);
  if (COM_last_HR != S_OK)
  {
    COM_HR_ErrMsg (COM_last_HR,_T("GetObject: CreateBindCtx failed!"));
    return;
  }
  COM_last_HR = MkParseDisplayName(vBindCtx, objname, &vChEaten, &vMoniker);
  if (COM_last_HR != S_OK)
  {
    COM_HR_ErrMsg (COM_last_HR,_T("GetObject: Receive Moniker failed!"));
    goto CleanGetObjectMon;
  }
  VariantInit(&obj->pObjects[0]);
  #ifdef __cplusplus
    COM_last_HR = vMoniker->BindToObject(vBindCtx, NULL, IID_IDispatch, 
                                        (void **)&obj->pObjects[0].pdispVal);
  #else
    COM_last_HR = vMoniker->lpVtbl->BindToObject(vMoniker, vBindCtx, NULL, &IID_IDispatch,
                                                (void **)&obj->pObjects[0].pdispVal);
  #endif
  if (COM_last_HR != S_OK)
  {
    COM_HR_ErrMsg (COM_last_HR,_T("GetObject: Moniker BindToObject failed!"));
    goto CleanGetObjectMon;
  }
  obj->p_unknown = NULL;
  obj->pObjects[0].vt = VT_DISPATCH;
  obj->pStatus = TRUE;
  obj->ipointer = 0;
  COM_objects_cnt++;
CleanGetObjectMon:
  #ifdef __cplusplus
    if (vMoniker) vMoniker->Release();
    if (vBindCtx) vBindCtx->Release();
  #else
    if (vMoniker) vMoniker->lpVtbl->Release (vMoniker);
    if (vBindCtx) vBindCtx->lpVtbl->Release (vBindCtx);
  #endif
}

void BCX_DispatchObject(IUnknown *iobj, OBJECT *obj, BOOL b_release) {
  if (!obj) return;
  static ULONG inc_inf_ussage=0;
  if (!COM_ole_initd) COM_ole_initialize();
  obj->p_unknown = iobj;
  #ifdef __cplusplus
    inc_inf_ussage = obj->p_unknown->AddRef();
  #else
    inc_inf_ussage = obj->p_unknown->lpVtbl->AddRef(obj->p_unknown);
  #endif
  VariantInit(&obj->pObjects[0]);
  #ifdef __cplusplus
    COM_last_HR = obj->p_unknown
  ->QueryInterface(IID_IDispatch, (void**)&obj->pObjects[0].pdispVal);
  #else
    COM_last_HR = obj->p_unknown->lpVtbl
  ->QueryInterface(obj->p_unknown, &IID_IDispatch,
     (void**)&obj->pObjects[0].pdispVal);
  #endif
  if (FAILED (COM_last_HR))
  {
    COM_HR_ErrMsg (COM_last_HR,_T("QueryInterface::IID_IDispatch  failed!"));
  #ifdef __cplusplus
      obj->p_unknown->Release();
  #else
      obj->p_unknown->lpVtbl->Release(obj->p_unknown);
  #endif
    return;
  }
  if (b_release)
  {
  #ifdef __cplusplus
     iobj->Release();
  #else
     iobj->lpVtbl->Release(obj->p_unknown);
  #endif
  }
  obj->pObjects[0].vt = VT_DISPATCH;
  obj->pStatus = TRUE;
  obj->ipointer = 0;
  COM_objects_cnt++;
}

void COM_create_safearray(void) {
  long lv_param_incr = 0;
  long lv_param_incr_rev = 0;
  HRESULT hr = NO_ERROR;
  COM_PTR_safearray = 
  SafeArrayCreateVector (VT_VARIANT,0,COM_plist_index);
  if (COM_PTR_safearray == NULL)
  {
    COM_clean_plist();
    COM_last_HR = E_OUTOFMEMORY;
    COM_HR_ErrMsg (COM_last_HR,_T("SafeArrayCreate failed."));
    return;
  }
  for (lv_param_incr = COM_plist_index-1; lv_param_incr>=0; lv_param_incr-=1)
  {
    hr = SafeArrayPutElement (COM_PTR_safearray,
    &lv_param_incr, &COM_PARAMS[lv_param_incr_rev]);
    if (FAILED(hr))
    {
      wsprintf(COM_ErrMsg, _T("Param Index = %d."),lv_param_incr);
      if (COM_PTR_safearray) SafeArrayDestroy (COM_PTR_safearray);
      COM_PTR_safearray = NULL;
      COM_clean_plist();
      COM_last_HR = hr;
      COM_HR_ErrMsg (COM_last_HR,_T("SafeArrayPutElement failed!"));
      return;
    }
    lv_param_incr_rev++;
  }
  COM_clean_plist();
  if(COM_PTR_safearray)
  {
     VariantInit (&COM_PARAMS[0]);
     COM_PARAMS[0].vt = VT_ARRAY|VT_VARIANT;
     V_ARRAY (&COM_PARAMS[0]) = COM_PTR_safearray;
     COM_plist_index = 1;
  }
}

HRESULT COM_WS2AS(LPCWSTR wide_string, UINT code_page) {
  if (wide_string==NULL) return (HRESULT) NO_ERROR;
  ULONG bytes_copied=0;
  ULONG temp_ansi_len = (ULONG)WideCharToMultiByte (
  code_page,0,wide_string,-1,COM_psz_tmp,0,NULL,NULL);
  if (temp_ansi_len == 0) return (HRESULT) NO_ERROR;
  if (COM_zstr_size < temp_ansi_len)
  {
    if (COM_psz_tmp) free(COM_psz_tmp);
    COM_psz_tmp = (char*)calloc(temp_ansi_len+1,sizeof(char));
    if (NULL == COM_psz_tmp)
    {
      COM_zstr_size = 0;
      return E_OUTOFMEMORY;
    } 
    COM_zstr_size = temp_ansi_len;
  }
  if ((bytes_copied = WideCharToMultiByte(code_page, 0, wide_string, -1,
       COM_psz_tmp, temp_ansi_len,NULL, NULL))==0)
  {
    return HRESULT_FROM_WIN32 (GetLastError());
  } 
  COM_psz_tmp [bytes_copied] = '\0';
  return (HRESULT)NO_ERROR;
}

HRESULT COM_AS2WS(LPCSTR ansi_string, UINT code_page) {
  if (!*ansi_string) return (HRESULT)NO_ERROR;
  ULONG ansi_str_len = (ULONG)strlen (ansi_string);
  if (!ansi_str_len) return (HRESULT)NO_ERROR;
  ULONG wide_str_len = (ansi_str_len * 2);
  if (COM_wstr_size < wide_str_len)
  {
    if (COM_LPWSTR_temp)
       CoTaskMemFree((void*)COM_LPWSTR_temp);
    COM_LPWSTR_temp = (LPOLESTR)CoTaskMemAlloc (wide_str_len);
    if (NULL == COM_LPWSTR_temp)
    {
      COM_wstr_size = 0;
      return E_OUTOFMEMORY;
    } 
    COM_wstr_size = wide_str_len;
  }
  if (MultiByteToWideChar (code_page, MB_PRECOMPOSED, ansi_string, 
      ansi_str_len, COM_LPWSTR_temp, wide_str_len)==0)
         return HRESULT_FROM_WIN32(GetLastError());
  COM_LPWSTR_temp[ansi_str_len] = L'\0';
  return (HRESULT)NO_ERROR;
}

void COM_FREE_TEMP_WIDE_STRING(void) {
  if (COM_LPWSTR_temp)
  {
    CoTaskMemFree ((void*)COM_LPWSTR_temp);
    COM_LPWSTR_temp = NULL;
  }
}

void COM_FREE_TEMP_ANSI_STRING (void) {
  if (COM_psz_tmp)
  {
    free (COM_psz_tmp);
    COM_psz_tmp = NULL;
  }
}

// *************************************************
//                  Main Program 
// *************************************************

int main(int argc, char *argv[])
{
  UprCase=(unsigned char*)calloc(257,1), UprCase = MakeUCaseTbl();
  Session=COM("Microsoft.Update.Session");
  lstrcpy(COM_ErrMsg, _T("CreateupdateSearcher()"));
  COM_plist_index = 0;
  COM_invoke(&Session, L"CreateupdateSearcher",COM_PROPS, &COM_vt_result);
  
  if (COM_vt_result.vt != VT_DISPATCH) {
      COM_last_HR = VariantChangeType(&COM_vt_result,&COM_vt_result,0,VT_DISPATCH);
  }
  
  if (FAILED(COM_last_HR)) {
    strcpy(COM_ErrMsg ,"Searcher");
    COM_HR_ErrMsg (COM_last_HR, _T("VariantChangeType failed. Expected IDispatch*"));
  } else {
    VariantInit(&Searcher.pObjects[0]);
    COM_last_HR = VariantCopy(&Searcher.pObjects[0],&COM_vt_result);
  
    if (FAILED(COM_last_HR)) {
      strcpy(COM_ErrMsg ,"Searcher");
      COM_HR_ErrMsg (COM_last_HR, _T("VariantCopy failed."));
    } else {
      COM_objects_cnt++;
      Searcher.pStatus = TRUE;
      Searcher.ipointer = 0;
    }  
  } 
  VariantClear(&COM_vt_result);
  if (0 == COM_reset_chain) COM_reset_disp_chain(&Session);
  lstrcpy(COM_ErrMsg, _T("ClientApplicationID"));
  COM_PARAMS[0].vt = VT_BSTR;
  COM_PARAMS[0].bstrVal = SysAllocString(L"uphide");
  COM_plist_index = 1;
  
  if (!COM_BCX_ERROR)
    COM_invoke(&Searcher, L"ClientApplicationID", DISPATCH_PROPERTYPUT, NULL);
  
  printf("%s\n","Checking for updates...");
  lstrcpy(COM_ErrMsg, _T("Search(\"IsInstalled=0 And IsHidden=0\")"));
  COM_PARAMS[0].vt = VT_BSTR;
  COM_PARAMS[0].bstrVal = SysAllocString(L"IsInstalled=0 And IsHidden=0");
  COM_plist_index = 1;
  COM_invoke(&Searcher, L"Search",COM_PROPS, &COM_vt_result);
  
  if (COM_vt_result.vt != VT_DISPATCH) {
      COM_last_HR = VariantChangeType(&COM_vt_result,&COM_vt_result,0,VT_DISPATCH);
  }
  
  if (FAILED(COM_last_HR)) {
    strcpy(COM_ErrMsg ,"SearchResult");
    COM_HR_ErrMsg (COM_last_HR, _T("VariantChangeType failed. Expected IDispatch*"));
  } else {
    VariantInit(&SearchResult.pObjects[0]);
    COM_last_HR = VariantCopy(&SearchResult.pObjects[0],&COM_vt_result);
  
    if (FAILED(COM_last_HR)) {
      strcpy(COM_ErrMsg ,"SearchResult");
      COM_HR_ErrMsg (COM_last_HR, _T("VariantCopy failed."));
    } else {
      COM_objects_cnt++;
      SearchResult.pStatus = TRUE;
      SearchResult.ipointer = 0;
    }  
  } 
  VariantClear(&COM_vt_result);
  if (0 == COM_reset_chain) COM_reset_disp_chain(&Searcher);
  lstrcpy(COM_ErrMsg, _T("Updates"));
  COM_get_next_dispatch(&SearchResult, L"Updates");
  lstrcpy(COM_ErrMsg, _T("Count"));
  COM_invoke(&SearchResult, L"Count",COM_PROPS, &COM_vt_result);
  
  if (COM_vt_result.vt != VT_INT)
    {
    VariantChangeType(&COM_vt_result,
   &COM_vt_result,0, VT_INT);
    }
  NumUpdates = COM_vt_result.intVal;
  VariantClear(&COM_vt_result);
  if (0 == COM_reset_chain) COM_reset_disp_chain(&SearchResult);
  if(NumUpdates==0 ){
      printf("%s\n","No updates were found.");
      UNCOM(SearchResult);
      UNCOM(Searcher);
      UNCOM(Session);
      fflush(stdout);
      ExitProcess(0);
    }
  printf("%s\n","Enter the numbers of the updates you want to hide, seperated by spaces, and press enter. Leave blank to exit.");
  lstrcpy(COM_ErrMsg, _T("Updates"));
  COM_invoke(&SearchResult, L"Updates",COM_PROPS, &COM_vt_result);
  
  if (COM_vt_result.vt != VT_DISPATCH) {
      COM_last_HR = VariantChangeType(&COM_vt_result,&COM_vt_result,0,VT_DISPATCH);
  }
  
  if (FAILED(COM_last_HR)) {
    strcpy(COM_ErrMsg ,"Updates");
    COM_HR_ErrMsg (COM_last_HR, _T("VariantChangeType failed. Expected IDispatch*"));
  } else {
    VariantInit(&Updates.pObjects[0]);
    COM_last_HR = VariantCopy(&Updates.pObjects[0],&COM_vt_result);
  
    if (FAILED(COM_last_HR)) {
      strcpy(COM_ErrMsg ,"Updates");
      COM_HR_ErrMsg (COM_last_HR, _T("VariantCopy failed."));
    } else {
      COM_objects_cnt++;
      Updates.pStatus = TRUE;
      Updates.ipointer = 0;
    }  
  } 
  VariantClear(&COM_vt_result);
  if (0 == COM_reset_chain) COM_reset_disp_chain(&SearchResult);
   VariantClear(&COM_ack_var);
   COM_enum_var = NULL;
   COM_long_coll = 0;
  OBJECT Item;
  ZeroMemory((PVOID)&Item,sizeof(OBJECT));
  COM_GetEnum_iface = TRUE;
  COM_invoke(&Updates, L"",
  DISPATCH_PROPERTYGET|DISPATCH_METHOD, &COM_ack_var);
  COM_reset_disp_chain(&Updates);
  COM_GetEnum_iface = FALSE;
  
  for(;;) { // for each construction ...
    if (FAILED(COM_last_HR)) {
      COM_HR_ErrMsg (COM_last_HR,_T("Get object enumerator: No Collections!"));
      break;
    } 
    if (COM_ack_var.vt != VT_DISPATCH && COM_ack_var.vt != VT_UNKNOWN) {
      COM_last_HR = E_NOINTERFACE;
      COM_HR_ErrMsg (COM_last_HR,_T("Enum interface not available! Collections unavailable!"));
      VariantClear(&COM_ack_var);
      break;
    } 
    if (COM_ack_var.vt == VT_DISPATCH) {
        #ifdef __cplusplus
        COM_last_HR = COM_ack_var.pdispVal
      ->QueryInterface(IID_IEnumVARIANT, (void **)&COM_enum_var);
        #else
         COM_last_HR = COM_ack_var.pdispVal->lpVtbl
       ->QueryInterface(COM_ack_var.pdispVal,
         &IID_IEnumVARIANT, (void **)&COM_enum_var);
      #endif
  
      if (FAILED(COM_last_HR)) {
         COM_HR_ErrMsg (COM_last_HR,_T("QueryInterface: Get enum variant: No Collections!"));
         VariantClear(&COM_ack_var);
         break;
      } 
    } else if (COM_ack_var.vt == VT_UNKNOWN) {
        #ifdef __cplusplus
          COM_last_HR = COM_ack_var.punkVal
         ->QueryInterface(IID_IEnumVARIANT, (void **)&COM_enum_var);
        #else
          COM_last_HR = COM_ack_var.punkVal->lpVtbl
        ->QueryInterface(COM_ack_var.punkVal, 
          &IID_IEnumVARIANT, (void **)&COM_enum_var);
      #endif
  
      if (FAILED(COM_last_HR)) {
        COM_HR_ErrMsg (COM_last_HR,_T("QueryInterface: Get enum variant: No Collections!"));
        VariantClear(&COM_ack_var);
        break;
      } 
    } 
    VariantClear(&COM_ack_var);
    break;
    }
  
      while(COM_enum_var) {
        BCX_SetNothing(&Item);
        #ifdef __cplusplus
          COM_last_HR = COM_enum_var
       ->Next(1, &Item.pObjects[0], &COM_long_coll);
        #else
          COM_last_HR = COM_enum_var->lpVtbl
       ->Next(COM_enum_var, 1, &Item.pObjects[0], &COM_long_coll);
        #endif
        if (FAILED(COM_last_HR)) {
          COM_HR_ErrMsg (COM_last_HR,_T("Enumeration failed! No Collections!"));
        #ifdef __cplusplus
           if(COM_enum_var) COM_enum_var->Release();
        #else
           if(COM_enum_var) COM_enum_var->lpVtbl
         ->Release(COM_enum_var);
        #endif
        COM_enum_var = NULL;
        break;
      } 
      if (Item.pObjects[0].vt != VT_DISPATCH)
        {
          VariantClear(&Item.pObjects[0]);
          COM_long_coll = 0;
        }
      else
        { 
          Item.pStatus = TRUE;
          Item.ipointer = 0;
          COM_objects_cnt++;
        }
      if (0 == COM_long_coll) {
      #ifdef __cplusplus
         if(COM_enum_var) COM_enum_var->Release();
      #else
         if(COM_enum_var) COM_enum_var->lpVtbl
       ->Release(COM_enum_var);
      #endif
      BCX_SetNothing(&Item);
      COM_enum_var = NULL;
      break;
      }
      char    Title[BCXSTRSIZE]= {0};
      lstrcpy(COM_ErrMsg, _T("Title"));
      COM_invoke(&Item, L"Title",COM_PROPS, &COM_vt_result);
  
      if(COM_vt_result.vt == VT_NULL){
        *Title=0;
      } else {
      COM_last_HR = VariantChangeType(&COM_vt_result,&COM_vt_result,0,VT_BSTR);
  
      if (FAILED(COM_last_HR)) {
        COM_HR_ErrMsg (COM_last_HR,_T("VariantChangeType (BSTR) Failed!"));
      } else {
        #ifndef UNICODE
        COM_last_HR = COM_WS2AS(COM_vt_result.bstrVal);
        if (FAILED(COM_last_HR)) {
          COM_HR_ErrMsg (COM_last_HR,_T("COM Get Property Failed! W2A failure!"));
        } else {
         strcpy(Title, COM_psz_tmp);
        }
        #else
        lstrcpy(Title,COM_vt_result.bstrVal);
        #endif
      }
      }
      VariantClear(&COM_vt_result);
      if (0 == COM_reset_chain) COM_reset_disp_chain(&Item);
      printf("%s%s%s\n",str(Index,1),": ",Title);
      Index++;
    }
  *Input=0, KBinput(),ScanError=scan(InputBuffer,"%s",Input);
  if(Input[0]==0 ){
      UNCOM(Updates);
      UNCOM(SearchResult);
      UNCOM(Searcher);
      UNCOM(Session);
      fflush(stdout);
      ExitProcess(0);
    }
  NumSelections=Split(ToHide,Input," ",0);
  for(I=0; I<=NumSelections-1; I++)
    {
      int     CurIndex= {0};
      CurIndex=VAL(ToHide[I]);
      if(CurIndex<=0||CurIndex>NumUpdates ){
          continue;
        }
      lstrcpy(COM_ErrMsg, _T("Item(CurIndex-1)"));
      COM_PARAMS[0].vt = VT_I4;
      COM_PARAMS[0].lVal = (long)CurIndex-1;
      COM_plist_index = 1;
      COM_invoke(&Updates, L"Item",COM_PROPS, &COM_vt_result);
  
      if (COM_vt_result.vt != VT_DISPATCH) {
          COM_last_HR = VariantChangeType(&COM_vt_result,&COM_vt_result,0,VT_DISPATCH);
      }
  
      if (FAILED(COM_last_HR)) {
        strcpy(COM_ErrMsg ,"Update");
        COM_HR_ErrMsg (COM_last_HR, _T(   "VariantChangeType failed. Expected IDispatch*"));
      } else {
        VariantInit(&Update.pObjects[0]);
        COM_last_HR = VariantCopy(&Update.pObjects[0],&COM_vt_result);
  
        if (FAILED(COM_last_HR)) {
          strcpy(COM_ErrMsg ,"Update");
          COM_HR_ErrMsg (COM_last_HR, _T("VariantCopy failed."));
        } else {
          COM_objects_cnt++;
          Update.pStatus = TRUE;
          Update.ipointer = 0;
        }  
      } 
      VariantClear(&COM_vt_result);
      if (0 == COM_reset_chain) COM_reset_disp_chain(&Updates);
      if(ISOBJECT(Update)){
          char    Title[BCXSTRSIZE]= {0};
          lstrcpy(COM_ErrMsg, _T("Title"));
          COM_invoke(&Update, L"Title",COM_PROPS, &COM_vt_result);
  
          if(COM_vt_result.vt == VT_NULL){
            *Title=0;
          } else {
          COM_last_HR = VariantChangeType(&COM_vt_result,&COM_vt_result,0,VT_BSTR);
  
          if (FAILED(COM_last_HR)) {
            COM_HR_ErrMsg (COM_last_HR,_T("VariantChangeType (BSTR) Failed!"));
          } else {
            #ifndef UNICODE
            COM_last_HR = COM_WS2AS(COM_vt_result.bstrVal);
            if (FAILED(COM_last_HR)) {
              COM_HR_ErrMsg (COM_last_HR,_T("COM Get Property Failed! W2A failure!"));
            } else {
             strcpy(Title, COM_psz_tmp);
            }
            #else
            lstrcpy(Title,COM_vt_result.bstrVal);
            #endif
          }
          }
          VariantClear(&COM_vt_result);
          if (0 == COM_reset_chain) COM_reset_disp_chain(&Update);
          lstrcpy(COM_ErrMsg, _T("IsHidden"));
          COM_PARAMS[0].vt = VT_BOOL;
          COM_PARAMS[0].boolVal = VARIANT_TRUE;
          COM_plist_index = 1;
  
          if (!COM_BCX_ERROR)
            COM_invoke(&Update, L"IsHidden", DISPATCH_PROPERTYPUT, NULL);
  
          printf("%s%s\n","Hid ",Title);
          UNCOM(Update);
        }
    }
  Pause();
  UNCOM(Updates);
  UNCOM(SearchResult);
  UNCOM(Searcher);
  UNCOM(Session);
  return EXIT_SUCCESS;   // End of main program
  }
