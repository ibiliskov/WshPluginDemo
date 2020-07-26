#define INITGUID
#define PI 3.14159265

#include <activscp.h>
#include <comutil.h>
#include <math.h>
#include <stdio.h>
#include <string.h>
#include <windows.h>

#include "CGraphics.h"
#include "resource.h"

HWND hwndMain;
HMODULE dll[1000];
HRESULT hr;
IActiveScript *script;
IActiveScriptParse *parse;

FILE *fp;
char *scriptFilename;
char scriptContent[2048];

int values[500];
PAINTSTRUCT Ps;
CGraphics g;

int WINAPI GetVal(int index)
{
  if (index >= 0 && index < 500)
    return values[index];
  else
    return 0;
}

void WINAPI SetVal(int index, double value)
{
  if (index >= 0 && index < 500)
    values[index] = (int)value;
}

void WINAPI Invalidate()
{
  InvalidateRect(hwndMain, NULL, TRUE);
}

void WINAPI Debug(int x)
{
}

struct ScriptInterface : public IDispatch
{
  DWORD cdcl;
  long WINAPI QueryInterface(REFIID riid, void **object)
  {
    *object = IsEqualIID(riid, IID_IDispatch) ? this : 0;
    return *object ? 0 : E_NOINTERFACE;
  }
  DWORD WINAPI AddRef()
  {
    return 0;
  }
  DWORD WINAPI Release()
  {
    return 0;
  }
  long WINAPI GetTypeInfoCount(UINT *)
  {
    return 0;
  }
  long WINAPI GetTypeInfo(UINT, LCID, ITypeInfo **)
  {
    return 0;
  }

  // Register methods (or vars)
  long WINAPI GetIDsOfNames(REFIID riid, WCHAR **name, UINT cnt, LCID lcid, DISPID *id)
  {
    for (DWORD j = 0; j < cnt; j++)
    {
      char buf[1000];
      DWORD k;
      WideCharToMultiByte(0, 0, name[j], -1, buf, sizeof(buf), 0, 0);

      // Two loops to check prefix _ added in script (one for sdcall, second for cdecl)
      for (k = 0; k < 2; k++)
      {
        // Check app functions
        if (strcmp(buf + k, "GetVal") == 0)
        {
          id[j] = (DISPID)GetVal;
          break;
        }
        else if (strcmp(buf + k, "SetVal") == 0)
        {
          id[j] = (DISPID)SetVal;
          break;
        }
        else if (strcmp(buf + k, "Invalidate") == 0)
        {
          id[j] = (DISPID)Invalidate;
          break;
        }
        else if (strcmp(buf + k, "Debug") == 0)
        {
          id[j] = (DISPID)Debug;
          break;
        }
        else

          // win32 api in known dlls
          for (int i = 0; (dll[i] || dll[i - 1]); i++)
          {
            if ((id[j] = (DISPID)GetProcAddress(dll[i], buf + k)))
              break;
          }
        if (id[j])
          break;
      }
      cdcl = k;
      if (!id[j])
        return E_FAIL;
    }
    return 0;
  }

  // Excuted from the script
  long WINAPI Invoke(DISPID id, REFIID riid, LCID lcid, WORD flags, DISPPARAMS *arg, VARIANT *ret, EXCEPINFO *excp,
                     UINT *err)
  {
    if (id == (DISPID)SetVal)
    {
      if (arg->cArgs >= 2 && arg->rgvarg[1].vt == VT_I4 && (arg->rgvarg[0].vt == VT_I4 || arg->rgvarg[0].vt == VT_R8))
        if (arg->rgvarg[0].vt == VT_I4)
          SetVal(arg->rgvarg[1].ulVal, arg->rgvarg[0].ulVal);
        else if (arg->rgvarg[0].vt == VT_R8)
          SetVal(arg->rgvarg[1].ulVal, arg->rgvarg[0].dblVal);
        else
          return E_FAIL;
      return 0;
    }

    if (id)
    {
      // Check stdcall and cdecl conventions
      int i = cdcl ? arg->cArgs : -1, result = 0, stack = arg->cArgs * 4;
      char *args[100] = {0};

      while ((cdcl ? (--i > -1) : (++i < arg->cArgs)))
      {
        DWORD param = arg->rgvarg[i].ulVal;

        // Convert unicode string params to ansi since most apis are ansi
        if (arg->rgvarg[i].vt == VT_BSTR)
        {
          WCHAR *w = arg->rgvarg[i].bstrVal;
          WideCharToMultiByte(0, 0, w, -1, args[i] = new char[wcslen(w) + 1], wcslen(w) + 1, 0, 0);
          param = (DWORD)args[i];
        }
        else if (arg->rgvarg[i].vt == VT_R8)
        {
          int d = (int)arg->rgvarg[i].dblVal;
          param = (DWORD)d;
        }
        _asm push param;
      }

      // cdecl: push params in reverse order and cleanup the stack after call
      _asm call id
      _asm mov result, eax
      if (cdcl)
      {
          _asm add esp, stack
      }
      
      i = -1;
      while (++i < arg->cArgs)
      {
        if (args[i])
        {
          delete args[i];
        }
      }

      // Return value to script (just unsigned integer and unicode string types)
      if (ret)
        ret->ullVal = result;
      char *c = (char *)result;
      if (ret)
        ret->vt = VT_UI4;
      __try
      {
        if (!c[1] && (*c < '0' || *c > '9'))
          ret->vt = VT_BSTR;
        ret->bstrVal = SysAllocString((WCHAR *)c);
      }
      __except (EXCEPTION_EXECUTE_HANDLER)
      {
      }
      return 0;
    }
    return E_FAIL;
  }
};

struct ScriptHost : public IActiveScriptSite
{
  ScriptInterface Interface;
  long WINAPI QueryInterface(REFIID riid, void **object)
  {
    *object = (IsEqualIID(riid, IID_IActiveScriptSite)) ? this : 0;
    return *object ? 0 : E_NOINTERFACE;
  }
  DWORD WINAPI AddRef()
  {
    return 0;
  }
  DWORD WINAPI Release()
  {
    return 0;
  }
  long WINAPI GetLCID(DWORD *lcid)
  {
    *lcid = LOCALE_USER_DEFAULT;
    return 0;
  }
  long WINAPI GetDocVersionString(BSTR *ver)
  {
    *ver = 0;
    return 0;
  }
  long WINAPI OnScriptTerminate(const VARIANT *, const EXCEPINFO *)
  {
    return 0;
  }
  long WINAPI OnStateChange(SCRIPTSTATE state)
  {
    return 0;
  }
  long WINAPI OnEnterScript()
  {
    return 0;
  }
  long WINAPI OnLeaveScript()
  {
    return 0;
  }
  long WINAPI GetItemInfo(const WCHAR *name, DWORD req, IUnknown **obj, ITypeInfo **type)
  {
    if (req & SCRIPTINFO_IUNKNOWN)
      *obj = &Interface;
    if (req & SCRIPTINFO_ITYPEINFO)
      *type = 0;
    return 0;
  }
  long WINAPI OnScriptError(IActiveScriptError *err)
  {
    EXCEPINFO e;
    err->GetExceptionInfo(&e);
    MessageBoxW(0, e.bstrDescription, e.bstrSource, 0);
    return 0;
  }
};

char *OpenFile()
{
  OPENFILENAMEA ofn;
  char fileName[MAX_PATH] = "";
  ZeroMemory(&ofn, sizeof(ofn));

  ofn.lStructSize = sizeof(OPENFILENAME);
  ofn.hwndOwner = NULL;
  ofn.lpstrFilter = "JavaScript Files (*.js)\0*.js\0All Files (*.*)\0*.*";
  ofn.lpstrFile = fileName;
  ofn.nMaxFile = MAX_PATH;
  ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
  ofn.lpstrDefExt = "js";

  GetOpenFileNameA(&ofn);

  for (int i = 0; i < strlen(fileName); i++)
  {
    if (fileName[i] == '\\')
      fileName[i] = '/';
  }

  char *fileNameReturn = (char *)malloc(MAX_PATH * sizeof(char));
  return strcpy(fileNameReturn, fileName);
}

char *SaveFile()
{
  OPENFILENAMEA ofn;
  char fileName[MAX_PATH] = "";
  ZeroMemory(&ofn, sizeof(ofn));

  ofn.lStructSize = sizeof(OPENFILENAME);
  ofn.hwndOwner = NULL;
  ofn.lpstrFilter = "JavaScript Files (*.js)\0*.js\0All Files (*.*)\0*.*";
  ofn.lpstrFile = fileName;
  ofn.nMaxFile = MAX_PATH;
  ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_HIDEREADONLY;
  ofn.lpstrDefExt = "js";

  GetSaveFileNameA(&ofn);

  for (int i = 0; i < strlen(fileName); i++)
  {
    if (fileName[i] == '\\')
      fileName[i] = '/';
  }

  char *fileNameReturn = (char *)malloc(MAX_PATH * sizeof(char));
  return strcpy(fileNameReturn, fileName);
}

void Initialize()
{
  for (int i = 0; i < 500; i++)
  {
    values[i] = 0;
  }
}

BOOL CALLBACK DlgProc(HWND hwnd, UINT Message, WPARAM wParam, LPARAM lParam)
{
  int wmId, wmEvent;
  switch (Message)
  {
  case WM_INITDIALOG:
    hwndMain = hwnd;
    // SetDlgItemText(hwnd, IDC_EDIT, L"number = 'text 1'.match(/\\d/);\r\nresult = app.test('hello'+' world
    // '+number);\r\n"
    //        L"app.MessageBoxA(0,'WIN'+32+' or any dll api called from JavaScript','hello world '+result,0);\r\n");
  case WM_COMMAND:
    wmId = LOWORD(wParam);
    wmEvent = HIWORD(wParam);

    switch (wmId)
    {
    case ID_MENU_OPEN:
      scriptFilename = OpenFile();
      fp = fopen(scriptFilename, "r");
      if (fp != NULL)
      {
        SetDlgItemTextA(hwnd, IDC_EDIT, "");

        while (fgets(scriptContent, sizeof scriptContent, fp) != NULL)
        {
          int ndx = GetWindowTextLength(GetDlgItem(hwnd, IDC_EDIT));
          SendDlgItemMessage(hwnd, IDC_EDIT, EM_SETSEL, (WPARAM)ndx, (LPARAM)ndx);
          SendDlgItemMessageA(hwnd, IDC_EDIT, EM_REPLACESEL, 0, (LPARAM)scriptContent);
        }
        fclose(fp);
      }
      break;
    case ID_MENU_SAVE:
      scriptFilename = SaveFile();
      fp = fopen(scriptFilename, "w");
      if (fp != NULL)
      {
        int len = GetWindowTextLength(GetDlgItem(hwnd, IDC_EDIT));
        if (len > 0)
        {
          char *buf = (char *)malloc(len * sizeof(char));
          GetDlgItemTextA(hwnd, IDC_EDIT, buf, len + 1);
          fputs(buf, fp);
        }
        fclose(fp);
      }
      break;
    case ID_MENU_RUN: {
      int len = GetWindowTextLength(GetDlgItem(hwnd, IDC_EDIT));
      if (len > 0)
      {
        int i;
        wchar_t *buf;
        buf = (wchar_t *)GlobalAlloc(GPTR, (len + 1) * 2);
        GetDlgItemText(hwnd, IDC_EDIT, buf, (len + 1) * 2);

        hr = parse->InitNew();
        hr = parse->ParseScriptText(buf, 0, 0, 0, 0, 0, 0, 0, 0);
        GlobalFree(buf);

        script->SetScriptState(SCRIPTSTATE_CONNECTED);
      }
    }
    break;
    case ID_MENU_EXIT:
      DestroyWindow(hwnd);
      break;
    default:
      return DefWindowProc(hwnd, Message, wParam, lParam);
    }

    break;
  case WM_CLOSE:
    EndDialog(hwnd, 0);
    break;
  case WM_PAINT:
    g.hDC = BeginPaint(hwnd, &Ps);
    g.DrawAxes(10, 530);
    g.DrawInterp(values, 500);
    EndPaint(hwnd, &Ps);
    break;
  default:
    return FALSE;
  }
  return TRUE;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
  CoInitialize(0);
  Initialize();

  wchar_t *name[] = {L"ntdll",   L"msvcrt", L"kernel32", L"user32", L"advapi32", L"shell32", L"wsock32",
                     L"wininet", L"<",      L">",        0};
  for (int i = 0; name[i]; i++)
    dll[i] = LoadLibrary(name[i]);

  // Set language
  GUID guid;
  CLSIDFromProgID(L"JavaScript", &guid);

  hr = CoCreateInstance(guid, 0, CLSCTX_ALL, IID_IActiveScript, (void **)&script);
  if (!script)
    return 1;
  hr = script->QueryInterface(IID_IActiveScriptParse, (void **)&parse);

  ScriptHost host;
  script->SetScriptSite((IActiveScriptSite *)&host);
  script->AddNamedItem(L"app", SCRIPTITEM_ISVISIBLE | SCRIPTITEM_NOCODE);

  DialogBox(hInstance, MAKEINTRESOURCE(IDD_MAIN), NULL, DlgProc);

  script->Close();

  return 0;
};