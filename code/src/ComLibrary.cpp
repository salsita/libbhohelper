#include "StdAfx.h"
#include "libbhohelper.h"

namespace LIB_BhoHelper
{

/*============================================================================
 * class ComLibrary
 */
//----------------------------------------------------------------------------
//  CTOR
ComLibrary::ComLibrary() :
    mhDll(NULL),
    mDllGetClassObject(NULL),
    mDllCanUnloadNow(NULL)
{
}

//----------------------------------------------------------------------------
//  DTOR
ComLibrary::~ComLibrary()
{
  // force unloading
  unload(TRUE);
}

//----------------------------------------------------------------------------
//  load a module from a certain path
HRESULT ComLibrary::load(LPOLESTR szDllName, HINSTANCE aHinstanceRel)
{
  if (mhDll)
    { return S_FALSE; }  // is already open

  OLECHAR fileName[MAX_PATH] = {0};

  // Add a bit of securty:
  // 1) deny network paths
  if (PathIsNetworkPath(szDllName))
      { return E_INVALIDARG; }

  // 2) If path is a relative path...
  if (PathIsRelative(szDllName)) {
    OLECHAR fileName2[MAX_PATH] = {0};
    // make it one based on aHinstanceRel
    if (
          !::GetModuleFileName(aHinstanceRel, fileName2, MAX_PATH)
      || !PathRemoveFileSpec(fileName2)
      || !PathCombine(fileName, fileName2, szDllName)
        )
    { return ATL::AtlHresultFromLastError(); }
  }
  else {
    wcscpy_s<MAX_PATH>(fileName, szDllName);
  }

  // load library
  mhDll = ::CoLoadLibrary(fileName, FALSE);
  if (!mhDll)
    { return ATL::AtlHresultFromLastError(); }

  // get DllGetClassObject and DllCanUnloadNow
  mDllGetClassObject = (FnDllGetClassObject)::GetProcAddress(mhDll, "DllGetClassObject");
  HRESULT hr = ATL::AtlHresultFromLastError();
  // DllCanUnloadNow is optional
  mDllCanUnloadNow = (FnDllCanUnloadNow)::GetProcAddress(mhDll, "DllCanUnloadNow");

  if (!mDllGetClassObject)
    { return hr; }
  return S_OK;
}

//----------------------------------------------------------------------------
//  unload
HRESULT ComLibrary::unload(BOOL bForce)
{
  if (!mhDll)
    { return S_OK; } // we are done already

  // call DllCanUnloadNow
  HRESULT hr = this->DllCanUnloadNow();
  ATLASSERT(SUCCEEDED(hr));

  // if can NOT unload AND NOT force...
  if ( (S_OK != hr) && !bForce )
    { return hr; } // ... don't close

  // close: free and clean up
  ::CoFreeLibrary(mhDll);
  mDllGetClassObject = NULL;
  mDllCanUnloadNow = NULL;
  mhDll = NULL;
  return S_OK;
}

//----------------------------------------------------------------------------
//  DllGetClassObject
HRESULT ComLibrary::DllGetClassObject(REFCLSID aClsid, 
  REFIID aIid, LPVOID FAR* ppv)
{
  return (mDllGetClassObject)
      ? mDllGetClassObject(aClsid, aIid, ppv)
      : E_UNEXPECTED;
}

//----------------------------------------------------------------------------
//  DllCanUnloadNow
HRESULT ComLibrary::DllCanUnloadNow()
{
  return (mDllCanUnloadNow)
      ? mDllCanUnloadNow()
      : S_OK; // default: yes, can unload
}

}// namespace LIB_BhoHelper
