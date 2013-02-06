#include "StdAfx.h"
#include "libbhohelper.h"

#include <algorithm>

#include <shlguid.h>
#include <OleAcc.h> // For ObjectFromLresult
#pragma comment(lib, "Oleacc.lib")

namespace LIB_BhoHelper
{
  HRESULT getHTMLDocumentForHWND(HWND hwnd, IHTMLDocument2** aRet)
  {
    DWORD dwMsg = ::RegisterWindowMessage(L"WM_HTML_GETOBJECT");
    LRESULT lResult = 0;
    ::SendMessageTimeout(hwnd, dwMsg, 0, 0, SMTO_ABORTIFHUNG, 1000, (DWORD*) &lResult);
    if (lResult) {
      return ::ObjectFromLresult(lResult, IID_IHTMLDocument2, 0, (void**) aRet);
    }
    else {
      return E_FAIL;
    }
  }

  //----------------------------------------------------------------------------
  //
  HRESULT getBrowserForHTMLDocument(IHTMLDocument2* aDocument, IWebBrowser2** aRet)
  {
    CComPtr<IHTMLWindow2> win;
    HRESULT hr = aDocument->get_parentWindow(&win);
    if (SUCCEEDED(hr)) {
      CComQIPtr<IServiceProvider> sp(win);
      CComPtr<IWebBrowser2> pWebBrowser;
      if (sp) {
        hr = sp->QueryService(IID_IWebBrowserApp, IID_IWebBrowser2, (void**) aRet);
      }
      else {
        return E_NOINTERFACE;
      }
    }
    return hr;
  }

  //----------------------------------------------------------------------------
  //
  HRESULT createIDispatchFromCreator(LPDISPATCH aCreator, VARIANT* aRet)
  {
    if (!aRet) {
      return E_POINTER;
    }
    CComQIPtr<IDispatchEx> creator(aCreator);
    if (!creator) {
      return E_NOINTERFACE;
    }
    DISPPARAMS params = {0};
    HRESULT hr = creator->InvokeEx(DISPID_VALUE, LOCALE_USER_DEFAULT, DISPATCH_CONSTRUCT, &params, aRet, NULL, NULL);
    if (hr != S_OK) {
      return hr;
    }
    return S_OK;
  }

  //----------------------------------------------------------------------------
  //
  HRESULT addJSArrayToVariantVector(LPDISPATCH aArrayDispatch, VariantVector &aVariantVector, bool aReverse)
  {
    CComQIPtr<IDispatch> dispArray(aArrayDispatch);
    if (!dispArray) {
        return E_NOINTERFACE;
    }

    // Get array length DISPID
    DISPID dispidLength;
    CComBSTR bstrLength(L"length");
    //HRESULT hr = dispexArray->GetDispID(bstrLength, fdexNameCaseSensitive, &dispidLength);
    HRESULT hr = dispArray->GetIDsOfNames(IID_NULL, &bstrLength, 1, LOCALE_SYSTEM_DEFAULT, &dispidLength);


    if (FAILED(hr)) {
        return hr;
    }

    // Get length value using InvokeEx()
    CComVariant varLength;
    DISPPARAMS dispParamsNoArgs = {0};
    //hr = dispexArray->InvokeEx(dispidLength, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispParamsNoArgs, &varLength, NULL, NULL);
    hr = dispArray->Invoke(dispidLength, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispParamsNoArgs, &varLength, NULL, NULL);
    if (FAILED(hr)) {
        return hr;
    }

    ATLASSERT(varLength.vt == VT_I4);
    const int count = varLength.intVal;
    aVariantVector.reserve(aVariantVector.size() + count); //ensure that we will not reallocate too often



    // For each element in source array:
    for (int i = 0 ; i < count; ++i) //values are reverted
    {
      CString strIndex;
      strIndex.Format(L"%d", i);

      // Convert to BSTR, as GetDispID() wants BSTR's
      CComBSTR bstrIndex(strIndex);
      DISPID dispidIndex;
      //hr = dispexArray->GetDispID(bstrIndex, fdexNameCaseSensitive, &dispidIndex);
      hr = dispArray->GetIDsOfNames(IID_NULL, &bstrIndex, 1, LOCALE_SYSTEM_DEFAULT, &dispidIndex);

      CComVariant varItem;
      //if we cannot obtain item with given index - it means that its value is undefined
      if (SUCCEEDED(hr)) {
      // Get array item value using InvokeEx()
        //hr = dispexArray->InvokeEx(dispidIndex, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispParamsNoArgs, &varItem, NULL, NULL);
        hr = dispArray->Invoke(dispidIndex, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispParamsNoArgs, &varItem, NULL, NULL);
        if (FAILED(hr)) {
          return hr;
        }
      }
      aVariantVector.push_back(varItem);
    }
    if (aReverse) {
      std::reverse(aVariantVector.begin(),aVariantVector.end());
    }
    return S_OK;
  }

  //----------------------------------------------------------------------------
  //
  HRESULT constructSafeArrayFromVector(const VariantVector &aVariantVector, VARIANT &aSafeArray)
  {
    HRESULT hr = VariantClear(&aSafeArray);
    if (FAILED(hr)) {
      return hr;
    }
    SAFEARRAYBOUND bounds [] = { static_cast<ULONG>(aVariantVector.size()), 0 };
    aSafeArray.vt = VT_ARRAY | VT_VARIANT;
    aSafeArray.parray = SafeArrayCreate(VT_VARIANT, 1, bounds);
    VARIANT *elements;
    SafeArrayAccessData(aSafeArray.parray, (void**)&elements);
    for (size_t i = 0; i < aVariantVector.size(); ++i) {
      hr = ::VariantCopy(&elements[i],&aVariantVector[i]);
      if (FAILED(hr)) {
        return hr;
      }
    }
    SafeArrayUnaccessData(aSafeArray.parray);
    return S_OK;
  }

  //----------------------------------------------------------------------------
  //
  HRESULT appendVectorToSafeArray(const VariantVector &aVariantVector, VARIANT &aSafeArray)
  {
    CComSafeArray<VARIANT> original;
    HRESULT hr = S_OK;
    if (aSafeArray.vt != VT_SAFEARRAY) {
      hr = VariantClear(&aSafeArray);
      if (FAILED(hr)) {
        return hr;
      }
      hr = original.Create();
      if (FAILED(hr)) {
        return hr;
      }
    } else {
      original.Attach(aSafeArray.parray);
      aSafeArray.vt = VT_EMPTY;
    }

    size_t originalSize = original.GetCount();
    original.Resize(originalSize + aVariantVector.size());
    for (size_t i = 0; i < aVariantVector.size(); ++i) {
      VARIANT tmp = {0};
      VariantCopy(&tmp, &(aVariantVector[i]));
      original.SetAt(originalSize + i, tmp);
    }
    CComVariant var(original.Detach());
    var.Detach(&aSafeArray);
    return S_OK;
  }

  //----------------------------------------------------------------------------
  //
  HRESULT removeUrlFragment(BSTR aUrl, BSTR* aRet) {
    ATLASSERT(aRet);
    *aRet = ::SysAllocString(aUrl);
    wchar_t* fragment = wcschr(*aRet, L'#');
    if (fragment) {
      *fragment = L'\0';
    }
    return S_OK;
  }

  //----------------------------------------------------------------------------
  //
  BOOL CALLBACK EnumBrowserWindows(HWND hwnd, LPARAM lParam)
  {
    wchar_t className[MAX_PATH];
    ::GetClassName(hwnd, className, MAX_PATH);
    if (wcscmp(className, L"Internet Explorer_Server") == 0) {
      // Now we need to get the IWebBrowser2 from the window.
      DWORD dwMsg = ::RegisterWindowMessage(L"WM_HTML_GETOBJECT");
      LRESULT lResult = 0;
      ::SendMessageTimeout(hwnd, dwMsg, 0, 0, SMTO_ABORTIFHUNG, 1000, (DWORD*) &lResult);
      if (lResult) {
        CComPtr<IHTMLDocument2> doc;
        HRESULT hr = ::ObjectFromLresult(lResult, IID_IHTMLDocument2, 0, (void**) &doc);
        if (SUCCEEDED(hr)) {
          CComPtr<IHTMLWindow2> win;
          hr = doc->get_parentWindow(&win);
          if (SUCCEEDED(hr)) {
            CComQIPtr<IServiceProvider> sp(win);
            CComPtr<IWebBrowser2> pWebBrowser;
            if (sp) {
              hr = sp->QueryService(IID_IWebBrowserApp, IID_IWebBrowser2, (void**) &pWebBrowser);
              if (SUCCEEDED(hr)) {
                CComPtr<IDispatch> container;
                pWebBrowser->get_Container(&container);
                // IWebBrowser2 doesn't have a container if it is an IE tab, so if we have a container
                // then we must be an embedded web browser (e.g. in an HTML toolbar).
                if (!container) {
                  // Now get the HWND associated with the tab so we can see if it is active.
                  sp = pWebBrowser;
                  if (sp) {
                    CComPtr<IOleWindow> oleWindow;
                    hr = sp->QueryService(SID_SShellBrowser, IID_IOleWindow, (void**) &oleWindow);
                    if (SUCCEEDED(hr)) {
                      HWND hTab;
                      hr = oleWindow->GetWindow(&hTab);
                      if (SUCCEEDED(hr) && ::IsWindowEnabled(hTab)) {
                        // Success, we found the active browser!
                        pWebBrowser.CopyTo((IWebBrowser2 **) lParam);
                        return FALSE;
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
    return TRUE;
  }

  ///////////////////////////////////////////////////////////
  // CIDispatchHelper
  // Put a property
  HRESULT CIDispatchHelper::SetProperty(LPOLESTR lpsName, CComVariant vtValue)
  {
    ATLASSERT(p);
    CComQIPtr<IDispatchEx> scriptDispatch(p);
    if (!scriptDispatch)
    {
      return E_UNEXPECTED;
    }
	  DISPID did = 0;
	  HRESULT hr = scriptDispatch->GetDispID(CComBSTR(lpsName), fdexNameEnsure, &did);
	  if (FAILED(hr))
		  return hr;
	  DISPID namedArgs[] = {DISPID_PROPERTYPUT};
	  DISPPARAMS params = {&vtValue, namedArgs, 1, 1};
	  return scriptDispatch->InvokeEx(did, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &params, NULL, NULL, NULL);
  }

  HRESULT CIDispatchHelper::SetPropertyByRef(LPOLESTR lpsName, CComVariant vtValue)
  {
    ATLASSERT(p);
    CComQIPtr<IDispatchEx> scriptDispatch(p);
    if (!scriptDispatch)
    {
      return E_UNEXPECTED;
    }
	  DISPID did = 0;
	  HRESULT hr = scriptDispatch->GetDispID(CComBSTR(lpsName), fdexNameEnsure, &did);
	  if (FAILED(hr))
		  return hr;
	  DISPID namedArgs[] = {DISPID_PROPERTYPUT};
	  DISPPARAMS params = {&vtValue, namedArgs, 1, 1};
	  return scriptDispatch->InvokeEx(did, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUTREF, &params, NULL, NULL, NULL);
  }

  HRESULT CIDispatchHelper::Call(LPOLESTR lpsName, DISPPARAMS* pParams, CComVariant* pvtRet)
  {
    ATLASSERT(p);
    DISPID did = 0;
    if (lpsName)
    {
      LPOLESTR lpNames[] = {lpsName};
      HRESULT hr = p->GetIDsOfNames(IID_NULL, lpNames, 1, LOCALE_USER_DEFAULT, &did);
      if (FAILED(hr))
        return hr;
    }

    DISPPARAMS params = {0};
    CComVariant vtResult;
    if (!pParams)
      pParams = &params;
    if (!pvtRet)
      pvtRet = &vtResult;
    return p->Invoke(did, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, pParams, pvtRet, NULL, NULL);
  }

#ifndef _NO_IWEBBROWSER
  CIDispatchHelper CIDispatchHelper::GetScriptDispatch(IWebBrowser2* pBrowser)
  {
    ATLASSERT(pBrowser);
    CComQIPtr<IDispatch> spRet;
    CComPtr<IDispatch> spDoc;
    HRESULT hr = pBrowser->get_Document(&spDoc);
    if (FAILED(hr))
      return spRet;

    CComQIPtr<IHTMLDocument2> spHTMLDoc(spDoc);
    if (!spHTMLDoc)
      return spRet;

    CComPtr<IHTMLWindow2> spWindow;
    hr = spHTMLDoc->get_parentWindow(&spWindow);
    if (FAILED(hr))
      return spRet;
    spRet = spWindow;
    return spRet;
  }
#endif


  HRESULT CIDispatchHelper::Call_exOnCloseWindow()
  {
    return Call(L"exOnCloseWindow");
  }

  HRESULT CIDispatchHelper::Call_exOnBeforeNavigate(LPDISPATCH pDisp, VARIANT *pURL)
  {
    // arguments in reversed order!
    CComVariant args[2] = {*pURL, pDisp};
    DISPPARAMS params = {args, NULL, 2, 0};
    return Call(L"exOnBeforeNavigate", &params);
  }

  HRESULT CIDispatchHelper::Call_exOnDocumentComplete(LPDISPATCH pDisp, VARIANT *pURL)
  {
    // arguments in reversed order!
    CComVariant args[2] = {*pURL, pDisp};
    DISPPARAMS params = {args, NULL, 2, 0};
    return Call(L"exOnDocumentComplete", &params);
  }

  HRESULT CIDispatchHelper::Call_exOnShowWindow()
  {
    return Call(L"exOnShowWindow");
  }

  HRESULT CIDispatchHelper::Call_exOnHideWindow()
  {
    return Call(L"exOnHideWindow");
  }

  BOOL CIDispatchHelper::Call_exOnDeactivateWindow()
  {
    // returns TRUE if callback returns true. FALSE otherwise.
    CComVariant vt;
    HRESULT hr = Call(L"exOnDeactivateWindow", NULL, &vt);
    if (FAILED(hr))
      return FALSE;
    vt.ChangeType(VT_BOOL);
    return (VARIANT_TRUE == vt.boolVal);
  }

  HRESULT CIDispatchHelper::CreateObject(LPOLESTR lpsName, LPDISPATCH* ppDisp, DISPPARAMS* pConstructorParams)
  {
    ATLASSERT(p);
    ATLASSERT(ppDisp);

    // get DISPID for lpsName
    DISPID did = 0;
    LPOLESTR lpNames[] = {lpsName};
    HRESULT hr = p->GetIDsOfNames(IID_NULL, lpNames, 1, LOCALE_USER_DEFAULT, &did);
    if (FAILED(hr))
      return hr;

    // invoke scriptdispatch with DISPATCH_PROPERTYGET for lpsName
    CComVariant vtRet;
    DISPPARAMS params = {0};
    CComVariant vtResult;
    hr = p->Invoke(did, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params, &vtResult, NULL, NULL);
    if (FAILED(hr))
      return hr;

    // check result: should be a VT_DISPATCH
    if ((VT_DISPATCH != vtResult.vt) || (NULL == vtResult.pdispVal))
      return DISP_E_TYPEMISMATCH;

    // get IDispatchEx on returned IDispatch
    CComQIPtr<IDispatchEx> prototype(vtResult.pdispVal);
    if (!prototype)
      return E_NOINTERFACE;

    // call InvokeEx with DISPID_VALUE and DISPATCH_CONSTRUCT to construct new object
    vtResult.Clear();
    if (!pConstructorParams)
      pConstructorParams = &params;
    hr = prototype->InvokeEx(DISPID_VALUE, LOCALE_USER_DEFAULT, DISPATCH_CONSTRUCT, pConstructorParams, &vtResult, NULL, NULL);
    if (FAILED(hr))
      return hr;

    // vtresult should contain the new object now.
    if ((VT_DISPATCH != vtResult.vt) || (NULL == vtResult.pdispVal))
      return DISP_E_TYPEMISMATCH;

    return vtResult.pdispVal->QueryInterface(IID_IDispatch, (void**)ppDisp);
  }

}// namespace LIB_BhoHelper
