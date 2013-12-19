#include "StdAfx.h"
#include "libbhohelper.h"

namespace LIB_BhoHelper
{
  _ATL_FUNC_INFO info_WebBrowserEvents2_WindowClosing = {CC_STDCALL,VT_EMPTY,2,{VT_BOOL,VT_BOOL|VT_BYREF}};
  _ATL_FUNC_INFO info_WebBrowserEvents2_NewWindow2 = {CC_STDCALL,VT_EMPTY,2,{VT_DISPATCH,VT_BOOL|VT_BYREF}};
  _ATL_FUNC_INFO info_WebBrowserEvents2_BeforeNavigate2 = {CC_STDCALL,VT_EMPTY,7,{VT_DISPATCH,VT_VARIANT|VT_BYREF,VT_VARIANT|VT_BYREF,VT_VARIANT|VT_BYREF,VT_VARIANT|VT_BYREF,VT_VARIANT|VT_BYREF,VT_BOOL|VT_BYREF}};
  _ATL_FUNC_INFO info_WebBrowserEvents2_DocumentComplete = {CC_STDCALL,VT_EMPTY,2,{VT_DISPATCH,VT_VARIANT|VT_BYREF}};
  _ATL_FUNC_INFO info_WebBrowserEvents2_WindowSetHeight = {CC_STDCALL,VT_EMPTY,1,{VT_I4}};
  _ATL_FUNC_INFO info_WebBrowserEvents2_WindowStateChanged = {CC_STDCALL,VT_EMPTY,2,{VT_UI4,VT_UI4}};

  // -------------------------------------------------------------------------
  // GetFuncInfoFromId
  HRESULT GetFuncInfoFromId(
    _In_ ITypeInfo* pTypeInfo, 
    _In_ const IID& /*iid*/, 
    _In_ DISPID dispidMember, 
    _In_ INVOKEKIND aInvokeKind,
    _In_ LCID /*lcid*/, 
    _Inout_ _ATL_FUNC_INFO& info)
  {
    if (pTypeInfo == NULL) {
      return E_INVALIDARG;
    }

    HRESULT hr = S_OK;
    FUNCDESC* pFuncDesc = NULL;
    TYPEATTR* pAttr;
    hr = pTypeInfo->GetTypeAttr(&pAttr);
    if (FAILED(hr)) {
      return hr;
    }
    int i;
    for (i=0;i<pAttr->cFuncs;i++) {
      hr = pTypeInfo->GetFuncDesc(i, &pFuncDesc);
      if (FAILED(hr)) {
        return hr;
      }
      if (pFuncDesc->memid == dispidMember && pFuncDesc->invkind == aInvokeKind) {
        break;
      }
      pTypeInfo->ReleaseFuncDesc(pFuncDesc);
      pFuncDesc = NULL;
    }
    pTypeInfo->ReleaseTypeAttr(pAttr);
    if (pFuncDesc == NULL) {
      return DISP_E_MEMBERNOTFOUND;
    }

    // If this assert occurs, then add a #define _ATL_MAX_VARTYPES nnnn
    // before including atlcom.h
    ATLASSERT(pFuncDesc->cParams <= _ATL_MAX_VARTYPES);
    if (pFuncDesc->cParams > _ATL_MAX_VARTYPES) {
      pTypeInfo->ReleaseFuncDesc(pFuncDesc);
      return E_UNEXPECTED;
    }

    for (i = 0; i < pFuncDesc->cParams; i++) {
      info.pVarTypes[i] = pFuncDesc->lprgelemdescParam[i].tdesc.vt;
      if (info.pVarTypes[i] == VT_PTR) {
        info.pVarTypes[i] = (VARTYPE)(pFuncDesc->lprgelemdescParam[i].tdesc.lptdesc->vt | VT_BYREF);
      }
      if (info.pVarTypes[i] == VT_SAFEARRAY) {
        info.pVarTypes[i] = (VARTYPE)(pFuncDesc->lprgelemdescParam[i].tdesc.lptdesc->vt | VT_ARRAY);
      }
      if (info.pVarTypes[i] == VT_USERDEFINED) {
        info.pVarTypes[i] = AtlGetUserDefinedType(pTypeInfo, pFuncDesc->lprgelemdescParam[i].tdesc.hreftype);
      }
    }

    VARTYPE vtReturn = pFuncDesc->elemdescFunc.tdesc.vt;
    switch(vtReturn) {
      case VT_INT:
        vtReturn = VT_I4;
        break;
      case VT_UINT:
        vtReturn = VT_UI4;
        break;
      case VT_VOID:
        vtReturn = VT_EMPTY; // this is how DispCallFunc() represents void
        break;
      case VT_HRESULT:
        vtReturn = VT_ERROR;
        break;
    }
    info.vtReturn = vtReturn;
    info.cc = pFuncDesc->callconv;
    info.nParams = pFuncDesc->cParams;
    pTypeInfo->ReleaseFuncDesc(pFuncDesc);
    return S_OK;
  }

}// namespace LIB_BhoHelper
