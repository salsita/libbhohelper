#pragma once
#include <vector>
#include <map>
#include "Mshtml.h"

// in case Win8 SDK is not installed
#ifndef __IDOMEvent_FWD_DEFINED__
#define __IDOMEvent_FWD_DEFINED__
typedef interface IDOMEvent IDOMEvent;
#endif 	/* __IDOMEvent_FWD_DEFINED__ */

#ifndef __IEventTarget_FWD_DEFINED__
#define __IEventTarget_FWD_DEFINED__
typedef interface IEventTarget IEventTarget;
#endif 	/* __IEventTarget_FWD_DEFINED__ */

#ifndef __IEventTarget_INTERFACE_DEFINED__
#define __IEventTarget_INTERFACE_DEFINED__
/* interface IEventTarget */
/* [object][uuid][dual][oleautomation] */ 

    DEFINE_GUID(IID_IEventTarget, 
        0x305104b9, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b);

    MIDL_INTERFACE("305104b9-98b5-11cf-bb82-00aa00bdce0b")
    IEventTarget : public IDispatch
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE addEventListener( 
            /* [in] */ __RPC__in BSTR type,
            /* [in] */ __RPC__in_opt IDispatch *listener,
            /* [in] */ VARIANT_BOOL useCapture) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE removeEventListener( 
            /* [in] */ __RPC__in BSTR type,
            /* [in] */ __RPC__in_opt IDispatch *listener,
            /* [in] */ VARIANT_BOOL useCapture) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE dispatchEvent( 
            /* [in] */ __RPC__in_opt IDOMEvent *evt,
            /* [out][retval] */ __RPC__out VARIANT_BOOL *pfResult) = 0;
        
    };
#endif 	/* __IEventTarget_INTERFACE_DEFINED__ */

#ifndef __IDOMEvent_INTERFACE_DEFINED__
#define __IDOMEvent_INTERFACE_DEFINED__
/* interface IDOMEvent */
/* [object][uuid][dual][oleautomation] */ 

    DEFINE_GUID(IID_IDOMEvent, 
        0x305104ba, 0x98b5, 0x11cf, 0xbb, 0x82, 0x00, 0xaa, 0x00, 0xbd, 0xce, 0x0b);

    MIDL_INTERFACE("305104ba-98b5-11cf-bb82-00aa00bdce0b")
    IDOMEvent : public IDispatch
    {
    public:
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_bubbles( 
            /* [out][retval] */ __RPC__out VARIANT_BOOL *p) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_cancelable( 
            /* [out][retval] */ __RPC__out VARIANT_BOOL *p) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_currentTarget( 
            /* [out][retval] */ __RPC__deref_out_opt IEventTarget **p) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_defaultPrevented( 
            /* [out][retval] */ __RPC__out VARIANT_BOOL *p) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_eventPhase( 
            /* [out][retval] */ __RPC__out USHORT *p) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_target( 
            /* [out][retval] */ __RPC__deref_out_opt IEventTarget **p) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_timeStamp( 
            /* [out][retval] */ __RPC__out ULONGLONG *p) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_type( 
            /* [out][retval] */ __RPC__deref_out_opt BSTR *p) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE initEvent( 
            /* [in] */ __RPC__in BSTR eventType,
            /* [in] */ VARIANT_BOOL canBubble,
            /* [in] */ VARIANT_BOOL cancelable) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE preventDefault( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE stopPropagation( void) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE stopImmediatePropagation( void) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_isTrusted( 
            /* [out][retval] */ __RPC__out VARIANT_BOOL *p) = 0;
        
        virtual /* [id][propput] */ HRESULT STDMETHODCALLTYPE put_cancelBubble( 
            /* [in] */ VARIANT_BOOL v) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_cancelBubble( 
            /* [out][retval] */ __RPC__out VARIANT_BOOL *p) = 0;
        
        virtual /* [id][propget] */ HRESULT STDMETHODCALLTYPE get_srcElement( 
            /* [out][retval] */ __RPC__deref_out_opt IHTMLElement **p) = 0;
        
    };
#endif 	/* __IDOMEvent_INTERFACE_DEFINED__ */



namespace LIB_BhoHelper
{

  ///////////////////////////////////////////////////////////
  // ATL function infos for WebBrowserEvents2
  extern _ATL_FUNC_INFO info_WebBrowserEvents2_WindowClosing;
  extern _ATL_FUNC_INFO info_WebBrowserEvents2_NewWindow2;
  extern _ATL_FUNC_INFO info_WebBrowserEvents2_BeforeNavigate2;
  extern _ATL_FUNC_INFO info_WebBrowserEvents2_DocumentComplete;
  extern _ATL_FUNC_INFO info_WebBrowserEvents2_WindowSetHeight;
  extern _ATL_FUNC_INFO info_WebBrowserEvents2_WindowStateChanged;

  ///////////////////////////////////////////////////////////
  // Utility functions
  HRESULT createIDispatchFromCreator(LPDISPATCH aCreator, VARIANT* aRet);

  typedef std::map<std::wstring, CComPtr<IDispatch>> DispatchMap;

  typedef std::vector<CComVariant> VariantVector;


  HRESULT addJSArrayToVariantVector(LPDISPATCH aArrayDispatch, VariantVector &aVariantVector, bool aReverse = false);
  HRESULT constructSafeArrayFromVector(const VariantVector &aVariantVector, VARIANT &aSafeArray);
  HRESULT appendVectorToSafeArray(const VariantVector &aVariantVector, VARIANT &aSafeArray);

  HRESULT getHTMLDocumentForHWND(HWND hwnd, IHTMLDocument2** aRet);
  HRESULT getBrowserForHTMLDocument(IHTMLDocument2* aDocument, IWebBrowser2** aRet);

  HRESULT removeUrlFragment(BSTR aUrl, BSTR* aRet);
  BOOL CALLBACK EnumBrowserWindows(HWND hwnd, LPARAM lParam);

  ///////////////////////////////////////////////////////////
  // CIDispatchHelper
  // makes IDispatch handling easier
  class CIDispatchHelper : public CComQIPtr<IDispatch>
  {
  public:
    // Methods overwritten from CComQIPtr<IDispatch>
    CIDispatchHelper() throw() :
        CComQIPtr<IDispatch>() {}
    CIDispatchHelper(IDispatch* lp) throw() :
        CComQIPtr<IDispatch>(lp) {}
    CIDispatchHelper(const CComQIPtr<IDispatch>& p) throw() :
        CComQIPtr<IDispatch>(p) {}
    IDispatch* operator=(IDispatch* lp) throw()
        {return CComQIPtr<IDispatch>::operator =(lp);};

    // Get a property from IDispatch.
    // T is the type of the return value
    // VT is the resulting VARIANT type. Get(..) trys to do a conversion to this type.
    // TCast is an optional type to cast the result to before assigning to the return value.
    // This is required e.g. for BSTR
    template <class T, VARTYPE VT, class TCast> HRESULT Get(LPOLESTR lpsName, T &Ret)
    {
      ATLASSERT(p);
      DISPID did = 0;
      LPOLESTR lpNames[] = {lpsName};
      HRESULT hr = p->GetIDsOfNames(IID_NULL, lpNames, 1, LOCALE_USER_DEFAULT, &did);
      if (FAILED(hr))
        return hr;

      DISPPARAMS params = {0};
      CComVariant vtResult;
      hr = p->Invoke(did, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params, &vtResult, NULL, NULL);
      if (FAILED(hr))
        return hr;
      hr = vtResult.ChangeType(VT);
      if (FAILED(hr))
        return hr;
      Ret = (T)(TCast)(vtResult.llVal);
      return S_OK;
    }

    // Put a property
    HRESULT SetProperty(LPOLESTR lpsName, CComVariant vtValue);
    HRESULT SetPropertyByRef(LPOLESTR lpsName, CComVariant vtValue);

    // Wrapper used if T == TCast
    template <class T, VARTYPE VT> HRESULT Get(LPOLESTR lpsName, T &Ret)
      {return Get<T, VT, T>(lpsName, Ret);}

    // Call a method. If lpsName is NULL method with DISPID 0 is called (used for js function objects)
    HRESULT Call(LPOLESTR lpsName = NULL, DISPPARAMS* pParams = NULL, CComVariant* pvtRet = NULL);

    // create a JS object
    HRESULT CreateObject(LPOLESTR lpsName, LPDISPATCH* ppDisp, DISPPARAMS* pConstructorParams = NULL);

#ifndef _NO_IWEBBROWSER
    // Get the IDispatch for javascript from a WebBrowserControl
    static CIDispatchHelper GetScriptDispatch(IWebBrowser2* pBrowser);
#endif
    // some methods for WebBrowserEvents2
    HRESULT Call_exOnCloseWindow();
    HRESULT Call_exOnBeforeNavigate(LPDISPATCH pDisp, VARIANT *pURL);
    HRESULT Call_exOnDocumentComplete(LPDISPATCH pDisp, VARIANT *pURL);

    // and often used ones
    HRESULT Call_exOnShowWindow();
    HRESULT Call_exOnHideWindow();
    BOOL Call_exOnDeactivateWindow();
  };

  ///////////////////////////////////////////////////////////
  // DOMEventHandlerAdapter
  // Adapter for DOM event handlers.
  // Uses addEventListener or "on.." property of mEventTarget
  // depending on what is available. An event added via
  // addEventListener gets removed in remove() and destructor.
  // When the "on.." method is used it overwrites existing
  // handlers, even if it has already attached to this event.
  class DOMEventHandlerAdapter
  {
  public:
    DOMEventHandlerAdapter() : mCapture(VARIANT_FALSE)
    {
    };

    ~DOMEventHandlerAdapter()
    {
      remove();
    };

    bool equals(IDispatch * aEventTarget, LPCOLESTR aEventName, IDispatch * aHandler, BOOL aCapture = FALSE)
    {
      return( mHandler == aHandler
              && mEventTarget.IsEqualObject(aEventTarget)
              && mEventName == aEventName
              && mCapture == aCapture);
    };

    HRESULT addTo(IDispatch * aEventTarget, LPCOLESTR aEventName, IDispatch * aHandler, BOOL aCapture = FALSE)
    {
      ATLASSERT(aEventTarget);
      ATLASSERT(aHandler);
      if (!aEventTarget || !aHandler) {
        return E_INVALIDARG;
      }

      if ( equals(aEventTarget, aEventName, aHandler, aCapture) && mEventTarget ) {
        // already attached a handler for this via addEventListener, so no need to do it again
        return S_OK;  
      }

      // In any case we add the event now. Whether it was done before using "on..." or
      // it was not done at all. Means, if "on.." is used it will overwrite the previous
      // handler.
      // Clean up first
      remove();

      mEventName = aEventName;
      mHandler = aHandler;
      mCapture = (aCapture) ? VARIANT_TRUE : VARIANT_FALSE;

      // see if we can get a IEventTarget
      mEventTarget = aEventTarget;
      if (mEventTarget) {
        // yes, so use addEventListener / removeEventListener
        return mEventTarget->addEventListener(mEventName, aHandler, mCapture);
      }

      // IEventTarget not supported, set via "on..." property
      CIDispatchHelper eventTarget(aEventTarget);
      ATLASSERT(eventTarget);
      CComBSTR name(L"on");
      name.Append(mEventName);
      return eventTarget.SetPropertyByRef(name, CComVariant(aHandler));
    }

    void remove()
    {
      // if IEventTarget is supported call removeEventListener
      if (mEventTarget) {
        mEventTarget->removeEventListener(mEventName, mHandler, mCapture);
        mEventTarget.Release();
      }
      mHandler.Release();
    }

  public:
    CComQIPtr<IEventTarget> mEventTarget;
    CComPtr<IDispatch> mHandler;
    CComBSTR mEventName;
    VARIANT_BOOL mCapture;
  };

  ///////////////////////////////////////////////////////////
  // ComLibrary
  // Allows in-proc server loading without the server being
  // registered.
  // It basically wraps DllGetClassObject(),
  // IClassFactory::CreateInstance() and DllCanUnloadNow().
  //
  // Usage:
  // Put a member somwhere:
  //   ComLibrary mMyLib;
  // Before you want to use it call:
  //   HRESULT hr = mMyLib.load(L"MyLib.dll");
  // And when you finished:
  //   mMyLib.unload();
  // If you use the library from within another library it's
  // probably a good idea to modify your DllCanUnloadNow()
  // and check there also your mMyLib.
  class ComLibrary
  {
  // -------------------------------------------------------------------------
  public: // members
    ComLibrary();
    virtual ~ComLibrary();

    HRESULT load(LPOLESTR szDllName, HINSTANCE aHinstanceRel = NULL);
    HRESULT unload(BOOL bForce = FALSE);

    // these two wrap the functions exported from the COM module
    HRESULT DllGetClassObject(REFCLSID aClsid, 
      REFIID aIid, LPVOID FAR* ppv);
    HRESULT DllCanUnloadNow();

    // wraps DllGetClassObject and IClassFactory::CreateInstance
    template<class Iface> HRESULT createInstance(
        REFCLSID aClsid,
        Iface** aRetPtrPtr,
        IUnknown* aUnknownOuter = NULL)
    {
      if (!mDllGetClassObject)
        { return E_UNEXPECTED; } // missing DllGetClassObject

      // get class factory
      CComPtr<IClassFactory> classFactory;
      HRESULT hr = this->DllGetClassObject(aClsid, IID_IClassFactory, (LPVOID*)&classFactory.p);
      if (FAILED(hr))
        { return hr; }
      return classFactory->CreateInstance(aUnknownOuter, __uuidof(Iface), (void**)aRetPtrPtr);
    }

  private:
    // type for DllGetClassObject
    typedef HRESULT (__stdcall *FnDllGetClassObject)(IN REFCLSID rclsid, 
      IN REFIID riid, OUT LPVOID FAR* ppv);
    // type for DllCanUnloadNow
    typedef HRESULT (__stdcall *FnDllCanUnloadNow)(void);

    // module handle
    HMODULE mhDll;
    // DllGetClassObject
    FnDllGetClassObject mDllGetClassObject;
    // DllCanUnloadNow
    FnDllCanUnloadNow   mDllCanUnloadNow;
  };

  // -------------------------------------------------------------------------
  // GetFuncInfoFromId
  // This is a modified copy of ALT::AtlGetFuncInfoFromId(..).
  // It allows also to get a function info for every INVOKEKIND, means,
  // also for PROPGET, PROPPUT etc.
  HRESULT GetFuncInfoFromId(
	  _In_ ITypeInfo* pTypeInfo, 
	  _In_ const IID& /*iid*/, 
	  _In_ DISPID dispidMember, 
    _In_ INVOKEKIND aInvokeKind,
	  _In_ LCID /*lcid*/, 
	  _Inout_ _ATL_FUNC_INFO& info);

  // DISPATCH_CALL_ENTRY
  // Copy of _ATL_EVENT_ENTRY with an additional INVOKEKIND member.
  // Used by the DISPCALL-map macros.
  template <class T>
    struct DISPATCH_CALL_ENTRY
  {
	  UINT nControlID;			//ID identifying object instance
	  const IID* piid;			//dispinterface IID
	  int nOffset;				  //offset of dispinterface from this pointer
	  DISPID dispid;				//DISPID of method/property
    INVOKEKIND invokeKind;
	  void (__stdcall T::*pfn)();	//method to invoke
	  _ATL_FUNC_INFO* pInfo;
  };

  // These macros work similar like the sink map macros (e.g. BEGIN_SINK_MAP(..) ).
  // The difference is, that it is possible to handle also PROPERTY_GET and
  // PROPERTY_PUT(REF) calls also.
#define BEGIN_DISPCALL_MAP(_class)\
	typedef _class _GetFuncMapFinder;\
	static const LIB_BhoHelper::DISPATCH_CALL_ENTRY<_class>* _GetFuncMap()\
	{\
		PTM_WARNING_DISABLE \
		typedef _class _disp_classtype;\
		static const LIB_BhoHelper::DISPATCH_CALL_ENTRY<_class> map[] = {

#define DISPCALL_MAP_ENTRY(id, iid, dispid, invokeKind, fn) \
  {id, &iid, (int)(INT_PTR)((_disp_classtype*)8)-8, dispid, invokeKind, (void (__stdcall _disp_classtype::*)())fn, NULL},

#define METHOD_ENTRY(id, iid, dispid, fn) \
  DISPCALL_MAP_ENTRY(id, iid, dispid, INVOKE_FUNC, fn)

#define PROPGET_ENTRY(id, iid, dispid, fn) \
  DISPCALL_MAP_ENTRY(id, iid, dispid, INVOKE_PROPERTYGET, fn)

#define PROPPUT_ENTRY(id, iid, dispid, fn) \
  DISPCALL_MAP_ENTRY(id, iid, dispid, INVOKE_PROPERTYPUT, fn)

#define PROPPUTREF_ENTRY(id, iid, dispid, fn) \
  DISPCALL_MAP_ENTRY(id, iid, dispid, INVOKE_PROPERTYPUTREF, fn)

#define PROPERTY_ENTRY(id, iid, dispid, fnGet, fnPut) \
  DISPCALL_MAP_ENTRY(id, iid, dispid, INVOKE_PROPERTYGET, fnGet) \
  DISPCALL_MAP_ENTRY(id, iid, dispid, INVOKE_PROPERTYPUT, fnPut)

#define PROPERTY_ENTRYREF(id, iid, dispid, fnGet, fnPut) \
  DISPCALL_MAP_ENTRY(id, iid, dispid, INVOKE_PROPERTYGET, fnGet) \
  DISPCALL_MAP_ENTRY(id, iid, dispid, INVOKE_PROPERTYPUTREF, fnPut)

#define END_DISPCALL_MAP() \
	  {0, NULL, 0, 0, INVOKE_FUNC, NULL, NULL} }; \
    return map;\
	  PTM_WARNING_RESTORE \
	}

  ///////////////////////////////////////////////////////////
  // template class DispinterfaceImpl
  // Derives from IDispatchImpl.
  // Allows (partial) impementation of dispinterfaces
  // (non-dual IDispatch interfaces).
  // Basically this class does the same like IDispEventImpl.
  // The methods implemented here are mainly copies of 
  // IDispEventImpl methods - IDispEventImpl is able to handle
  // partial implementations of dispinterface based interfaces.
  // IDispEventImpl though allows only methods, no properties.
  template <UINT nID, class T, class Tif, const IID* piid = &__uuidof(Tif),
      const GUID* plibid = &CAtlModule::m_libid, WORD wMajor = 1,
      WORD wMinor = 0, class tihclass = CComTypeInfoHolder>
        class ATL_NO_VTABLE DispinterfaceImpl :
          public IDispatchImpl<Tif, piid, plibid, wMajor, wMinor, tihclass>
  {
  public:
	  STDMETHOD(Invoke)(
		  _In_ DISPID dispidMember, 
		  _In_ REFIID /*riid*/,
		  _In_ LCID lcid, 
		  _In_ WORD wFlags, 
		  _In_ DISPPARAMS* pdispparams, 
		  _Out_opt_ VARIANT* pvarResult,
		  _In_opt_ EXCEPINFO* /*pexcepinfo*/, 
		  _In_opt_ UINT* /*puArgErr*/)
	  {
      INVOKEKIND aInvokeKind = INVOKE_FUNC;
      if (wFlags & DISPATCH_PROPERTYGET) {
        aInvokeKind = INVOKE_PROPERTYGET;
      }
      else if (wFlags & DISPATCH_PROPERTYPUT) {
        aInvokeKind = INVOKE_PROPERTYPUT;
      }
		  const LIB_BhoHelper::DISPATCH_CALL_ENTRY<T>* pMap = T::_GetFuncMap();
		  const LIB_BhoHelper::DISPATCH_CALL_ENTRY<T>* pFound = NULL;
		  while (pMap->piid != NULL) {
			  if ((pMap->nControlID == nID) &&
            (pMap->dispid == dispidMember) && 
				    (IsEqualIID(*(pMap->piid), *piid)) &&
            (pMap->invokeKind == aInvokeKind))
			  {
				  pFound = pMap;
				  break;
			  }
			  pMap++;
		  }
      if (pFound == NULL) {
			  return S_OK;
      }

		  _ATL_FUNC_INFO info;
		  _ATL_FUNC_INFO* pInfo;
      if (pFound->pInfo != NULL) {
			  pInfo = pFound->pInfo;
      }
		  else {
			  pInfo = &info;
			  HRESULT hr = GetFuncInfoFromId(*piid, dispidMember, aInvokeKind, lcid, info);
			  if (FAILED(hr))
				  return S_OK;
		  }
		  return InvokeFromFuncInfo(pFound->pfn, *pInfo, pdispparams, pvarResult);
	  }

	  //Helper for invoking the event
	  HRESULT InvokeFromFuncInfo(
		  /* _In_ */ void (__stdcall T::*pEvent)(), 
		  _In_ _ATL_FUNC_INFO& info, 
		  _In_ DISPPARAMS* pdispparams, 
		  _Out_opt_ VARIANT* pvarResult)
	  {
		  ATLASSERT(pdispparams->cArgs == (UINT)info.nParams);

		  T* pT = static_cast<T*>(this);

		  // If this assert occurs, then add 
		  // #define _ATL_MAX_VARTYPES nnnn
		  // before including atlcom.h
		  ATLASSERT(info.nParams <= _ATL_MAX_VARTYPES);
		  if (info.nParams > _ATL_MAX_VARTYPES) {
			  return E_FAIL;
		  }
		  VARIANTARG* rgVarArgs[_ATL_MAX_VARTYPES];
		  VARIANTARG** pVarArgs = info.nParams ? rgVarArgs : 0;

		  UINT nIndex = 0;

#ifndef _ATL_IGNORE_NAMED_ARGS
		  for (nIndex; nIndex < pdispparams->cNamedArgs; nIndex++) {
			  ATLASSERT( ( NULL != pVarArgs ) && ( pdispparams->rgdispidNamedArgs[nIndex] < _countof(rgVarArgs) ) );
			  if( ( NULL == pVarArgs ) || ( pdispparams->rgdispidNamedArgs[nIndex] >= _countof(rgVarArgs) ) ) {
				  return E_FAIL;
			  }
			  pVarArgs[pdispparams->rgdispidNamedArgs[nIndex]] = &pdispparams->rgvarg[nIndex];
		  }
#endif

		  for (; nIndex < pdispparams->cArgs; nIndex++) {
			  ATLASSERT( NULL != pVarArgs );
			  if( NULL == pVarArgs )
			  {
				  return E_FAIL;
			  }
			  pVarArgs[info.nParams-nIndex-1] = &pdispparams->rgvarg[nIndex];
		  }

		  CComStdCallThunk<T> thunk;
		  thunk.Init(pEvent, pT);

		  CComVariant tmpResult;
		  if (pvarResult == NULL)
			  pvarResult = &tmpResult;

		  HRESULT hr = DispCallFunc(
			  &thunk,
			  0,
			  info.cc,
			  info.vtReturn,
			  info.nParams,
			  info.pVarTypes,
			  pVarArgs,
			  pvarResult);
		  ATLASSERT(SUCCEEDED(hr));
		  return hr;
	  }

	  //Helper for finding the function index for a DISPID
	  HRESULT GetFuncInfoFromId(
		  _In_ const IID& iid, 
		  _In_ DISPID dispidMember, 
      _In_ INVOKEKIND aInvokeKind,
		  _In_ LCID lcid, 
		  _Inout_ _ATL_FUNC_INFO& info)
	  {
		  CComPtr<ITypeInfo> spTypeInfo;
      HRESULT hr = _tih.GetTI(lcid, &spTypeInfo);
      if (FAILED(hr)) {
			  return hr;
      }
		  return ::LIB_BhoHelper::GetFuncInfoFromId(spTypeInfo, iid, dispidMember, aInvokeKind, lcid, info);
	  }
  };

}// namespace LIB_BhoHelper
