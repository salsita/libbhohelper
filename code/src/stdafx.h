// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include <SDKDDKVer.h>

#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit

#include "gtest/gtest.h"
#include "gmock/gmock.h"

#include <atlbase.h>
#include <atlstr.h>
#include <atlapp.h>

#include <atlcom.h>
#include <atltypes.h>
#include <atlwin.h>
#include <atlframe.h>
#include <atlctl.h>
#include <atlcoll.h>
#include <atlfile.h>
#include <atlsafe.h>

#include <shlguid.h>
#include <shlobj.h>
#include <exdispid.h>
#include <dispex.h>

using namespace ATL;

