#pragma once

// Including SDKDDKVer.h defines the highest available Windows platform.

// If you wish to build your application for a previous Windows platform, include WinSDKVer.h and
// set the _WIN32_WINNT macro to the platform you wish to support before including SDKDDKVer.h.

#include <WinSDKVer.h>

#ifndef _WIN32_WINNT
// _WIN32_WINNT_VISTA - XP SP3
# define _WIN32_WINNT  0x0502
#endif

#ifndef _WIN32_IE
// _WIN32_IE_IE80 - IE8
# define _WIN32_IE     0x0800
#endif

#include <SDKDDKVer.h>
