#pragma once

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files
#include <windows.h>

// export symbols for DLL and Specify naming conventions
extern "C" __declspec(dllexport) void __cdecl multiplyBy(double* in, double* out);