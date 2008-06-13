/*
 *------------------------------------------------------------------------------
 *	optcl.cpp
 *	Declares the OpTcl's main entry points.
 *
 *	Copyright (C) 1999  Farzad Pezeshkpour, University of East Anglia
 *
 *	This program is free software; you can redistribute it and/or
 *	modify it under the terms of the GNU General Public License
 *	as published by the Free Software Foundation; either version 2
 *	of the License, or (at your option) any later version.
 *
 *	This program is distributed in the hope that it will be useful,
 *	but WITHOUT ANY WARRANTY; without even the implied warranty of
 *	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *	GNU General Public License for more details.
 *
 *	You should have received a copy of the GNU General Public License
 *	along with this program; if not, write to the Free Software
 *	Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
 *------------------------------------------------------------------------------
 */

#ifndef _OPTCL_H_B229D2A0_616A_11d4_8004_0040055861F2
#define _OPTCL_H_B229D2A0_616A_11d4_8004_0040055861F2



// debugging symbols
#ifdef _DEBUG
#define _ATL_DEBUG_INTERFACES	
#define _ATL_DEBUG_REFCOUNT		
#define _ATL_DEBUG_QI
#endif



extern "C" DLLEXPORT int Optcl_Init (Tcl_Interp *pInterp);
extern "C" DLLEXPORT BOOL WINAPI DllMain (HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved);
int		TypeLib_Init (Tcl_Interp *pInterp);


extern CComPtr<IMalloc> g_pmalloc;
extern bool g_bTkInit;
#endif// _OPTCL_H_B229D2A0_616A_11d4_8004_0040055861F2