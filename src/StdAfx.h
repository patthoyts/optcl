/*
 *------------------------------------------------------------------------------
 *	stdafx.cpp
 *	include file for standard system include files, or project specific 
 *	include files that are used frequently, but are changed infrequently
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


// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#if !defined(AFX_STDAFX_H__1363E007_C12C_11D2_8003_0040055861F2__INCLUDED_)
#define AFX_STDAFX_H__1363E007_C12C_11D2_8003_0040055861F2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


// Insert your headers here
//#define WIN32_LEAN_AND_MEAN		// Exclude rarely-used stuff from Windows headers

#include <Atlbase.h>
extern CComModule _Module;
#include <atlcom.h>
#include <atlhost.h>
#include <atlwin.h>

#include <windows.h>
#include <comdef.h>
#include <tcl.h>
#include <tk.h>

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__1363E007_C12C_11D2_8003_0040055861F2__INCLUDED_)
