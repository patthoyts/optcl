/*
 *------------------------------------------------------------------------------
 *	utility.cpp
 *	Declares a collection of often used, general purpose functions.
 *	I've also placed the variant/Tcl_Obj conversion functions here. 
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
#ifndef UTILITY_418A3400_56FC_11d3_86E8_0000B482A708
#define UTILITY_418A3400_56FC_11d3_86E8_0000B482A708


#ifndef ASSERT
#	ifdef _DEBUG
#		include <crtdbg.h> 
#		define ASSERT(x) _ASSERTE(x)
#	else
#		define ASSERT(x)
#	endif
#endif



// TRACE functionality - works like printf, only in debug mode 
// - output to the debug console
#ifdef _DEBUG
#	define TRACE OptclTrace
void OptclTrace(LPCTSTR lpszFormat, ...);
#else
#	define TRACE
#endif

#define TCL_CMDEF(fname) int fname (ClientData cd, Tcl_Interp *pInterp, int objc, Tcl_Obj *CONST objv[])
#define CHECKHR(hr) if (FAILED(hr)) throw(hr)
#define CHECKHR_TCL(hr, i, v) if (FAILED(hr)) {Tcl_SetResult (i, HRESULT2Str(hr), TCL_DYNAMIC); return v;}

#define SETDISPPARAMS(dp, numArgs, pvArgs, numNamed, pNamed) \
    {\
    (dp).cArgs=numArgs;\
    (dp).rgvarg=pvArgs;\
    (dp).cNamedArgs=numNamed;\
    (dp).rgdispidNamedArgs=pNamed;\
    }

#define SETNOPARAMS(dp) SETDISPPARAMS(dp, 0, NULL, 0, NULL)

#define _countof(x) (sizeof(x)/sizeof(x[0]))


template <class T> void		delete_ptr (T* &ptr)
{
	if (ptr != NULL) {
		delete ptr;
		ptr = NULL;
	}
}


template <class T> T* delete_array (T *&ptr) {
	if (ptr != NULL) {
		delete []ptr;
		ptr = NULL;
	}
	return ptr;
}



class OptclObj;

bool		var2obj (Tcl_Interp *pInterp, VARIANT &var, TObjPtr &presult, OptclObj **ppObj = NULL);
bool		obj2var_ti (Tcl_Interp *pInterp, TObjPtr &obj, VARIANT &var, ITypeInfo *pInfo, TYPEDESC *pdesc);
bool		obj2var_vt (Tcl_Interp *pInterp, TObjPtr &obj, VARIANT &var, VARTYPE vt);
bool		obj2var_vt_byref (Tcl_Interp *pInterp, TObjPtr &obj, VARIANT &var, VARTYPE vt);
void		obj2var (TObjPtr &obj, VARIANT &var);


void		OptclVariantClear (VARIANT *pvar);


char	*	HRESULT2Str (HRESULT hr);
void		FreeBSTR (BSTR &bstr);
void		FreeBSTRArray (BSTR * bstr, UINT count);
char	*	ExceptInfo2Str (EXCEPINFO *pe);
DISPID		Name2ID (IDispatch *, const LPOLESTR name);
DISPID		Name2ID (IDispatch *, const char *name);
int			ObjectNotFound (Tcl_Interp *pInterp, const char *name);
void		SplitTypedString (char *pstr, char ** ppsecond);
bool		SplitObject (Tcl_Interp *pInterp, Tcl_Obj *pObj, 
						 const char * tokens, Tcl_Obj **ppResult);
bool		SplitBrackets (Tcl_Interp *pInterp, Tcl_Obj *pObj,
						   TObjPtr & result);

/// TESTS
TCL_CMDEF (Obj2VarTest);

#endif // UTILITY_418A3400_56FC_11d3_86E8_0000B482A708