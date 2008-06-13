/*
 *------------------------------------------------------------------------------
 *	utility.cpp
 *	Implements a collection of often used, general purpose functions.
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


#include "stdafx.h"
#include "tbase.h"
#include "utility.h"
#include "optcl.h"
#include "objmap.h"
#include "typelib.h"
#include "optclobj.h"
#include "optcltypeattr.h"

#ifdef _DEBUG
/*
 *-------------------------------------------------------------------------
 * OptclTrace --
 *	Performs a debugging service similar to printf. Works only under debug.
 *	#defined to TRACE(formatstring, ....)
 * Result:
 *	None.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void OptclTrace(LPCTSTR lpszFormat, ...)
{
	va_list args;
	va_start(args, lpszFormat);

	int nBuf;
	TCHAR szBuffer[512];

	nBuf = _vsntprintf(szBuffer, _countof(szBuffer), lpszFormat, args);

	// was there an error? was the expanded string too long?
	ASSERT(nBuf >= 0);

	OutputDebugString (szBuffer);
	va_end(args);
}
#endif //_DEBUG


/*
 *-------------------------------------------------------------------------
 * HRESULT2Str --
 *	Converts an HRESULT to a Tcl allocated string.
 *
 * Result:
 *	The string if not null.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
char * HRESULT2Str (HRESULT hr)
{

	USES_CONVERSION;

    LPTSTR   szMessage;
	char	*message;
	char	*tclmessage;
	
    if (HRESULT_FACILITY(hr) == FACILITY_WINDOWS) 
		hr = HRESULT_CODE(hr); 
	
    FormatMessage( 
		FORMAT_MESSAGE_ALLOCATE_BUFFER | 
		FORMAT_MESSAGE_FROM_SYSTEM, 
		NULL, 
		hr, 
		MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), //The user default language 
		(LPTSTR)&szMessage, 
		0, 
		NULL ); 
	
	// conversion to char * if unicode
	message = T2A (szMessage);
	tclmessage = Tcl_Alloc(strlen(message)+1);
	strcpy(tclmessage, message);
	for (char *i = tclmessage; *i != 0; i++)
		if (*i == '\r') *i = ' ';
    LocalFree(szMessage);
	return tclmessage;
}




/*
 *-------------------------------------------------------------------------
 * FreeBSTR --
 *	If not NULL, releases the BSTR and sets it to NULL.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void FreeBSTR (BSTR &bstr)
{
	if (bstr != NULL) {
		SysFreeString (bstr);
		bstr = NULL;
	}
}


/*
 *-------------------------------------------------------------------------
 * FreeBSTRArray --
 *	Releases an array of BSTR and sets them to NULL, if not already.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void FreeBSTRArray (BSTR * bstr, UINT count)
{
	if (bstr == NULL) return;
	for (UINT index = 0; index < count; index++)
	{
		FreeBSTR (bstr[index]);
	}
}



/*
 *-------------------------------------------------------------------------
 * ExceptInfo2Str --
 *	Converts an EXCEPINFO structure to a Tcl allocated string.
 *
 * Result:
 *	The string if not null.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
char * ExceptInfo2Str (EXCEPINFO *pe)
{
	USES_CONVERSION;

	ASSERT (pe != NULL);
	char * str = NULL;
	HRESULT hr;

	char* stderror = "unknown error";

	if (pe->bstrDescription == NULL) {
		if (pe->pfnDeferredFillIn != NULL) {
			hr = (pe->pfnDeferredFillIn)(pe);
			if (FAILED (hr) || pe->bstrDescription==NULL)
				return HRESULT2Str(hr);
		}
		else
		{
			str = Tcl_Alloc (strlen(stderror)+1);
			strcpy (str, stderror);
			return str;
		}
	}

	TDString s;
	s.set("error - ");

	if (pe->bstrSource != NULL)
		s << "source: \"" << W2A(pe->bstrSource) << "\" ";
	s << "description: \"" << W2A(pe->bstrDescription) << "\"";
	str = Tcl_Alloc(s.length () + 1);
	strcpy (str, s);
	return str;
}



/*
 *-------------------------------------------------------------------------
 * Name2ID --
 *	Converts a name of a dispatch member to an id.
 * Result:
 *	Either DISPID_UNKNOWN if failed or the dispid.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
DISPID	Name2ID (IDispatch *pdisp, const char *name)
{
	USES_CONVERSION;
	ASSERT (pdisp != NULL && name != NULL);
	LPOLESTR olestr = A2OLE ((char*)name);
	DISPID dispid = DISPID_UNKNOWN;

	pdisp->GetIDsOfNames (IID_NULL, &olestr, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	return dispid;
}


/*
 *-------------------------------------------------------------------------
 * Name2ID --
 *	Converts a name (OLE string) of a dispatch member to an id.
 * Result:
 *	Either DISPID_UNKNOWN if failed or the dispid.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
DISPID		Name2ID (IDispatch *pdisp, const LPOLESTR olename)
{
	DISPID dispid = DISPID_UNKNOWN;
	pdisp->GetIDsOfNames (IID_NULL, (LPOLESTR*)&olename, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
	return dispid;
}



void OptclVariantClear (VARIANT *pvar)
{
	ASSERT (pvar != NULL);
	if ((pvar->vt & VT_BYREF) || pvar->vt == VT_VARIANT) {
		switch (pvar->vt & (~VT_BYREF)) {
		case VT_VARIANT:
			OptclVariantClear (pvar->pvarVal);
			g_pmalloc->Free (pvar->pvarVal);
			break;
		case VT_ERROR:
		case VT_I2:
		case VT_UI1:
			g_pmalloc->Free (pvar->piVal);
			break;
		// long
		case VT_HRESULT:
		case VT_I4:
		case VT_UI2:
		case VT_INT:
			g_pmalloc->Free (pvar->plVal);
			break;
		// float
		case VT_R4:
			g_pmalloc->Free (pvar->pfltVal);
			break;

		// double
		case VT_R8:
			g_pmalloc->Free (pvar->pdblVal);
			break;

		// boolean
		case VT_BOOL:
			g_pmalloc->Free (pvar->pboolVal);
			break;
		// object
		case VT_UNKNOWN:
		case VT_DISPATCH:
			if (pvar->ppunkVal != NULL) {
				(*(pvar->ppunkVal))->Release();
				g_pmalloc->Free (pvar->ppunkVal);
			}
			break;
		case VT_CY:
			g_pmalloc->Free (pvar->pcyVal);
			break;
		case VT_DATE:
			g_pmalloc->Free (pvar->pdate);
			break;
		case VT_BSTR:
			if (pvar->pbstrVal != NULL) {
				SysFreeString (*(pvar->pbstrVal));
				g_pmalloc->Free (pvar->pbstrVal);
			}
		break;
		case VT_RECORD:
		case VT_VECTOR:
		case VT_ARRAY:
		case VT_SAFEARRAY:
			ASSERT (FALSE); // case not handled yet
			break;

		default:
			ASSERT (FALSE); // unknown type
		}
	}
	else
		VariantClear (pvar);
}




bool var2obj_byref (Tcl_Interp *pInterp, VARIANT &var, TObjPtr &presult, OptclObj **ppObj)
{
	ASSERT (var.ppunkVal != NULL);

	USES_CONVERSION;

	bool		bOk = false;
	BSTR		bstr = NULL;
	HRESULT		hr = S_OK;
	OptclObj *	pObj = NULL;


	presult.create();
	if (var.ppunkVal == NULL) {
		presult = 0;
		return true;
	}
	try {
		switch (var.vt & ~VT_BYREF)
		{
		case VT_DISPATCH:
		case VT_UNKNOWN:
			if (*var.ppunkVal != NULL) {
				pObj = g_objmap.Add (pInterp, *var.ppunkVal);
				presult = (const char*)(*pObj); // cast to char*
				if (ppObj != NULL)
					*ppObj = pObj;
			}
			else
				presult = 0;
			break;
		case VT_BOOL:
			presult = (bool)(*var.pboolVal != 0);
			break;
		
		case VT_ERROR:
		case VT_I2:
			presult = *var.piVal;
			break;
		
		case VT_HRESULT:
		case VT_I4:
		case VT_UI2:
		case VT_INT:
			presult = *var.plVal;
			break;
		case VT_R4:
			presult = (double)(*var.pfltVal);
			break;
		case VT_R8:
			presult = (double)(*var.pdblVal);
			break;
		case VT_BSTR:
			presult = OLE2A(*var.pbstrVal);
			break;
		case VT_CY:
			hr = VarBstrFromCy (*var.pcyVal, LOCALE_SYSTEM_DEFAULT, NULL, &bstr);
			CHECKHR_TCL(hr, pInterp, false);
			break;
		case VT_DATE:
			hr = VarBstrFromDate (*var.pdblVal, LOCALE_SYSTEM_DEFAULT, NULL, &bstr);
			CHECKHR_TCL(hr, pInterp, false);
			break;
		case VT_VARIANT:
			if (var.pvarVal == NULL) {
				Tcl_SetResult (pInterp, "pointer to null", TCL_STATIC);
				bOk = false;
			} else {
				bOk = var2obj (pInterp, *var.pvarVal, presult, ppObj);
			}
			break;
		default:
			presult = "?unhandledtype?";
		}
		bOk = true;
		if (bstr != NULL) {
			presult = OLE2A(bstr);
			SysFreeString (bstr); bstr = NULL;
		}
	}
	catch (char *err) {
		Tcl_SetResult (pInterp, err, TCL_VOLATILE);
	}
	return bOk;
}


/*
 *-------------------------------------------------------------------------
 * var2obj --
 *	Converts a variant to a Tcl_Obj without type information.
 * Result:
 *	true iff successful, else interpreter holds error string.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
bool var2obj (Tcl_Interp *pInterp, VARIANT &var, TObjPtr &presult, OptclObj **ppObj /* = NULL*/)
{
	USES_CONVERSION;

	ASSERT (pInterp != NULL);
	ASSERT (ppObj == NULL || *ppObj ==  NULL);


	OptclObj *	pObj = NULL;
	_variant_t	comvar;
	HRESULT		hr = S_OK;
	_bstr_t		name;
	bool		bOk = false;
	

	if ((var.vt & VT_ARRAY) || (var.vt & VT_VECTOR)) {
		Tcl_SetResult (pInterp, "can't handle arrays or vectors for now", TCL_STATIC);
		return false;
	}

	if (var.vt == VT_VARIANT) {
		ASSERT (var.pvarVal != NULL);
		return var2obj (pInterp, *(var.pvarVal), presult, ppObj);
	}

	if (var.vt & VT_BYREF)
		return var2obj_byref (pInterp, var, presult, ppObj);

	presult.create();

	try {
		switch (var.vt)
		{
		case VT_DISPATCH:
		case VT_UNKNOWN:
			if (var.punkVal != NULL) {
				pObj = g_objmap.Add (pInterp, var.punkVal);
				presult = (const char*)(*pObj); // cast to char*
				if (ppObj != NULL)
					*ppObj = pObj;
			}
			else
				presult = 0;
			break;
		case VT_BOOL:
			presult = (bool)(var.boolVal != 0);
			break;
		case VT_I2:
			presult = var.iVal;
			break;
		case VT_I4:
			presult = var.lVal;
			break;
		case VT_R4:
			presult = (double)(var.fltVal);
			break;
		case VT_R8:
			presult = (double)(var.dblVal);
			break;
		default: // standard string conversion required
			comvar = var;
			name = comvar;
			presult = (char*)name;
		}
		bOk = true;
	}

	catch (HRESULT hr) {
		Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
	}
	catch (_com_error ce) {
		Tcl_SetResult (pInterp, T2A((TCHAR*)ce.ErrorMessage()), TCL_VOLATILE);
	}
	catch (char *err) {
		Tcl_SetResult (pInterp, err, TCL_VOLATILE);
	}

	return bOk;
}







/*
 *-------------------------------------------------------------------------
 * obj2var_ti --
 *	converts a Tcl_Obj to a variant using type information.
 *
 * Result:
 *	true iff successful, else interpreter holds error string.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
// nb - pInfo is the context for pdesc
bool obj2var_ti (Tcl_Interp *pInterp, TObjPtr &obj, VARIANT &var, 
						  ITypeInfo *pInfo, TYPEDESC *pdesc)
{
	ASSERT ((pInfo == NULL && pdesc == NULL) || (pInfo != NULL && pdesc != NULL));
	ASSERT (pInterp != NULL);

	OptclTypeAttr		ota;
	CComPtr<ITypeInfo>	pcurrent;
	CComPtr<IUnknown>	ptmpunk;
	HRESULT				hr;
	TObjPtr				ptmp;
	bool				bOk = false;
	OptclObj *			pOptclObj = NULL;
	long				lValue;

	// if no type description has been provided, do a simple conversion
	if (pdesc == NULL) {
		obj2var (obj, var);
		bOk = true;
	}

	// a simple type
	else if (pdesc->vt != VT_USERDEFINED && pdesc->vt != VT_SAFEARRAY) {
		if (pdesc->vt != VT_PTR) 
			return obj2var_vt (pInterp, obj, var, pdesc->vt);
		else {
			ASSERT (pdesc->lptdesc->vt != VT_PTR &&
					pdesc->lptdesc->vt != VT_USERDEFINED &&
					pdesc->lptdesc->vt != VT_SAFEARRAY);

			if     (pdesc->lptdesc->vt == VT_PTR || 
					pdesc->lptdesc->vt == VT_USERDEFINED || 
					pdesc->lptdesc->vt == VT_SAFEARRAY)
			{
				Tcl_SetResult (pInterp, "can't convert - optcl doesn't support level of de-referencing", TCL_STATIC);
				return false;
			}	
			return obj2var_vt_byref (pInterp, obj, var, pdesc->lptdesc->vt);
		}
	}

	// arrays - should be easy to do - not enough time right now...
	else if (pdesc->vt == VT_SAFEARRAY) {
		// wont do arrays for now.
		Tcl_SetResult (pInterp, "optcl doesn't currently handle array types", TCL_STATIC);
	}

	else {
		// type information provided and it refers to a user defined type
		// resolve the initial type

		hr = pInfo->GetRefTypeInfo (pdesc->hreftype, &ota.m_pti);
		CHECKHR(hr);
		g_libs.EnsureCached (ota.m_pti);
		hr = ota.GetTypeAttr();
		CHECKHR(hr);
		ASSERT (ota.m_pattr != NULL);
		pcurrent = pInfo;

		while (ota->typekind == TKIND_ALIAS && 
			   ota->tdescAlias.vt == VT_USERDEFINED)
		{
			HREFTYPE href = ota->tdescAlias.hreftype;
			pcurrent = ota.m_pti;
			ota = NULL; // release the type attribute and type info 
			pcurrent->GetRefTypeInfo (href, &ota.m_pti);
			hr = ota.GetTypeAttr();
			CHECKHR(hr);
		}
		
		// we've now climbed back up the alias chain and have one of the following:
		// enum, record, module, interface, dispatch, coclass, union or alias to a basic type
		// The following we can't (currently) do anything useful with: record, union, module.

		if (ota.m_pattr->typekind == TKIND_ALIAS) 
			return obj2var_ti (pInterp, obj, var, pcurrent, &(ota->tdescAlias));


		TYPEKIND tk = ota->typekind;	// the metaclass
		GUID intfguid = ota->guid;	


		switch (tk)
		{
		case TKIND_ENUM:
			if (bOk = (Tcl_GetLongFromObj (NULL, obj, &lValue) == TCL_OK)) 
				obj2var(obj, var);
			else if (bOk = TypeLib_ResolveConstant (pInterp, obj, ptmp, ota.m_pti)) 
				obj2var (ptmp, var);
			break;
		
		case TKIND_DISPATCH:
		case TKIND_INTERFACE:
			// both these case require an object with the correct interface
			pOptclObj = g_objmap.Find (obj);
			if (pOptclObj != NULL) {
				ptmpunk = (IUnknown*)(*pOptclObj);
				ASSERT (ptmpunk != NULL);
				hr = ptmpunk->QueryInterface (intfguid, (void**)&(var.punkVal));
				CHECKHR(hr);
				V_VT(&var) = VT_UNKNOWN;
				bOk = true;
			} else 
				ObjectNotFound (pInterp, obj);
			break;

		case TKIND_COCLASS:
			pOptclObj = g_objmap.Find (obj);
			if (pOptclObj != NULL) {
				var.punkVal = (IUnknown*)(*pOptclObj);
				var.punkVal->AddRef();
				V_VT(&var) = VT_UNKNOWN;
				bOk = true;
			} else 
				ObjectNotFound (pInterp, obj);
			break;

		case TKIND_ALIAS: 
			ASSERT (FALSE); // should be hanlded above.
			break;

		// can't handle these types
		case TKIND_MODULE:
		case TKIND_RECORD:
		case TKIND_UNION:
			obj2var (obj, var);
			bOk = true;
			break;

		default:
			break;
		}
	}

	return bOk;
}







/*
 *-------------------------------------------------------------------------
 * obj2var --
 *	Converts a Tcl object to a variant without type information.
 *	If the Tcl object is null, then sets the value to zero.
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void	obj2var (TObjPtr &obj, VARIANT &var)
{
	_variant_t v;
	ASSERT (var.vt == VT_EMPTY);
	try {
		if (obj.isnull()) {
			var.lVal = 0; 
			var.vt = VT_I4;
		} else {

			if (Tcl_GetLongFromObj (NULL, obj, &V_I4(&var)) == TCL_OK)
				V_VT(&var) = VT_I4;

			else if (Tcl_GetDoubleFromObj (NULL, obj, &V_R8(&var)) == TCL_OK)
				V_VT(&var) = VT_R8;

			else {
				v.Attach (var);
				v = (char*)(obj);
				var = v.Detach();
			}

	#if _DEBUG
			if (obj->typePtr != NULL) {
				TRACE ("%s\n", obj->typePtr->name);
			}
	#endif // _DEBUG
		}
	}

	catch (_com_error ce) {
		throw (HRESULT(ce.Error()));
	}
}


static char memerr[] = "out of memory";

#define CHECKMEM_TCL(x, interp, action) if ((x) == NULL) { \
	Tcl_SetResult (interp, memerr, TCL_STATIC); \
	action; \
}
	
bool	obj2var_vt_byref (Tcl_Interp *pInterp, TObjPtr &obj, VARIANT &var, VARTYPE vt)
{
	USES_CONVERSION;

	ASSERT ((vt & VT_BYREF) == 0); // we know that this is a BYREF variant - we don't want it in the vt
	OptclObj *	pOptclObj = NULL;
	bool		bok = true;
	IUnknown *	pUnk= NULL;
	VARIANT		temp;


	if (vt == VT_VARIANT) {
		var.pvarVal = (VARIANT*)g_pmalloc->Alloc (sizeof(VARIANT));
		CHECKMEM_TCL(var.pvarVal, pInterp, return false);
		VariantInit (var.pvarVal);
		if (!obj2var_vt (pInterp, obj, *var.pvarVal, vt)) {
			g_pmalloc->Free(var.pvarVal);
			var.pvarVal = NULL;
			return false;
		}
		var.vt = vt | VT_BYREF;
		return true;
	}


	VariantInit(&temp);
	// perform the conversion into a temporary variant
	if (!obj2var_vt (pInterp, obj, temp, vt))
		return false;

	
	switch (temp.vt) {
	// short
	case VT_ERROR:
	case VT_I2:
	case VT_UI1:
		var.piVal = (short*)g_pmalloc->Alloc (sizeof(short));
		CHECKMEM_TCL(var.pvarVal, pInterp, bok = false);
		if (bok) *var.piVal = temp.iVal;
		break;

	// long
	case VT_HRESULT:
	case VT_I4:
	case VT_UI2:
	case VT_INT:
		var.plVal = (long*)g_pmalloc->Alloc (sizeof(long));
		CHECKMEM_TCL(var.plVal, pInterp, bok = false);
		if (bok) *var.plVal = temp.lVal;
		break;

	// float
	case VT_R4:
		var.pfltVal = (float*)g_pmalloc->Alloc(sizeof(float));
		CHECKMEM_TCL(var.pfltVal, pInterp, bok = false);
		if (bok) *var.pfltVal = temp.fltVal;
		break;

	// double
	case VT_R8:
		var.pdblVal = (double*)g_pmalloc->Alloc(sizeof(double));
		CHECKMEM_TCL(var.pdblVal, pInterp, bok = false);
		if (bok) *var.pdblVal = temp.dblVal;
		break;

	// boolean
	case VT_BOOL:
		var.pboolVal = (VARIANT_BOOL*)g_pmalloc->Alloc(sizeof(VARIANT_BOOL));
		CHECKMEM_TCL(var.pboolVal, pInterp, bok = false);
		if (bok) *var.pboolVal = temp.boolVal;
		break;

	// object
	case VT_UNKNOWN:
	case VT_DISPATCH:
		// now allocate the memory
		var.ppunkVal = (LPUNKNOWN*)g_pmalloc->Alloc(sizeof (LPUNKNOWN));
		CHECKMEM_TCL(var.ppunkVal, pInterp, bok = false);
		if (bok) {
			*var.ppunkVal = temp.punkVal;
			if (*var.ppunkVal != NULL)
				(*var.ppunkVal)->AddRef();
		}
		break;

	case VT_CY:
		var.pcyVal = (CY*)g_pmalloc->Alloc(sizeof(CY));
		CHECKMEM_TCL(var.pcyVal, pInterp, bok = false);
		if (bok) *var.pcyVal = temp.cyVal;
		break;
	case VT_DATE:
		var.pdate = (DATE*)g_pmalloc->Alloc(sizeof(DATE));
		CHECKMEM_TCL(var.pdate, pInterp, bok = false);
		if (bok) *var.pdate = temp.date;
		break;
	case VT_BSTR:
		var.pbstrVal = (BSTR*)g_pmalloc->Alloc(sizeof(BSTR));
		CHECKMEM_TCL(var.pdate, pInterp, bok = false);
		if (bok) {
			*var.pbstrVal = SysAllocStringLen (temp.bstrVal, SysStringLen(temp.bstrVal));
			if (*var.pbstrVal == NULL) {
				g_pmalloc->Free (var.pbstrVal); var.pbstrVal = NULL;
				Tcl_SetResult (pInterp, memerr, TCL_STATIC);
				bok = false;
			}
		}

		break;
	case VT_RECORD:
	case VT_VECTOR:
	case VT_ARRAY:
	case VT_SAFEARRAY:
		ASSERT (FALSE); // case not handled yet
		break;

	default:
		ASSERT (FALSE); // should never get here.
	}

	var.vt = temp.vt | VT_BYREF;
	VariantClear(&temp);
	return bok;
}


/*
 *-------------------------------------------------------------------------
 * obj2var_vt --
 *	Converts a Tcl object to a variant of a certain type.
 * Result:
 *	None.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
bool	obj2var_vt (Tcl_Interp *pInterp, TObjPtr &obj, VARIANT &var, VARTYPE vt)
{
	ASSERT ((vt & VT_BYREF) == 0); // house rules: no by-reference variants here!

	OptclObj *pOptclObj = NULL;
	IUnknown * ptmpunk = NULL;
	bool bOk = true;
	HRESULT hr;

	switch (vt)
	{
	case VT_DISPATCH:
	case VT_UNKNOWN:
		V_VT(&var) = vt;
		if (obj.isnull()) 
			var.punkVal = NULL;
		else {
			// attempt to cast from an optcl object
			pOptclObj = g_objmap.Find (obj);
			

			if (pOptclObj != NULL) { // found it?
				ptmpunk = (IUnknown*)(*pOptclObj); // pull out the IUnknown pointer
				ASSERT (ptmpunk != NULL);
				if (vt == VT_DISPATCH) {		   // query to IDispatch iff required
					hr = ptmpunk->QueryInterface (IID_IDispatch, (VOID**)&ptmpunk);
					CHECKHR_TCL(hr, pInterp, false);
				}
				else
					ptmpunk->AddRef();				// if not IDispatch, make sure we incr the refcount
				var.punkVal = ptmpunk;
			}
			else {
				ObjectNotFound (pInterp, obj);
				bOk = false;
			}
		}
		break;
	default:
		obj2var (obj, var);
		if (vt != VT_VARIANT) {
			HRESULT hr = VariantChangeType (&var, &var, NULL, vt);
			if (FAILED (hr)) {
				Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
				bOk = false;
			}
		}
		break;
	}
	return bOk;
}



/*
 *-------------------------------------------------------------------------
 * ObjectNotFound --
 *	Standard error message when an optcl object is not found.
 * Result:
 *	TCL_ERROR always.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
int ObjectNotFound (Tcl_Interp *pInterp, const char *name)
{
	Tcl_SetResult (pInterp, "could not find object '", TCL_STATIC);
	Tcl_AppendResult (pInterp, (char*)name, "'", NULL);
	return TCL_ERROR;
}



/*
 *-------------------------------------------------------------------------
 * SplitTypedString --
 *	If pstr is of the format "a.b.c" then it is modified such that
 *		pstr == "a.b" and *ppsecond = "c"
 *	Otherwise, pstr will point to the original string and *ppsecond will
 *	be NULL.
 * Result:
 *	None.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void SplitTypedString (char *pstr, char ** ppsecond)
{
	const char * token = ".";
	ASSERT (pstr != NULL && ppsecond != NULL);

	char *p = pstr;
	while (*p != '.' && *p != '\0') p++;
	if (*p == '\0') {
		*ppsecond = NULL;
		return;
	}
		
	pstr = strtok (pstr, token);
	pstr[strlen(pstr)] = '.';

	for (short i = 0; i < 2; i++)
	{
		*ppsecond = strtok (NULL, token);
		if (*ppsecond == NULL)
			break;
	}
}



/*
 *-------------------------------------------------------------------------
 * SplitObject --
 *	Splits a string held within a Tcl object (pObj) into its constituent
 *	objects (ppResult), using a collection tokens.
 *
 * Result:
 *	true iff successful. Else, error string in interpreter.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
bool	SplitObject (Tcl_Interp *pInterp, Tcl_Obj *pObj, 
						 const char * tokens, Tcl_Obj **ppResult)
{
	ASSERT (pInterp != NULL && pObj != NULL && tokens != NULL && ppResult != NULL);
	TObjPtr cmd;
	cmd.create();
	cmd = "split";
	cmd.lappend (pObj);
	cmd.lappend(tokens);
	if (Tcl_EvalObj (pInterp, cmd) == TCL_ERROR)
		return false;
	*ppResult = Tcl_GetObjResult (pInterp);
	Tcl_IncrRefCount (*ppResult);
	return true;
}


bool SplitBrackets (Tcl_Interp *pInterp, Tcl_Obj *pObj,
						   TObjPtr & result)
{
	ASSERT (pInterp != NULL && pObj != NULL);
	TObjPtr pcmd ("regexp -nocase {^([^\\(\\)])+(\\([^\\(\\)]+\\))?$} ");
	pcmd.lappend(pObj);

	if (Tcl_EvalObj (pInterp, pcmd) == TCL_ERROR)
		return false;

	CONST84 char * okstr = Tcl_GetStringResult (pInterp);
	if (okstr[0] == '0') {
		Tcl_SetResult (pInterp, "property format is incorrect: ", TCL_STATIC);
		Tcl_AppendResult (pInterp, Tcl_GetStringFromObj(pObj, NULL), NULL);
		return false;
	}

	pcmd = "split";
	pcmd.lappend (pObj).lappend("(),");
	if (Tcl_EvalObj (pInterp, pcmd) == TCL_ERROR)
		return false;
	result.copy(Tcl_GetObjResult (pInterp));

	// the last element will be a null string
	char *str = Tcl_GetStringFromObj (pObj, NULL);
	if (str[strlen (str) - 1] == ')')
		Tcl_ListObjReplace (NULL, result, result.llength() - 1, 1, 0, NULL);
	return true;
}

/// Tests
TCL_CMDEF (Obj2VarTest)
{
	if (objc < 2) {
		Tcl_WrongNumArgs(pInterp, 1, objv, "value");
		return TCL_ERROR;
	}

	VARIANT var;
	VARIANT * pvar;
	HRESULT hr;

	pvar = (VARIANT*)CoTaskMemAlloc(sizeof(VARIANT));
	
	VariantInit(pvar);
	var.vt = VT_VARIANT;
	var.pvarVal = pvar;

	TObjPtr ptr(objv[1], false);

	obj2var (ptr, *pvar);
	CoTaskMemFree((LPVOID)pvar);
	hr = VariantClear(&var);
	Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
	return FAILED(hr)?TCL_ERROR:TCL_OK;
}


