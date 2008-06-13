/*
 *------------------------------------------------------------------------------
 *	optclobj.cpp
 *	Implements the functionality for the internal  representation of 
 *	an optcl object.
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
#include <comdef.h>
#include "tbase.h"
#include "utility.h"
#include "optcl.h"
#include "OptclObj.h"
#include "typelib.h"
#include "ObjMap.h"
#include "dispparams.h"
#include "eventbinding.h"
#include "optclbindptr.h"
#include "optcltypeattr.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

OptclObj::OptclObj ()
: m_refcount(0), m_cmdtoken(NULL), m_pta(NULL),
m_destroypending(false), m_container(this)
{
}

bool OptclObj::Create (Tcl_Interp *pInterp, const char *strid, 
					const char * windowpath, bool start)
{
	m_pInterp = pInterp;

	USES_CONVERSION;
	ASSERT (strid != NULL);

	if (windowpath == NULL) {
		LPOLESTR 	lpolestrid = A2OLE(strid);
		CLSID		clsid;
		HRESULT		hr;
	
		// convert strid to CLSID
		hr = CLSIDFromString (lpolestrid, &clsid);
		if (FAILED (hr))
			hr = CLSIDFromProgID (lpolestrid, &clsid);
		CHECKHR_TCL(hr, pInterp, false);

		if (!start)
			hr = GetActiveObject(clsid, NULL, &m_punk);		
		if (start || FAILED(hr)) 
			hr = CoCreateInstance (clsid, NULL, CLSCTX_SERVER, IID_IUnknown, (void**)&m_punk);
		CHECKHR_TCL(hr, pInterp, false);
		
	}
	else {
		m_punk = m_container.Create(pInterp, Tk_MainWindow(pInterp), windowpath, strid);
		if (m_punk == NULL)
			return false;
	}
	try {
		CreateName (m_punk);
		InitialiseClassInfo(m_punk);
		InitialisePointers (m_punk);
	}
	catch (HRESULT hr) {
		CHECKHR_TCL(hr, pInterp, false);
	}
	return true;
}




/*
 *-------------------------------------------------------------------------
 * OptclObj::Attach --
 *	Connects this object to an existing interface
 * Result:
 *	true iff successful
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
bool OptclObj::Attach (Tcl_Interp *pInterp, LPUNKNOWN punk)
{
	ASSERT (m_punk == NULL);
	ASSERT (punk != NULL);

	m_pInterp = pInterp;
	try {
		CreateName (punk);
		InitialiseUnknown(punk);
		InitialiseClassInfo(m_punk);
		InitialisePointers (m_punk);
	}
	catch (HRESULT hr) {
		m_punk = NULL;
		CHECKHR_TCL(hr, pInterp, false);
	}
	return true;
}





/*
 *-------------------------------------------------------------------------
 * OptclObj::~OptclObj --
 *	Destructor
 * Result:
 *
 * Side effects:
 *
 *-------------------------------------------------------------------------
 */
OptclObj::~OptclObj()
{
	m_destroypending = true;
	ReleaseBindingTable();
	ReleaseTypeAttr();
}






/*
 *-------------------------------------------------------------------------
 * OptclObj::CreateName --
 *	Creates the string representation for this object - a unique name is 
 *	created from the object's IUnknown pointer.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void OptclObj::CreateName (LPUNKNOWN punk)
{
	ASSERT (punk != NULL);
	char str[10];
	sprintf (str, "%x", punk);
	m_name = "optcl0x";
	m_name += str;
}







/*
 *-------------------------------------------------------------------------
 * OptclObj::InitialiseClassInfo --
 *	Attempts to find the typeinfo for this object's coclass.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void OptclObj::InitialiseClassInfo (LPUNKNOWN punk)
{
	CComQIPtr<IProvideClassInfo> pcli;

	// try to pull out the coclass information
	pcli = punk; // implicit query interface
	if (pcli != NULL)
		pcli->GetClassInfo (&m_pti_class);
}




/*
 *-------------------------------------------------------------------------
 * OptclObj::InitialiseUnknown --
 *	Initialises the OptclObj's 'true' IUnknown pointer. 
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	Throws HRESULT on error.
 *-------------------------------------------------------------------------
 */
void OptclObj::InitialiseUnknown (LPUNKNOWN punk)
{
	ASSERT (punk != NULL);
	HRESULT hr;

	hr = punk->QueryInterface (IID_IUnknown, (void**)(&m_punk));
	CHECKHR(hr);
}




/*
 *-------------------------------------------------------------------------
 * OptclObj::InitialisePointersFromCoClass --
 *	Called when we have the coclass information for this object. The 
 *	function identifies the default interface and binds to it.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
HRESULT OptclObj::InitialisePointersFromCoClass()
{
	ASSERT (m_pti_class != NULL);
	TYPEATTR *pta = NULL;
	HRESULT hr;
	
	// retrieve the type attribute
	hr = m_pti_class->GetTypeAttr (&pta);
	CHECKHR(hr);

	// store the number of implemented interfaces
	WORD impcount = pta->cImplTypes;
	m_pti_class->ReleaseTypeAttr (pta); pta = NULL;

	// iterate through the type looking for the default interface
	for (WORD i = 0; i < impcount; i++)
	{
		INT flags;
		hr = m_pti_class->GetImplTypeFlags (i, &flags);
		if (FAILED (hr))
			return hr;
		if (flags == IMPLTYPEFLAG_FDEFAULT)
			break;
	}

	// if not found return an error
	if (i == impcount)
		return E_FAIL;

	// we found the interface - now to get its iid...
	// first retrieve the type info;

	CComPtr<ITypeInfo> reftype;
	CComPtr<ITypeLib>  reftypelib;

	HREFTYPE href;

	hr = m_pti_class->GetRefTypeOfImplType (i, &href);
	if (FAILED(hr))
		return hr;

	hr = m_pti_class->GetRefTypeInfo (href, &reftype);
	if (FAILED(hr))
		return hr;

	// now set the interface from typeinfo
	return SetInterfaceFromType (reftype);
}




/*
 *-------------------------------------------------------------------------
 * OptclObj::GetTypeAttr --
 *	Retrieves the type attribute for the type of this object's current 
 *	interface.
 * Result:
 *	Standard HRESULT.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
HRESULT OptclObj::GetTypeAttr()
{
	ASSERT (m_pta == NULL);
	ASSERT (m_pti != NULL);
	return m_pti->GetTypeAttr(&m_pta);
}


/*
 *-------------------------------------------------------------------------
 * OptclObj::ReleaseTypeAttr --
 *	Releases the current type attribute, and sets it to NULL.
 * Result:
 *	None.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void OptclObj::ReleaseTypeAttr()
{
	if (m_pti != NULL && m_pta != NULL) {
		m_pti->ReleaseTypeAttr(m_pta);
		m_pta = NULL;
	}
}


/*
 *-------------------------------------------------------------------------
 * OptclObj::SetInterfaceFromType --
 *	Queries for the interface described in the typeinfo. The interface,
 *	if found, becomes the current interface.
 *
 * Result:
 *	HRESULT giving success code.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
HRESULT OptclObj::SetInterfaceFromType (ITypeInfo *reftype)
{
	HRESULT hr;
	CComPtr<ITypeLib> reftypelib;
	UINT libindex;
	TYPEATTR *pta;

	hr = reftype->GetContainingTypeLib(&reftypelib, &libindex);
	if (FAILED(hr))
		return hr;

	hr = reftype->GetTypeAttr (&pta);
	if (FAILED(hr))
		return hr;

	if (pta->typekind != TKIND_DISPATCH) {
		reftype->ReleaseTypeAttr (pta);
		return E_NOINTERFACE;
	}

	GUID guid = pta->guid;
	reftype->ReleaseTypeAttr (pta);
	
	hr = m_punk->QueryInterface(guid, (void**)(&m_pcurrent));
	if (FAILED(hr))
		return hr;

	
	// nice! now we cache the result of all our hard work
	ReleaseTypeAttr ();
	m_pti = reftype;
	m_ptl = reftypelib;
	m_ptc = NULL;
	m_pti->GetTypeComp (&m_ptc);
	// now that we got the interface ok, retrieve the type attributes again
	return GetTypeAttr();
}





/*
 *-------------------------------------------------------------------------
 * OptclObj::InitialisePointers --
 *	Called to initialise this objects interface pointers
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void OptclObj::InitialisePointers (LPUNKNOWN punk, ITypeLib *plib, ITypeInfo *pinfo)
{
	HRESULT hr;
	ASSERT (punk != NULL);
	CComQIPtr<IDispatch> pdisp;

	ASSERT ((plib!=NULL && pinfo!=NULL) || (plib==NULL && pinfo==NULL));

	if (plib != NULL && pinfo != NULL) {
		m_pcurrent = punk;
		m_ptl = plib;
		m_pti = pinfo;
		m_ptc = NULL;
		m_pti->GetTypeComp (&m_ptc);
		GetTypeAttr();
	} 

	// else, if we have the coclass information, try building on its default
	// interface
	else if (m_pti_class == NULL || FAILED(InitialisePointersFromCoClass())) {
		// failed to build using coclass information
		// Query Interface cast to a dispatch interface
		m_pcurrent = punk;
		try {
			if (m_pcurrent == NULL)
				throw (HRESULT(0));
			// get the type information and library.
			hr = m_pcurrent->GetTypeInfo (0, LOCALE_SYSTEM_DEFAULT, &m_pti);
			CHECKHR(hr);
			UINT index;
			hr = m_pti->GetContainingTypeLib(&m_ptl, &index);
			CHECKHR(hr);
			m_ptc = NULL;
			m_pti->GetTypeComp (&m_ptc);
			GetTypeAttr();
		}
		

		catch (HRESULT) {
			// there isn't a interface that we can use
			ReleaseTypeAttr();
			m_pcurrent.Release();
			m_pti = NULL;
			m_ptl = NULL;
			m_ptc = NULL;
			return;
		}
	}
	// inform the typelibrary browser system of the library
	g_libs.EnsureCached (m_ptl);
}




/*
 *-------------------------------------------------------------------------
 * OptclObj::operator LPUNKNOWN --
 *	Gives the 'true' IUnknown pointer for this object.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
OptclObj::operator LPUNKNOWN()
{
	return m_punk;
}




/*
 *-------------------------------------------------------------------------
 * OptclObj::operator const char * --
 *	Gives the string representation for this object.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
OptclObj::operator const char * ()
{
	return m_name.c_str();
}



/*
 *-------------------------------------------------------------------------
 * void OptclObj::CoClassName --
 *	Returns the class name in the tcl object smart ptr or ??? if unknown.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void OptclObj::CoClassName (TObjPtr &pObj)
{
	pObj.create();
	if (m_pti_class == NULL)
		pObj = "???";
	else
		TypeLib_GetName (NULL, m_pti_class, pObj);
}


/*
 *-------------------------------------------------------------------------
 * OptclObj::InterfaceName --
 *	Returns the name of this objects current interface.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void OptclObj::InterfaceName (TObjPtr &pObj)
{
	pObj.create();
	if (m_pti == NULL)
		pObj = "???";
	else
		TypeLib_GetName (NULL, m_pti, pObj);
}



/*
 *-------------------------------------------------------------------------
 * OptclObj::SetInterfaceName --
 *	Sets the current interface to that named by pObj.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void OptclObj::SetInterfaceName (TObjPtr &pObj)
{
	ASSERT (pObj.isnotnull());
	TypeLib *ptl;
	CComPtr<ITypeInfo> pti;
	CComPtr<IUnknown> punk;
	TYPEATTR ta, *pta = NULL;
	HRESULT hr;

	TypeLib_ResolveName (pObj, &ptl, &pti);
	// we need to insert some alias type resolution here.

	hr = pti->GetTypeAttr (&pta);
	CHECKHR(hr);
	ta = *pta;
	pti->ReleaseTypeAttr (pta);


	if (ta.typekind != TKIND_INTERFACE &&
		ta.typekind != TKIND_DISPATCH)
		throw ("type does not resolve to an interface");

	
	hr = m_punk->QueryInterface (ta.guid, (void**)(&punk));
	CHECKHR(hr);
	InitialisePointers (punk, ptl->m_ptl, pti);
}




/*
 *-------------------------------------------------------------------------
 * OptclObj::InvokeCmd --
 *	Called by the object map as a result of invoking the object command
 *	on this object. Format of the command is as follows
 *
 *		obj : ?-with subprop? prop ?value? ?prop value? ...
 *		obj method ?arg? ...
 *
 * Result:
 *	true iff successful. Error string in interpreter.
 *
 * Side effects:
 *	Depends on the parameters.
 *-------------------------------------------------------------------------
 */
bool OptclObj::InvokeCmd (Tcl_Interp *pInterp, int objc, Tcl_Obj *CONST objv[])
{
	ASSERT (pInterp != NULL);
	CComPtr<IDispatch> pdisp;
	CComPtr<ITypeComp> ptc;
	CComPtr<ITypeInfo> pti;
	TObjPtr name;
	
	int		invkind = DISPATCH_METHOD;

	char * msg = 			
		"\n\tobj : ?-with subprop? prop ?value? ?prop value? ..."
		"\n\tobj method ?arg? ...";

	if (objc == 0) {
		Tcl_WrongNumArgs (pInterp, 0, NULL, msg);
		return TCL_ERROR;
	}

	if (CheckInterface (pInterp) == false)
		return TCL_ERROR;

	

	// parse for a -with flag
	name.attach(objv[0]);
	if (strncmp (name, "-with", strlen(name)) == 0)
	{
		// check that we have enough parameters
		if (objc < 3) {
			Tcl_WrongNumArgs (pInterp, 0, NULL, msg);
			return false;
		}

		name.attach(objv[1]);
		if (!ResolvePropertyObject (pInterp, name, &pdisp, &pti, &ptc))
			return false;
		objc -= 2;
		objv += 2;
	}
	else {
		pdisp = m_pcurrent;
		ptc = m_ptc;
		pti = m_pti;
	}

	// check the first argument for a ':'
	char * str = Tcl_GetStringFromObj (objv[0], NULL);
	ASSERT (str != NULL);
	
	if (*str == ':') {
		objc--;
		objv++;

		if (objc == 1) 
			return GetProp (pInterp, objv[0], pdisp, pti, ptc);
		else {
			if (objc % 2 != 0) {
				Tcl_SetResult (pInterp, "property set requires pairs of parameters", TCL_STATIC);
				return false;
			}
			return SetProp (pInterp, objc/2, objv, pdisp, pti, ptc);
		}
	}

	if (ptc == NULL)
		return InvokeNoTypeInf (pInterp, invkind, objc, objv, pdisp);
	else
		return InvokeWithTypeInf (pInterp, invkind, objc, objv, pdisp, pti, ptc);
}





/*
 *-------------------------------------------------------------------------
 * OptclObj::CheckInterface --
 *	Checks for current interface being valid.
 * Result:
 *	Currently, returns true iff an interface exists.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
bool OptclObj::CheckInterface (Tcl_Interp *pInterp)
{
	if (m_pcurrent == NULL) {
		Tcl_SetResult (pInterp, "no interface available", TCL_STATIC);
		return false;
	}

	/* -- not needed now that we are only working with dispatch interfaces

	if (m_pta != NULL) {
		if (m_pta->typekind == TKIND_INTERFACE && ((m_pta->wTypeFlags&TYPEFLAG_FDUAL)==0))
		{
			Tcl_SetResult (pInterp, "interface is a pure vtable - optcl can't call these ... yet!", TCL_STATIC);
			return false;
		}
	}
	*/
	return true;
}


/*
 *-------------------------------------------------------------------------
 * OptclObj::BuildParamsWithBindPtr --
 *	Builds the dispatch parameters using the values found in a bindptr 
 *	object.
 *
 * Result:
 *	true iff successful - else error string in interpreter.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
bool OptclObj::BuildParamsWithBindPtr (Tcl_Interp *pInterp, int objc, Tcl_Obj *CONST objv[],
									   OptclBindPtr & bp,  DispParams & dp)
{
	ASSERT (pInterp != NULL && objv != NULL);
	bool	con_ok = true;
	TObjPtr obj;


	// check for the last parameter being the return value and take
	// this into account when checking parameter counts
	int params = bp.cParams ();
	if (params > 0 && bp.param(params - 1)->paramdesc.wParamFlags & PARAMFLAG_FRETVAL)
		--params;

	if (objc <= params &&
		objc >= (params - bp.cParamsOpt()))
	{
		// set up the dispatch arguments - must be in reverse order
		dp.Args (objc);
		for (int count = objc-1; count >= 0 && con_ok; count--)
		{
			con_ok = false;
			ELEMDESC *pdesc = bp.param(count);
			ASSERT (pdesc != NULL);
			// cases for parameters : [in] - value
			//						  [inout] - reference (variable must exist) 
			//						  [out] - reference (variable doesn have to exist)

			// is it an in* type 
			if ((pdesc->paramdesc.wParamFlags  & PARAMFLAG_FIN) || (pdesc->paramdesc.wParamFlags == PARAMFLAG_NONE)) {
				// is it [inout]?
				if (pdesc->paramdesc.wParamFlags  & PARAMFLAG_FOUT) {
					obj.attach(Tcl_ObjGetVar2 (pInterp, objv[count], NULL, TCL_LEAVE_ERR_MSG));
					if (obj.isnull()) 
						return false;
				}
				else // just [in]
					obj.attach(objv[count]);

				con_ok = obj2var_ti(pInterp, obj, dp[objc - count - 1], bp.m_pti, &(pdesc->tdesc));
			}

			else if (pdesc->paramdesc.wParamFlags  & PARAMFLAG_FOUT)
			{ // a pure out flag - we'll set up the type of the parameter correctly, but fill it with a null
				con_ok = obj2var_ti(pInterp, TObjPtr(NULL), dp[objc - count - 1], bp.m_pti, &(pdesc->tdesc));
			}

			else  {
				// unknown parameter type
				ASSERT(false);
			}
		}
	}
	else
	{
		Tcl_SetResult (pInterp, "wrong # args", TCL_STATIC);
		con_ok = false;
	}

	return con_ok;
}



/*
 *-------------------------------------------------------------------------
 * OptclObj::RetrieveOutParams --
 *	Scans the parameter types in a bind pointer and pulls out those that
 *	are either out or in/out, and sets the appropriate Tcl variable to 
 *	their value.
 *
 * Result:
 *	true iff successful. Else, error string in interpreter.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
bool OptclObj::RetrieveOutParams (Tcl_Interp *pInterp, int objc, Tcl_Obj *CONST objv[],
								  OptclBindPtr & bp,  DispParams & dp)
								 
{ 
	TObjPtr presult;
	bool bok = true;
	// now loop through the parameters again, pulling out the [*out] values
	for (int count = objc - 1; bok && count >= 0; count--)
	{
		ELEMDESC *pdesc = bp.param(count);
		ASSERT (pdesc != NULL);
		// is it an out parameter?
		if (pdesc->paramdesc.wParamFlags & PARAMFLAG_FOUT)
		{
			// convert the value back to a tcl object
			bok = (!var2obj (pInterp, dp[objc - count - 1], presult) ||
					Tcl_ObjSetVar2 (pInterp, objv[count], NULL, 
					                presult, TCL_LEAVE_ERR_MSG) == NULL);
				
		}
	}
	return bok;
}






bool OptclObj::InvokeWithTypeInfVariant (Tcl_Interp *pInterp, long invokekind,
								  int objc, Tcl_Obj *CONST objv[], 
								  IDispatch *pDisp, ITypeInfo *pti, ITypeComp *pCmp, VARIANT &varResult)
{
	USES_CONVERSION;
	DispParams	dp;
	LPOLESTR	olename;
	

	DISPID		dispid;
	HRESULT		hr;
	EXCEPINFO	ei;
	UINT		ea = 0;

	bool		bOk = false;
	TObjPtr		obj;
	TObjPtr		presult;
	static  DISPID		propput = DISPID_PROPERTYPUT;
	OptclBindPtr	obp;
	OptclTypeAttr	ota;

	ASSERT (objc >= 1);
	ASSERT (pDisp != NULL);
	ASSERT (pti != NULL);
	ASSERT (varResult.vt == VT_EMPTY);
	ota = pti;

	ASSERT (ota->typekind == TKIND_DISPATCH || (ota->wTypeFlags & TYPEFLAG_FDUAL));

	try {
		olename = A2OLE(Tcl_GetStringFromObj (objv[0], NULL));
		hr = pCmp->Bind (olename, LHashValOfName(LOCALE_SYSTEM_DEFAULT, olename), 
			invokekind, &obp.m_pti, &obp.m_dk, &obp.m_bp);
		CHECKHR(hr);

		if (obp.m_dk == DESCKIND_NONE) {
			Tcl_SetResult (pInterp, "member not found: ", TCL_STATIC);
			Tcl_AppendResult (pInterp, (char*)obj, NULL);
		} else {
			dispid = obp.memid();
			// check the number of parameters

			objc--; // count of parameters provided
			objv++; // the parameters
			if (!BuildParamsWithBindPtr (pInterp, objc, objv, obp, dp))
				return false;

			if (invokekind == DISPATCH_PROPERTYPUT) {
				dp.cNamedArgs = 1;
				dp.rgdispidNamedArgs = &propput;
			}

			// can't invoke through the typelibrary for local objects
			//hr = pti->Invoke(pDisp, dispid, invokekind, &dp, &varResult, &ei, &ea);
			hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, invokekind, 
				&dp, &varResult, &ei, &ea);

			if (invokekind == DISPATCH_PROPERTYPUT) {
				dp.rgdispidNamedArgs = NULL;
			}

			// error check
			if (hr == DISP_E_EXCEPTION)
				Tcl_SetResult (pInterp, ExceptInfo2Str (&ei), TCL_DYNAMIC);

			else if (hr == DISP_E_TYPEMISMATCH) {
				TDString td("type mismatch in parameter #");
				td << (long)(ea);
				Tcl_SetResult (pInterp, td, TCL_VOLATILE);
			}
			else
				CHECKHR_TCL(hr, pInterp, false);
			if (FAILED(hr))
				return false;

			if (!RetrieveOutParams (pInterp, objc, objv, obp, dp))
				return false;
			bOk = true;
		}
	}
	catch (HRESULT hr)
	{
		Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
	}
	return bOk;
}






/*
 *-------------------------------------------------------------------------
 * OptclObj::InvokeWithTypeInf --
 *	Performs a method invocation, given a dispatch interface and a 
 *	ITypeComp interface for typing.
 *
 * Result:
 *	true iff successful. Error string in interpreter.
 * Side effects:
 *	Depends on the method being invoked.
 *-------------------------------------------------------------------------
 */
bool OptclObj::InvokeWithTypeInf (Tcl_Interp *pInterp, long invokekind,
								  int objc, Tcl_Obj *CONST objv[], 
								  IDispatch *pDisp, ITypeInfo *pti, ITypeComp *pCmp)
{
	VARIANT varResult;
	VariantInit(&varResult);
	TObjPtr presult;

	bool bok;
	bok = InvokeWithTypeInfVariant (pInterp, invokekind, objc, objv, pDisp, pti, pCmp, varResult);

	// set the result of the operation to the return value of the function
	if (bok && (bok = var2obj(pInterp, varResult, presult)))
			Tcl_SetObjResult (pInterp, presult);
	VariantClear(&varResult);
	return bok;
}






/*
 *-------------------------------------------------------------------------
 * OptclObj::InvokeNoTypeInf --
 *	Performs a member invocation without any type information on an 
 *	IDispatch interface.
 *
 * Result:
 *	true iff successful. Else, error string in interpreter.
 * Side effects:
 *	Depends on the methods being invoked.
 *-------------------------------------------------------------------------
 */

bool OptclObj::InvokeNoTypeInf(	Tcl_Interp *pInterp, long invokekind, 
								int objc, Tcl_Obj *CONST objv[], 
								IDispatch *pDisp)
{
	VARIANT var;
	VariantInit (&var);
	TObjPtr presult;
	bool bok;

	if (bok = InvokeNoTypeInfVariant (pInterp, invokekind, objc, objv, pDisp, var)) {
		if (bok = var2obj(pInterp, var, presult))
			Tcl_SetObjResult (pInterp, presult);
		VariantClear(&var);
	}

	return bok;
}



/*
 *-------------------------------------------------------------------------
 * OptclObj::InvokeNoTypeInfVariant --
 *	The same as InvokeNoTypeInf, but instead of placing the result in 
 *	the interpreter, returns within a variant.
 * Result:
 *	true iff successful. Else, error string in interpreter
 * Side effects:
 *	Depends on member being invoked
 *-------------------------------------------------------------------------
 */
bool OptclObj::InvokeNoTypeInfVariant (	Tcl_Interp *pInterp, long invokekind, 
										int objc, Tcl_Obj *CONST objv[], 
										IDispatch *pDisp, VARIANT &varResult)
{
	ASSERT (varResult.vt == VT_EMPTY);

	DispParams dp;
	DISPID dispid;
	HRESULT hr;
	TObjPtr obj;
	EXCEPINFO ei;
	UINT ea = 0;
	bool bOk = false;

	ASSERT (objc >= 1);
	ASSERT (pDisp != NULL);

	obj.attach(objv[0]);
	dispid = Name2ID(pDisp, obj);
	if (dispid == DISPID_UNKNOWN) {
		Tcl_SetResult (pInterp, "member not found: ", TCL_STATIC);
		Tcl_AppendResult (pInterp, obj, NULL);
	} else {
		objc--; // count of parameters
		// set up the dispatch arguments - must be in reverse order
		dp.Args (objc);
		for (int i = objc-1; i >= 0; i--)
		{
			obj.attach(objv[i+1]);
			obj2var(obj, dp[i]);
		}

		hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, invokekind, 
			&dp, &varResult, &ei, &ea);
		if (hr == DISP_E_EXCEPTION)
			Tcl_SetResult (pInterp, ExceptInfo2Str (&ei), TCL_DYNAMIC);
		else if (hr == DISP_E_TYPEMISMATCH) {
			TDString td("type mismatch in parameter #");
			td << (long)(ea);
			Tcl_SetResult (pInterp, td, TCL_VOLATILE);
		}
		else if (FAILED(hr))
			Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
		else {
		}
	}

	return bOk;
}




/*
 *-------------------------------------------------------------------------
 * OptclObj::GetProp --
 *	Called to get the value of a property (a property can be indexed)
 *	If type information is provided, then it will be used in the invocation
 * Result:
 *	true iff ok. Else, error string in interpreter
 * Side effects:
 *	Depends on the property and its value.
 *-------------------------------------------------------------------------
 */
bool OptclObj::GetProp (Tcl_Interp *pInterp, Tcl_Obj *name, 
			  IDispatch *pdisp, ITypeInfo *pti, ITypeComp *ptc)
{
	ASSERT (pInterp != NULL && name != NULL && pdisp != NULL);
	TObjPtr params;
	bool bok;

	if (bok = SplitBrackets (pInterp, name, params)) {
		int length = params.llength();
		ASSERT (length >= 1);
		Tcl_Obj ** pplist = (Tcl_Obj **)malloc(sizeof(Tcl_Obj*) * length);
		if (pplist == NULL) {
			Tcl_SetResult (pInterp, "out of memory", TCL_STATIC);
			return false;
		}

		for (int p = 0; p < length; p++)
			pplist[p] = params.lindex(p);

		if (pti != NULL) {
			ASSERT (ptc != NULL);
			bok = InvokeWithTypeInf(pInterp, DISPATCH_PROPERTYGET, length, pplist, pdisp, pti, ptc);
		}
		else {
			bok = InvokeNoTypeInf (pInterp, DISPATCH_PROPERTYGET, length, pplist, pdisp);
		}

		free(pplist);
	}
	return bok;
}




/*
 *-------------------------------------------------------------------------
 * OptclObj::GetIndexedVariant --
 *	Called to get the value of a property or the return type of method, with
 *	bracket indexing.
 *	If type information is provided, then it will be used in the invocation
 * Result:
 *	true iff ok. Else, error string in interpreter
 * Side effects:
 *	Depends on the property and its value.
 *-------------------------------------------------------------------------
 */
bool OptclObj::GetIndexedVariant (Tcl_Interp *pInterp, Tcl_Obj *name, 
			  IDispatch *pdisp, ITypeInfo *pti, ITypeComp *ptc, VARIANT &varResult)
{
	ASSERT (pInterp != NULL && name != NULL && pdisp != NULL);
	ASSERT (varResult.vt == VT_EMPTY);

	TObjPtr params;
	TObjPtr presult;
	bool bok;
	static const int invkind = DISPATCH_PROPERTYGET|DISPATCH_METHOD;
	if (bok = SplitBrackets (pInterp, name, params)) {
		int length = params.llength();
		ASSERT (length >= 1);
		Tcl_Obj ** pplist = (Tcl_Obj **)malloc(sizeof(Tcl_Obj*) * length);
		if (pplist == NULL) {
			Tcl_SetResult (pInterp, "out of memory", TCL_STATIC);
			return false;
		}

		for (int p = 0; p < length; p++)
			pplist[p] = params.lindex(p);

		if (pti != NULL) {
			ASSERT (ptc != NULL);
			bok = InvokeWithTypeInfVariant (pInterp, invkind, length, pplist, pdisp, pti, ptc, varResult);
		}
		else {
			bok = InvokeNoTypeInfVariant (pInterp, invkind, length, pplist, pdisp, varResult);
		}
		free(pplist);
	}
	return bok;
}

bool	OptclObj::SetProp (Tcl_Interp *pInterp, 
						   int paircount, Tcl_Obj * CONST namevalues[], 
						   IDispatch *pdisp, ITypeInfo *pti, ITypeComp *ptc)
{
	bool bok = true;
	ASSERT (pInterp != NULL && paircount > 0 && namevalues != NULL && pdisp != NULL);
	for (int i = 0; bok && i < paircount; i++)
	{
		TObjPtr params;
		if (bok = SplitBrackets (pInterp, namevalues[0], params)) {
			params.lappend(namevalues[1]);
			int length = params.llength();
			ASSERT (length >= 1);
			Tcl_Obj ** pplist = (Tcl_Obj **)malloc(sizeof(Tcl_Obj*) * length);
			if (pplist == NULL) {
				Tcl_SetResult (pInterp, "out of memory", TCL_STATIC);
				return false;
			}


			for (int p = 0; p < length; p++)
				pplist[p] = params.lindex(p);

			if (pti != NULL) {
				ASSERT (ptc != NULL);
				bok = InvokeWithTypeInf(pInterp, DISPATCH_PROPERTYPUT, length, pplist, pdisp, pti, ptc);
			}
			else {
				bok = InvokeNoTypeInf (pInterp, DISPATCH_PROPERTYPUT, length, pplist, pdisp);
			}
			namevalues += 2;
			free(pplist);
		}
	}
	if (bok)
		Tcl_ResetResult (pInterp);
	return true;
}







/*
 *-------------------------------------------------------------------------
 * OptclObj::GetPropVariantDispatch --
 *	Retrieves the value of property as a variant, relative to a dispatch
 *	interface.
 *
 * Result:
 *	true iff successful.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
bool OptclObj::GetPropVariantDispatch (Tcl_Interp *pInterp, const char*name, 
									IDispatch* pcurrent, VARIANT &varResult)

{
	USES_CONVERSION;

	ASSERT (pcurrent != NULL && pInterp != NULL);

	DISPID		dispid;
	HRESULT		hr;
	DISPPARAMS	dispparamsNoArgs; SETNOPARAMS (dispparamsNoArgs);
	EXCEPINFO	ei;
	bool		bOk = false;

	dispid = Name2ID (pcurrent, name);
	if (dispid == DISPID_UNKNOWN) {
		Tcl_SetResult (pInterp, "property not found: ", TCL_STATIC);
		Tcl_AppendResult (pInterp, name, NULL);
		return false;
	}

	hr = pcurrent->Invoke (dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, 
		&dispparamsNoArgs, &varResult, &ei, NULL);
	if (hr == DISP_E_EXCEPTION) 
		Tcl_SetResult (pInterp, ExceptInfo2Str (&ei), TCL_DYNAMIC);
	else if (FAILED(hr)) 
		Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
	else
		bOk = true;

	return bOk;
}










/*
 *-------------------------------------------------------------------------
 * ResolvePropertyObject --
 *	Resolves a property list of objects in the dot format
 *	e.g. application.documents(1).pages(2)
 * Result:
 *	true iff successful to bind the ppunk parameter to a valid 
 * Side effects:
 *
 *-------------------------------------------------------------------------
 */
bool	OptclObj::ResolvePropertyObject (Tcl_Interp *pInterp, const char *sname, 
								   IDispatch **ppdisp, ITypeInfo **ppinfo, ITypeComp **ppcmp /* = NULL*/)
{
	USES_CONVERSION;
	ASSERT (pInterp != NULL && ppdisp != NULL && sname != NULL);
	// copy the string onto the stack
	char *		szname;
	char *		seps = ".";
	char *		szprop = NULL;
	_variant_t	varobj;
	VARIANT		varResult;

	HRESULT		hr;
	
	TObjPtr pcmd;
	TObjPtr plist;
	TObjPtr pokstring;

	szname = (char*)_alloca (strlen (sname) + 1);
	strcpy (szname, sname);
	szprop = strtok(szname, seps);
	CComQIPtr <IDispatch> current;
	CComPtr <ITypeInfo> pti;
	CComPtr <ITypeComp> pcmp;

	UINT	typecount = 0;

	current = m_pcurrent;
	pti = m_pti;
	pcmp = m_ptc;

	pcmd.create();

	VariantInit (&varResult);

	try {
		while (szprop != NULL)
		{
			TObjPtr prop(szprop);

			VariantClear(&varResult);
			if (!GetIndexedVariant (pInterp, prop, current, pti, pcmp, varResult))
				break;
			
			// check that it's an object
			if (varResult.vt != VT_DISPATCH && varResult.vt != VT_UNKNOWN)
			{
				Tcl_SetResult (pInterp, "'", TCL_STATIC);
				Tcl_AppendResult (pInterp, szprop, "' is not an object", NULL);
				break;
			}

			else
			{
				current = varResult.punkVal;
				if (current == NULL)
				{
					Tcl_SetResult (pInterp, "'", TCL_STATIC);
					Tcl_AppendResult (pInterp, szprop, "' is not a dispatchable object", NULL);
					break;
				}
				typecount = 0;
				pti = NULL;
				pcmp = NULL;

				current->GetTypeInfoCount (&typecount);
				if (typecount > 0) {
					hr = current->GetTypeInfo (0, LOCALE_SYSTEM_DEFAULT, &pti);
					if (SUCCEEDED(hr)) {
						g_libs.EnsureCached (pti);
					}
					pti->GetTypeComp(&pcmp);
				}
			}
			
			// get the next property
			szprop = strtok(NULL, seps);
		}
		
		*ppinfo = pti.Detach();
		*ppcmp = pcmp.Detach();
		*ppdisp = current.Detach ();
	}

	catch (HRESULT hr)
	{
		Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
	}

	catch (char * error)
	{
		Tcl_SetResult (pInterp, error, TCL_STATIC);
	}
	VariantClear(&varResult);
	return (szprop == NULL);
}





/*
 *-------------------------------------------------------------------------
 * OptclObj::GetBinding --
 *	Retrieves the current binding, if any for a properly formed event.
 *	Event is in the form of either 
 *		'event_name' on default interface
 *		'lib.type.event_name'
 *
 * Result:
 *	true iff successful. Error in interpreter.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
bool OptclObj::GetBinding (Tcl_Interp *pInterp, char *pname)
{
	ASSERT (pInterp != NULL && pname != NULL);

	EventBindings *	pbinding = NULL;
	GUID			guid;
	CComPtr<ITypeInfo> 
					pti;
	int				ne;
	TObjPtr			name(pname);
	TObjPtr			sr;


	// split the name
	sr.create();
	if (!SplitObject(pInterp, name, ".", &sr))
		return false;
	
	ne = sr.llength();
	if (ne <= 0)
		return false;

	// check for a split on strings ending with a token
	if ((*(char*)(sr.lindex(ne - 1))) == '\0') {
		if (--ne <= 0)
			return false;
	}

	if (ne != 1 && ne != 3) {
		Tcl_SetResult (pInterp, "wrong event format: should be either 'eventname' or 'lib.type.eventname'", TCL_STATIC);
		return NULL;
	}

	if (ne == 1 && !FindDefaultEventInterface(pInterp, &pti, &guid))
		return false;
	else if (ne == 3 && FindEventInterface (pInterp, sr.lindex(0), sr.lindex(1), &pti, &guid))
		return false;

	if (m_bindings.find(&guid, &pbinding) == NULL) {
		Tcl_ResetResult (pInterp);
		return true;
	}
	else 
		return pbinding->GetBinding (pInterp, name);
}



/*
 *-------------------------------------------------------------------------
 * OptclObj::SetBinding --
 *	Sets an event binding for the event pointed by 'pname' to the tcl command
 *	stored in 'pcmd'.
 * Result:
 *	true iff successful. Else error in interpreter.
 * Side effects:
 *	Any earlier binding for this event will be removed.
 *-------------------------------------------------------------------------
 */
bool OptclObj::SetBinding (Tcl_Interp *pInterp, char *pname, Tcl_Obj *pcmd)
{
	ASSERT (pInterp != NULL && pname != NULL && pcmd != NULL);
	ASSERT (m_punk != NULL);

	TObjPtr	name(pname);
	TObjPtr	cmd(pcmd, false);

	TObjPtr	sr; // split result
	int		ne;		// name elements
	GUID	guid; // id of the event interface
	HRESULT	hr;

	CComPtr<ITypeInfo> 
			pti;  // typeinfo for the event interface
	
	EventBindings *	// the bindings for this interface
			pbinding = NULL;
	

	// split the name
	sr.create();
	if (!SplitObject(pInterp, name, ".", &sr))
		return false;
	
	ne = sr.llength();
	if (ne <= 0)
		return false;

	// check for a split on strings ending with a token
	if ((*(char*)(sr.lindex(ne - 1))) == '\0') {
		if (--ne <= 0)
			return false;
	}

	if (ne != 1 && ne != 3) {
		Tcl_SetResult (pInterp, "wrong event format: should be either 'eventname' or 'lib.type.eventname'", TCL_STATIC);
		return NULL;
	}

	if (ne == 1 && !FindDefaultEventInterface(pInterp, &pti, &guid))
		return false;
	else if (ne == 3 && !FindEventInterface (pInterp, sr.lindex(0), sr.lindex(1), &pti, &guid))
		return false;

	
	if (m_bindings.find(&guid, &pbinding) == NULL)
	{
		pbinding = new EventBindings (this, guid, pti);
		// initiate the advise
		hr = m_punk.Advise((IUnknown*)(pbinding), guid, &(pbinding->m_cookie));
		if (FAILED(hr)) {
			delete pbinding;
			Tcl_SetResult (pInterp, HRESULT2Str (hr), TCL_DYNAMIC);
			return false;
		}
		m_bindings.set(&guid, pbinding);
	}

	// deleting a single event binding
	if ((*(char*)cmd) == '\0') {
		if (!pbinding->DeleteBinding (pInterp, sr.lindex(ne==1?0:2)))
			return false;
		// total number of bindings for this interface is now zero?
		if (pbinding->TotalBindings() == 0) {
			// unadvise - CComPtr doesn't have this!!
			CComQIPtr<IConnectionPointContainer> pcpc;
			CComPtr<IConnectionPoint> pcp;
			pcpc = m_punk;
			ASSERT (pcpc != NULL);
			pcpc->FindConnectionPoint (guid, &pcp);
			ASSERT (pcp != NULL);
			pcp->Unadvise (pbinding->m_cookie);
			m_bindings.delete_entry(&guid);
			delete pbinding;
		}
	}
	else if (!pbinding->SetBinding (pInterp, sr.lindex(ne==1?0:2), pcmd))
		return false;

	return true;
}




/*
 *-------------------------------------------------------------------------
 * OptclObj::FindDefaultEventInterface --
 *	Retrieves the default event interface for this object.
 * Result:
 *	true iff successful. Else, error string in interpreter.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
bool OptclObj::FindDefaultEventInterface (Tcl_Interp *pInterp, ITypeInfo **ppinfo, GUID *pguid)
{
	ASSERT (pInterp != NULL && ppinfo != NULL && pguid != NULL);
	ASSERT (m_punk != NULL);

	bool		bOk = false;
	TYPEATTR *	pattr = NULL;
	HRESULT		hr;
	USHORT		impltypes;
	CComPtr<ITypeInfo> peti;
	

	if (m_pti_class == NULL) 
		Tcl_SetResult (pInterp, "class-less object doesn't have a default event interface", TCL_STATIC);
	else
	{
		hr = m_pti_class->GetTypeAttr (&pattr);
		CHECKHR_TCL(hr, pInterp, false);
		impltypes = pattr->cImplTypes;
		m_pti_class->ReleaseTypeAttr(pattr); pattr = NULL;

		for (USHORT i = 0; i < impltypes; i++)
		{
			INT flags;
			HREFTYPE href;
			if (SUCCEEDED(m_pti_class->GetImplTypeFlags (i, &flags))
				&& (flags & IMPLTYPEFLAG_FDEFAULT) // default interface and ..
				&& (flags & IMPLTYPEFLAG_FSOURCE) // an event source
				&& SUCCEEDED(m_pti_class->GetRefTypeOfImplType(i, &href))
				&& SUCCEEDED(m_pti_class->GetRefTypeInfo (href, &peti)))
			{
				i = impltypes; // quits this loop
				// while we're here, we'll make sure that this type is cached
				g_libs.EnsureCached(peti);
			}
		}

		if (peti != NULL)
		{
			hr = peti->GetTypeAttr (&pattr);
			CHECKHR_TCL(hr, pInterp, false);
			*pguid = pattr->guid;
			peti->ReleaseTypeAttr(pattr); pattr = NULL;
			*ppinfo = peti;
			(*ppinfo)->AddRef();
			bOk = true;
		}
	}
	return bOk;
}





/*
 *-------------------------------------------------------------------------
 * OptclObj::FindEventInterface --
 *	Called to find an event type info and guid given library and type of 
 *	event interface.
 *
 * Result:
 *	true iff successful. Else, error string in interpreter.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
bool OptclObj::FindEventInterface (Tcl_Interp *pInterp, const char * lib, const char *type,
								ITypeInfo **ppinfo, GUID * pguid)
{
	
	ASSERT (pInterp != NULL && lib != NULL && type != NULL && ppinfo != NULL);
	ASSERT (pguid != NULL);
	ASSERT (m_punk != NULL);

	USES_CONVERSION;
	CComPtr<ITypeInfo>	peti;
	CComPtr<ITypeLib>	petl;

	bool				bOk = false;
	TYPEATTR *			pattr = NULL;
	BSTR				bType = NULL;
	BSTR				bLib = NULL;
	HRESULT				hr;
	UINT				dummy;


	if (m_pti_class == NULL) {
		// we don't have any class information
		// try going through the typelibraries
		try {
			TypeLib_ResolveName (lib, type, NULL, &peti);
			if (peti == NULL)
				Tcl_SetResult (pInterp, "binding through typelib for class-less object failed", TCL_STATIC);
		}

		catch (char *err) {
			Tcl_SetResult (pInterp, err, TCL_VOLATILE);
		}

		catch (HRESULT hr) {
			Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
		}
	}

	else
	{
		// we do have class information
		// this will ensure that even if the type library for the event
		// interface is not loaded (yes, it can be different to that of the
		// object that's using it, it will still be found.

		// convert the event interface name to a bstring
		bType = A2BSTR(type);
		bLib = A2BSTR(lib);
		// get the number of implemented types
		hr = m_pti_class->GetTypeAttr (&pattr);
		CHECKHR_TCL(hr, pInterp, false); // beware, conditional return here
		USHORT types = pattr->cImplTypes;
		m_pti_class->ReleaseTypeAttr (pattr); pattr = NULL;

		// loop throught the implemented types
		for (USHORT intf = 0; intf < types; intf++)
		{
			INT flags;
			HREFTYPE href;
			CComBSTR btypename,
				     blibname;

			// if we the implementation flags is an event source and 
			// the name of the referenced type is the same 
			if (SUCCEEDED(m_pti_class->GetImplTypeFlags (intf, &flags)) 
				&& (flags&IMPLTYPEFLAG_FSOURCE) 
				&& SUCCEEDED(m_pti_class->GetRefTypeOfImplType (intf, &href))
				&& SUCCEEDED(m_pti_class->GetRefTypeInfo (href, &peti))
				&& SUCCEEDED(peti->GetContainingTypeLib(&petl, &dummy))
				&& SUCCEEDED(peti->GetDocumentation(MEMBERID_NIL, &btypename, NULL, NULL, NULL))
				&& SUCCEEDED(petl->GetDocumentation(-1, &blibname, NULL, NULL, NULL)))
			{
				if ((btypename == bType) && (blibname == bLib)) {
					intf = types; // quits this loop
					// while we're at it, lets make sure that this typelibrary is
					// registered.
					g_libs.EnsureCached(petl);
				}
			}
			else {
				peti = NULL;
				petl = NULL;
			}
			btypename.Empty();
			blibname.Empty();
		}

		if (peti == NULL)
			Tcl_SetResult (pInterp, "couldn't find event interface", TCL_STATIC);
	}

	if (peti != NULL) {
		// if we've got a typeinfo, find the GUID
		hr = peti->GetTypeAttr (&pattr);
		CHECKHR_TCL (hr, pInterp, false);
		*pguid = pattr->guid;
		peti->ReleaseTypeAttr(pattr); pattr = NULL;
		*ppinfo = peti;
		(*ppinfo)->AddRef();
		bOk = true;
	}
	return bOk;
}




/*
 *-------------------------------------------------------------------------
 * OptclObj::ReleaseBindingTable --
 *	Releases the bindings withing the event bindings hash table.
 *	It's probably very important that this isn't called within the 
 *	evaluation of an event.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void	OptclObj::ReleaseBindingTable()
{
	CComQIPtr<IConnectionPointContainer> pcpc;
	EventBindings * pbinding;
	CComPtr<IConnectionPoint> pcp;


	pcpc = m_punk;
	if (pcpc == NULL)
		return;

	EventBindingsTbl::iterator i;
	for (i = m_bindings.begin(); i != m_bindings.end(); i++)
	{
		pbinding = *i;
		pcpc->FindConnectionPoint (*(i.key()), &pcp);
		if (pcp != NULL) {
			pcp->Unadvise(pbinding->m_cookie);
			pcp = NULL;
		}
		else {
			// this case occurs when the com object has been destroyed
			// before this object is destroyed
		}
		delete pbinding;
	}
}




/*
 *-------------------------------------------------------------------------
 * OptclObj::GetState --
 *	This is some prelim code for persistence support - yanked out of the 
 *	old container code - more soon!
 *
 * Result:
 *	
 * Side effects:
 *
 *-------------------------------------------------------------------------
 */
bool	OptclObj::GetState (Tcl_Interp *pInterp)
{
	ASSERT (pInterp != NULL);
	USES_CONVERSION;

	CComPtr<IStream>				pStream;
	CComQIPtr<IPersistStream>		pPS;
	CComQIPtr<IPersistStreamInit>	pPSI;

	HGLOBAL		hMem;		// handle to memory
	LPVOID		pMem;		// pointer to memory
	HRESULT		hr;		
	DWORD		dwSize;		// size of memory used
	TObjPtr		pObjs[2];	// objects to automanage the tcl_obj lifetimes
	Tcl_Obj	*	pResult;	// pointers used to create the result list
	CLSID		clsid;
	LPOLESTR	lpOleStr;
	

	pPS = m_punk;
	pPSI = m_punk;
	if (!pPS && !pPSI) {
		Tcl_SetResult (pInterp, "object does not support stream persistance model", TCL_STATIC);
		return false;
	}

	hMem = GlobalAlloc (GHND, 0);
	if (hMem == NULL) {
		Tcl_SetResult (pInterp, "unable to initialise global memory", TCL_STATIC);
		return false;
	}

	hr = CreateStreamOnHGlobal (hMem, TRUE, &pStream);
	if (FAILED(hr)) {
		GlobalFree (hMem);
		Tcl_SetResult (pInterp, "unable to create a stream on global memory", TCL_STATIC);
		return false;
	}

	if (pPS)
		hr = pPS->Save (pStream, TRUE);
	else
		hr = pPSI->Save(pStream, TRUE);
	

	if (FAILED(hr)) {
		if (hr == STG_E_CANTSAVE)
			Tcl_SetResult (pInterp, "failed to save object", TCL_STATIC);
		else if (hr == STG_E_MEDIUMFULL)
			Tcl_SetResult (pInterp, "failed to aquire enough memory", TCL_STATIC);
		return false;
	}
	dwSize = GlobalSize(hMem);
	pMem = GlobalLock (hMem);

	ATLASSERT (pMem);
	pObjs[1] = Tcl_NewStringObj ((char*)pMem, dwSize);
	GlobalUnlock (hMem);

	// now get the clsid
	if (pPS)
		hr = pPS->GetClassID (&clsid);
	else
		hr = pPSI->GetClassID (&clsid);

	if (FAILED(hr)) 
	{
		Tcl_SetResult (pInterp, "failed to retrieve the clsid", TCL_STATIC);
		return false;
	}

	hr = StringFromCLSID (clsid, &lpOleStr);
	if (FAILED(hr)) {
		Tcl_SetResult (pInterp, "failed to convert clsid to a string", TCL_STATIC);
		return false;
	}
	pObjs[0] = Tcl_NewStringObj (OLE2A(lpOleStr), -1);
	pResult = Tcl_NewListObj (0, NULL);
	Tcl_ListObjAppendElement (NULL, pResult, pObjs[0]);
	Tcl_ListObjAppendElement (NULL, pResult, pObjs[1]);
	Tcl_SetObjResult (pInterp, pResult);
	return true;
}


/*
 *-------------------------------------------------------------------------
 * OptclObj::ContainerWantsToDie --
 *	Called by the Tk widget container, when it is about to be destroyed.
 *	If we are not currently destroying this object, then instigate it.
 *
 * Result:
 *	None.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void OptclObj::ContainerWantsToDie ()
{
	if (!m_destroypending)
		g_objmap.Delete(this);
}


