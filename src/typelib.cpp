/*
 *------------------------------------------------------------------------------
 *	typelib.cpp
 *	Implements access to typelibraries. Currently this only includes 
 *	browsing facilities. In the future, this may contain typelib building
 *	functionality.
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
#include "typelib.h"
#include "objmap.h"
#include "optclbindptr.h"
#include "optcltypeattr.h"
#include <strstream>

//----------------------------------------------------------------
//				\/\/\/\/\/\ Declarations /\/\/\/\/\/\/


void		TypeLib_Exit	(ClientData);
const char *TYPEKIND2Str (TYPEKIND tkind);
void		FUNCDESC2Obj (ITypeInfo *pti, FUNCDESC *pfd, TObjPtr &fdesc);
void		VARDESC2Obj (ITypeInfo *pti, VARDESC *pdesc, TObjPtr &presult);
bool		TYPEDESC2Obj (ITypeInfo *pti, TYPEDESC *pdesc, TObjPtr &pobj);

void		VariantToObj (VARIANT *pvar, TObjPtr &obj);

inline void	ReleaseTypeAttr (ITypeInfo *pti, TYPEATTR *&pta);

void		Guid2LibName (GUID &guid, TObjPtr &plibname);

void		TypeLib_GetImplTypes (ITypeInfo *pti, TObjPtr &inherited);
void		TypeLib_ProcessFunctions (ITypeInfo *pti, TObjPtr &methods, TObjPtr &properties);
void		TypeLib_ProcessVariables (ITypeInfo *pti, TObjPtr &properties);
void		TypeLib_GetVariable (ITypeInfo *pti, UINT index, TObjPtr &properties);

HRESULT		BindTypeInfo (ITypeComp *, const char *, ITypeInfo **);

TCL_CMDEF(TypeLib_LoadedLibs);
TCL_CMDEF(TypeLib_LoadLib);
TCL_CMDEF(TypeLib_UnloadLib);
TCL_CMDEF(TypeLib_IsLibLoaded);
TCL_CMDEF(TypeLib_TypesInLib);
TCL_CMDEF(TypeLib_TypeInfo);
TCL_CMDEF(TypeLib_GetRegLibPath);
TCL_CMDEF(TypeLib_GetLoadedLibPath);
TCL_CMDEF(TypeLib_GetDetails);

//// TEST CODE ////
TCL_CMDEF(TypeLib_ResolveConstantTest);

//----------------------------------------------------------------
//			\/\/\/\/\/\/ Globals \/\/\/\/\/\/\/

// this class uses a Tcl hash table - this usually wouldn't be
// safe, except that this hash table is initialised (courtsey of THash<>)
// only on first uses (lazy). So it should be okay. Not sure how 
// this will behave in a multithreaded application

TypeLibsTbl	g_libs;

//----------------------------------------------------------------
// Implementation for TypeLibsTbl class

TypeLibsTbl::TypeLibsTbl () : THash<char, TypeLib*> ()
{

}


TypeLibsTbl::~TypeLibsTbl ()
{
	DeleteAll();
}

void TypeLibsTbl::DeleteAll ()
{
	for (iterator i = begin(); i != end(); i++)
	{
		ASSERT ((*i) != NULL);
		delete (*i);
	}
	deltbl();
}

/*
 *-------------------------------------------------------------------------
 * TypeLibsTbl::LoadLib --
 *	Load a Type Library by it's pathname.
 *
 * Result:
 *	Pointer to the TypeLib object iff successful.
 * Side Effects:
 *	Library is added to the cache
 *-------------------------------------------------------------------------
 */
TypeLib * TypeLibsTbl::LoadLib (Tcl_Interp *pInterp, const char *pathname)
{
	USES_CONVERSION;
	CComPtr<ITypeLib> pLib;
	
	try {
		CComPtr<ITypeComp> pComp;
		TObjPtr progname, fullname;
		Tcl_HashEntry *pEntry = NULL;
		HRESULT hr;

		hr = LoadTypeLibEx(A2OLE(pathname), REGKIND_NONE, &pLib);
		CHECKHR(hr);
		ASSERT(pLib != NULL);
		if (pLib == NULL)
			throw ("failed to bind to a type library");

		return Cache (pInterp, pLib, pathname);
	} 

	catch (char *error) {
		Tcl_SetResult (pInterp, error, TCL_VOLATILE);
	}
	catch (HRESULT hr) {
		Tcl_SetResult (pInterp, (char*)HRESULT2Str(hr), TCL_DYNAMIC);
	}

	return NULL;
}




/*
 *-------------------------------------------------------------------------
 * TypeLibsTbl::Cache --
 *	Called in order to cache a library. 
 *	Pre: The library does *not* exist in the cache
 *
 * Result:
 *	A standard OLE HRESULT
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
TypeLib* TypeLibsTbl::Cache (Tcl_Interp *pInterp, ITypeLib *ptl, const char * path /* = NULL */)
{
	TLIBATTR * ptlattr = NULL;
	CComPtr<ITypeComp> ptc;
	TypeLib *pLib = NULL;
	Tcl_HashEntry *pEntry = NULL;


	if (FAILED(ptl->GetLibAttr(&ptlattr))) {
		if (pInterp)
			Tcl_SetResult (pInterp, "couldn't retrieve type library attributes", TCL_STATIC);
		return NULL;
	}

	ASSERT(ptlattr != NULL);

	TypeLibUniqueID uid (ptlattr->guid, ptlattr->wMajorVerNum, ptlattr->wMinorVerNum);


	ptl->ReleaseTLibAttr(ptlattr);
	
	// search for this guid
	if (m_loadedlibs.find(&uid, &pEntry) != NULL) {
		ASSERT (pEntry != NULL);
		pLib = (TypeLib *)Tcl_GetHashValue (pEntry);
		return pLib;
	} 

	// now generate the names, and do a search on the programmatic name
	TObjPtr progname, fullname;
	GenerateNames(progname, fullname, ptl);

	if (g_libs.find((char*)progname, &pLib) != NULL) {
		if (pInterp)
			Tcl_SetResult (pInterp, "library already loaded with the same programmatic name", TCL_STATIC);
		return NULL;
	} 

	if (FAILED(ptl->GetTypeComp(&ptc))) {
		if (pInterp)
			Tcl_SetResult (pInterp, "failed to retrieve the ITypeComp interface", TCL_STATIC);
		return NULL;
	}

	pLib = new TypeLib ();
	if (FAILED(pLib->Init(ptl, ptc, progname, fullname, path))) {
		delete pLib;
		pLib = NULL;
	} else {
		pEntry = set(progname, pLib);
		ASSERT (pEntry != NULL);
		m_loadedlibs.set (&uid, pEntry);
	}

	return pLib;
}


TypeLib * TypeLibsTbl::TypeLibFromUID (const GUID &guid, WORD maj, WORD min)
{
	TypeLibUniqueID uid(guid, maj, min);
	TypeLib *plib = NULL;
	Tcl_HashEntry *pEntry = NULL;
	m_loadedlibs.find(&uid, &pEntry);
	if (pEntry)
		plib = (TypeLib*)Tcl_GetHashValue (pEntry);
	return plib;
}


char * TypeLibsTbl::GetFullName (char * szProgName)
{
	ASSERT (szProgName != NULL);
	TypeLib * pLib = NULL;
	char * result = NULL;
	if (find(szProgName, &pLib) != end()) {
		ASSERT (pLib != NULL);
		result = pLib->m_fullname;
	}
	return result;
}



GUID * TypeLibsTbl::GetGUID (char * szProgName)
{
	ASSERT (szProgName != NULL);
	TypeLib * pLib = NULL;
	GUID * result = NULL;
	if (find(szProgName, &pLib) != end()) {
		ASSERT (pLib != NULL && pLib->m_libattr != NULL);
		result = &(pLib->m_libattr->guid);
	}
	return result;
}


/*
 *-------------------------------------------------------------------------
 * TypeLibsTbl::UnloadLib --
 *	Given the programmatic name of a library, the routine unloads it, if it is 
 *	loaded.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void TypeLibsTbl::UnloadLib (Tcl_Interp *pInterp, const char *szprogname)
{
	Tcl_HashEntry *pEntry = NULL;
	TypeLib *ptl  = NULL;
	pEntry = g_libs.find(szprogname, &ptl);

	if (pEntry == NULL) 
		return;

	ASSERT (ptl != NULL && ptl->m_ptl != NULL);

	TObjPtr progname, fullname;
	HRESULT hr = GenerateNames(progname, fullname, ptl->m_ptl);

	if (FAILED(hr)) {
		Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
		return;
	}
	ASSERT (fullname != (Tcl_Obj*)(NULL));
	TypeLibUniqueID uid (ptl->m_libattr->guid, ptl->m_libattr->wMajorVerNum, ptl->m_libattr->wMinorVerNum);
	m_loadedlibs.delete_entry(&(uid));
	delete ptl;
	Tcl_DeleteHashEntry (pEntry);
}

/*
 *-------------------------------------------------------------------------
 * TypeLibsTbl::GenerateNames --
 *	Given a type library, generates the programmatic and full name for the
 *	library.
 *
 * Result:
 *	S_OK iff successful.
 *
 * Side Effects:
 *	The objects, progname and username allocate memory to store the 
 *	names.
 *-------------------------------------------------------------------------
 */
HRESULT TypeLibsTbl::GenerateNames (TObjPtr &progname, TObjPtr &username, ITypeLib *pLib)
{
	USES_CONVERSION;
	ASSERT (pLib != NULL);
	CComBSTR bprogname, busername;
	HRESULT hr;
	hr = pLib->GetDocumentation(-1, &bprogname, &busername, NULL, NULL);
	if (FAILED(hr)) return hr;

	TLIBATTR * pattr = NULL;
	hr = pLib->GetLibAttr (&pattr);
	if (FAILED(hr)) return hr;

	ASSERT (pattr != NULL);
	TDString str;
	if (busername != NULL)
		str << OLE2A(busername);
	else
		str << OLE2A(bprogname);
	str << " (Ver " << pattr->wMajorVerNum << "." << pattr->wMinorVerNum << ")";
	pLib->ReleaseTLibAttr(pattr);

	username.create();
	username = str;
	progname.create();
	progname = OLE2A(bprogname);
	return hr;
}



/*
 *-------------------------------------------------------------------------
 * TypeLibsTbl::EnsureCached --
 *	Given a typelibrary, the routine ensures that it is stored in the cache.
 *
 * Result:
 *	A pointer to the caches TypeLib object.
 *
 * Side effects:
 *	Throws HRESULT.
 *-------------------------------------------------------------------------
 */
TypeLib *TypeLibsTbl::EnsureCached (ITypeLib  *ptl)
{
	return Cache(NULL, ptl);
}


/*
 *-------------------------------------------------------------------------
 * TypeLibsTbl::EnsureCached --
 *	Sames as EnsureChached(ITypeLib *), but uses a type info.
 *
 * Result:
 *	Non NULL iff successful - result points to the cached TypeLib structure.
 *
 * Side effects:
 *	Throws HRESULT.
 *-------------------------------------------------------------------------
 */
TypeLib *TypeLibsTbl::EnsureCached (ITypeInfo *pInfo)
{
	ASSERT (pInfo != NULL);
	CComPtr<ITypeLib> pLib;
	UINT tmp;
	HRESULT hr;
	hr = pInfo->GetContainingTypeLib(&pLib, &tmp);
	if (FAILED(hr)) return NULL;
	return EnsureCached (pLib);
}







// ------------------- TypeLib initialisation and shutdown routines -------------------------
int TypeLib_Init (Tcl_Interp *pInterp)
{
	OleInitialize(NULL);
	Tcl_CreateExitHandler (TypeLib_Exit, NULL);
	Tcl_CreateObjCommand (pInterp, "typelib::loadedlibs", TypeLib_LoadedLibs, NULL, NULL);
	Tcl_CreateObjCommand (pInterp, "typelib::_load", TypeLib_LoadLib, NULL, NULL);
	Tcl_CreateObjCommand (pInterp, "typelib::unload", TypeLib_UnloadLib, NULL, NULL);
	Tcl_CreateObjCommand (pInterp, "typelib::types", TypeLib_TypesInLib, NULL, NULL);
	Tcl_CreateObjCommand (pInterp, "typelib::typeinfo", TypeLib_TypeInfo, NULL, NULL);
	Tcl_CreateObjCommand (pInterp, "typelib::isloaded", TypeLib_IsLibLoaded, NULL, NULL);
	Tcl_CreateObjCommand (pInterp, "typelib::reglib_path", TypeLib_GetRegLibPath, NULL, NULL);
	Tcl_CreateObjCommand (pInterp, "typelib::loadedlib_path", TypeLib_GetLoadedLibPath, NULL, NULL);
	Tcl_CreateObjCommand (pInterp, "typelib::loadedlib_details", TypeLib_GetDetails, NULL, NULL);

	
	//// TESTS ////
	Tcl_CreateObjCommand (pInterp, "typelib::resolveconst", TypeLib_ResolveConstantTest, NULL, NULL);
	
	return TCL_OK;
}



void TypeLib_Exit (ClientData)
{
	g_libs.DeleteAll ();
	OleUninitialize();
}
// ------------------------------------------------------------------------------------------




/*
 *-------------------------------------------------------------------------
 * ReleaseTypeAttr --
 *	Release at type attribute from the specified type info.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
inline void	ReleaseTypeAttr (ITypeInfo *pti, TYPEATTR *&pta)
{	
	ASSERT (pti != NULL);
	if (pta != NULL) {
		pti->ReleaseTypeAttr(pta);
		pta = NULL;
	}
}




/*
 *-------------------------------------------------------------------------
 * ReleaseBindPtr --
 *	Releases a bind ptr (if not null), according to its type description.
 *	Sets the value of the pointer to null.
 *
 * Result:
 *	None
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void	ReleaseBindPtr (ITypeInfo *pti, DESCKIND dk, BINDPTR &ptr)
{
	if (ptr.lpfuncdesc != NULL) {
		switch (dk) {
		case DESCKIND_FUNCDESC:
			ASSERT (pti != NULL);
			pti->ReleaseFuncDesc (ptr.lpfuncdesc);
			ptr.lpfuncdesc = NULL;
			break;
		case DESCKIND_IMPLICITAPPOBJ: // same as a vardesc
		case DESCKIND_VARDESC:
			ASSERT (pti != NULL);
			pti->ReleaseVarDesc (ptr.lpvardesc);
			ptr.lpvardesc = NULL;
			break;
		case DESCKIND_TYPECOMP:
			ptr.lptcomp->Release();
			ptr.lptcomp = NULL;
			break;
		}
	}
}



const char *TYPEKIND2Str (TYPEKIND tkind)
{
	switch (tkind)
	{
	case TKIND_ENUM:
		return "enum";
	case TKIND_RECORD:
		return "struct";
	case TKIND_MODULE:
		return "module";
	case TKIND_INTERFACE:
		return "interface";
	case TKIND_DISPATCH:
		return "dispatch";
	case TKIND_COCLASS:
		return "class";
	case TKIND_ALIAS:
		return "typedef";
	case TKIND_UNION:
		return "union";
	default:
		return "???";
	}
}






const char *VARTYPE2Str (VARTYPE vt)
{
	vt = vt & ~VT_ARRAY & ~VT_BYREF;
	switch (vt) {
	case VT_EMPTY:
	case VT_NULL:
		return "_null_";
	case VT_I1:
		return "char";
	case VT_UI1:
		return "uchar";
	case VT_I2:
		return "short";
	case VT_UI2:
		return "ushort";
	case VT_INT:
	case VT_I4:
	case VT_ERROR:
		return "long";
	case VT_UI4:
	case VT_UINT:
		return "ulong";
	case VT_I8:
		return "super_long";
	case VT_UI8:
		return "usuper_long";
	case VT_R4:
		return "float";
	case VT_R8:
		return "double";
	case VT_CY:
		return "currency";
	case VT_DATE:
		return "date";
	case VT_BSTR:
		return "string";
	case VT_DISPATCH:
		return "dispatch";
	case VT_BOOL:
		return "bool";
	case VT_VARIANT:
		return "any";
	case VT_UNKNOWN:
		return "interface";
	case VT_DECIMAL:
		return "decimal";
	case VT_VOID:
		return "void";
	case VT_HRESULT:
		return "scode";
	case VT_LPSTR:
	case VT_LPWSTR:
		return "string";
	case VT_CARRAY:
		return "carray";
	default:
		return "???";
	}
}





/*
 *-------------------------------------------------------------------------
 * TYPEDESC2Obj --
 *	
 * Result:
 *
 * Side effects:
 *
 *-------------------------------------------------------------------------
 */
bool TYPEDESC2Obj (ITypeInfo *pti, TYPEDESC *pdesc, TObjPtr &pobj)
{
	USES_CONVERSION;
	ASSERT (pdesc != NULL && pti != NULL);
	bool array = ((pdesc->vt & VT_ARRAY) != 0);
	pdesc->vt = pdesc->vt & ~VT_ARRAY;
	HRESULT hr;

	if (pdesc->vt == VT_USERDEFINED) {
		// resolve the referenced type
		CComPtr<ITypeInfo> prefti;
		TYPEATTR *pta = NULL;
		WORD flags; 
		hr = pti->GetRefTypeInfo (pdesc->hreftype, &prefti);
		CHECKHR(hr);
		hr = prefti->GetTypeAttr (&pta);
		CHECKHR(hr);
		flags = pta->wTypeFlags;

		ReleaseTypeAttr (prefti, pta);
		if ((flags & TYPEFLAG_FRESTRICTED)) {
			pobj.create();
			pobj = "!!!"; // unaccessable type
			return false;
		}
		g_libs.EnsureCached(prefti);
		TypeLib_GetName (NULL, prefti, pobj);
	} else if ((pdesc->vt == VT_SAFEARRAY) || (pdesc->vt == VT_PTR)) {
		if (!TYPEDESC2Obj (pti, pdesc->lptdesc, pobj))
			return false;
		ASSERT (pobj.isnotnull());
		if (pdesc->vt == VT_SAFEARRAY)
			pobj += " []";
	} else {
		pobj.create();
		pobj = VARTYPE2Str(pdesc->vt);
	}

	return true;
}





/*
 *-------------------------------------------------------------------------
 * TypeLib_GetName --
 *	Converts a library or type to a name stored in a Tcl_Obj. If pLib
 *	is not NULL and pInfo is, then the name is the name of the library.
 *	Otherwise, pInfo must be non-null (pLib can always be derived from pInfo)
 *	The result is stored in pname.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void	TypeLib_GetName (ITypeLib *pLib, ITypeInfo *pInfo, TObjPtr &pname)
{
	ASSERT (pLib!=NULL || pInfo!=NULL);

	USES_CONVERSION;
	BSTR	progname = NULL,
			typname = NULL;
	HRESULT hr;

	UINT tmp;
	bool bLibcreate = false;

	// ensure we have a library to work with
	if (pLib == NULL) {
		hr = pInfo->GetContainingTypeLib(&pLib, &tmp);
		CHECKHR(hr);
		bLibcreate = true;
	}
	// get the library programmatic name

	
	hr = pLib->GetDocumentation (-1, &progname, NULL, NULL, NULL);
	CHECKHR(hr);

	if (pInfo == NULL) {
		pname.create(); 
		pname = W2A(progname);
	} else {
		hr = pInfo->GetDocumentation(MEMBERID_NIL, &typname, NULL, NULL, NULL);
		CHECKHR(hr);
		TDString str;
		str.set(W2A(progname)) << "." << W2A(typname);
		pname.create(); 
		pname = str;
	}

	FreeBSTR(progname);
	FreeBSTR(typname);
	if (bLibcreate)
		pLib->Release();
}





/*
 *-------------------------------------------------------------------------
 * BindTypeInfo --
 *	Given a type compiling interface (ptc) and a typename (szTypeName),
 *	resolves to a ITypeInfo interface (stored in ppti).
 *
 * Result:
 *	Returns a standard OLE HRESULT.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
HRESULT BindTypeInfo (ITypeComp *ptc, const char *szTypeName, ITypeInfo **ppti)
{
	USES_CONVERSION;
	LPOLESTR oleTypename = NULL;
	UINT hash;
	ASSERT (ptc != NULL && ppti != NULL);
	CComPtr<ITypeComp> ptemp;
	oleTypename  = A2OLE(szTypeName);
	hash = LHashValOfName(LOCALE_SYSTEM_DEFAULT, oleTypename);
	return ptc->BindType (oleTypename, hash, ppti, &ptemp);
}






/*
 *-------------------------------------------------------------------------
 * FUNCDESC2Obj --
 *	
 * Result:
 *
 * Side effects:
 *
 *-------------------------------------------------------------------------
 */
void FUNCDESC2Obj (ITypeInfo *pti, FUNCDESC *pfd, TObjPtr &fdesc)
{
	ASSERT (pfd != NULL && pti != NULL);
	ASSERT (!(pfd->wFuncFlags & FUNCFLAG_FRESTRICTED));
	USES_CONVERSION;
	BSTR	*	fnames = NULL;
	char *		szfname = NULL;
	HRESULT		hr;
	UINT		totalread = 0;
	UINT		total = 0;
	TObjPtr		type;
	TObjPtr		flags;
	TObjPtr		param;
	TObjPtr		optionparam;

	fdesc.create();

	try {
		// get the names
		total = pfd->cParams + 1;
		fnames = new BSTR[total];
		hr = pti->GetNames(pfd->memid, fnames, total, &totalread);
		CHECKHR(hr);
		if (totalread != total)
			throw ("couldn't retrieve all the parameter names");

		
		
		TYPEDESC2Obj(pti, &(pfd->elemdescFunc.tdesc), type);
		fdesc.lappend (type); // return type
		fdesc.lappend (W2A(fnames[0])); // the function name
		
		
		// now build up the parameters
		for (SHORT index = 0; index < pfd->cParams; index++)
		{
			ELEMDESC *pdesc = pfd->lprgelemdescParam + index;
			flags.create();
			if ((pdesc->paramdesc.wParamFlags & PARAMFLAG_FIN) ||
				(pdesc->paramdesc.wParamFlags == PARAMFLAG_NONE))
				flags.lappend("in");
			if (pdesc->paramdesc.wParamFlags & PARAMFLAG_FOUT)
				flags.lappend("out");
			if (pdesc->paramdesc.wParamFlags & PARAMFLAG_FRETVAL)
				flags.lappend("retval");

			// type of parameter
			TYPEDESC2Obj(pti, &(pdesc->tdesc), type);

			// setup the result
			param.create();
			param.lappend(flags).lappend(type).lappend(W2A(fnames[index+1]));

			
			// is it optional and does it have a default value
			if ((pdesc->paramdesc.wParamFlags & PARAMFLAG_FHASDEFAULT)
				&& (pdesc->paramdesc.wParamFlags & PARAMFLAG_FOPT)) {
				VariantToObj (&(pdesc->paramdesc.pparamdescex->varDefaultValue), optionparam);
				param.lappend(optionparam);
			}
			else
			if ((pfd->cParams - index)<=pfd->cParamsOpt) 
				param.lappend ("?");
			
			fdesc.lappend(param);
		}


		FreeBSTRArray (fnames, totalread);
		delete fnames;
	}

	catch (char *error) {
		FreeBSTRArray (fnames, totalread);
		delete fnames;
		throw (error);
	}

	catch (HRESULT hr) {
		FreeBSTRArray (fnames, totalread);
		delete fnames;
		throw (hr);
	}
}



/*
 *-------------------------------------------------------------------------
 * VARDESC2Obj --
 *	
 * Result:
 *
 * Side effects:
 *
 *-------------------------------------------------------------------------
 */
void VARDESC2Obj (ITypeInfo *pti, VARDESC *pdesc, TObjPtr &presult)
{
	ASSERT (pti != NULL && pdesc != NULL);

	USES_CONVERSION;
	HRESULT			hr;
	BSTR			name = NULL;
	char		*	szname = NULL;
	TObjPtr			tdesc; // stores the description of the type
	TObjPtr			tflags;// read write flags for this variable
	
	hr = pti->GetDocumentation(pdesc->memid, &name, NULL, NULL, NULL);
	CHECKHR(hr);
	szname = W2A (name);
	FreeBSTR (name);

	TYPEDESC2Obj (pti, &(pdesc->elemdescVar.tdesc), tdesc);
	tflags.create();
	if (pdesc->wVarFlags & VARFLAG_FREADONLY)
		tflags = "read";
	else
		tflags = "read write";
	
	presult.create();
	presult.lappend (tflags).lappend(tdesc).lappend(szname);
	if (pdesc->varkind == VAR_CONST) { // its a constant
		TObjPtr cnst;
		VariantToObj (pdesc->lpvarValue, cnst);
		presult.lappend (cnst);
	}
}


void VariantToObj (VARIANT *pvar, TObjPtr &obj)
{
	ASSERT (pvar != NULL);

	USES_CONVERSION;

	VARTYPE vt = pvar->vt;
	CComVariant var;
	HRESULT hr;

	vt = vt & ~VT_BYREF;
	obj.create();

	if (vt == VT_UNKNOWN || vt == VT_DISPATCH) 
		obj = "object";
	else if ((vt & VT_ARRAY) == VT_ARRAY) 
		obj = "array";
	else {
		hr = var.Copy (pvar);
		CHECKHR(hr);
		var.ChangeType(VT_BSTR);
		ASSERT (var.bstrVal != NULL);
		obj = W2A (var.bstrVal);
	}
}



/*
 *-------------------------------------------------------------------------
 * ImplFlags2Obj --
 *
 *	Converts implementation flags to a tcl object.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	Uses TObjPtr functions, and hence throws (char *) in case of any errors.
 *-------------------------------------------------------------------------
 */
void ImplFlags2Obj (UINT implflags, TObjPtr &flags)
{
	flags.create();
	if (implflags & IMPLTYPEFLAG_FDEFAULT)
		flags.lappend("default");
	if (implflags & IMPLTYPEFLAG_FSOURCE)
		flags.lappend("source");
}




/*
 *-------------------------------------------------------------------------
 * TypeLib_DescOfRefType --
 *
 *	Called to describe a referenced type from another type. If bclassinfo 
 *	is true, the function prepends additional flags to describe the role of
 *	the referenced type to the class type.
 *
 * Result:
 *	return true iff successful.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
bool TypeLib_DescOfRefType (ITypeInfo *pti, UINT index, TObjPtr &desc, bool bclassinf)
{
	ASSERT (pti != NULL);

	HRESULT hr;
	TObjPtr name;
	TObjPtr flags;
	CComPtr <ITypeInfo>	ptmp;
	HREFTYPE href;
	INT implflags;
	TYPEATTR * pta = NULL;
	WORD typeflags;

	hr = pti->GetRefTypeOfImplType (index , &href);
	CHECKHR(hr);

	hr = pti->GetRefTypeInfo (href, &ptmp);
	CHECKHR(hr);
	
	g_libs.EnsureCached(ptmp);
	hr = pti->GetImplTypeFlags(index, &implflags);
	CHECKHR(hr);

	hr = ptmp->GetTypeAttr (&pta);
	CHECKHR(hr);
	typeflags = pta->wTypeFlags;
	ReleaseTypeAttr(pti, pta);

	if ((typeflags & TYPEFLAG_FRESTRICTED) || 
		(implflags & IMPLTYPEFLAG_FRESTRICTED))
		return false;


	TypeLib_GetName (NULL, ptmp, name);
	if (bclassinf) {
		ImplFlags2Obj (implflags, flags);
	} else {
		flags.create();
	}

	desc.create();
	desc.lappend(flags).lappend(name);
	return true;
}





/*
 *-------------------------------------------------------------------------
 * TypeLib_GetImplTypes --	
 *	Compiles a list of inherited interfaces from a type information pointer.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	throws char * and HRESULT.
 *-------------------------------------------------------------------------
 */
void TypeLib_GetImplTypes (ITypeInfo *pti, TObjPtr &inherited)
{
	ASSERT (pti!=NULL && inherited.isnotnull());

	HRESULT		hr;
	TYPEATTR *	pattr = NULL;
	WORD		count;		// total number of references
	TObjPtr		desc;
	TYPEKIND	tkind;


	hr = pti->GetTypeAttr (&pattr);
	CHECKHR(hr);

	count = pattr->cImplTypes;
	tkind = pattr->typekind;
	ReleaseTypeAttr (pti, pattr);
	
	for (WORD index = 0; index < count; index++)
	{
		if (TypeLib_DescOfRefType (pti, index, desc, (tkind == TKIND_COCLASS)))
			inherited.lappend(desc);
	}
}







/*
 *-------------------------------------------------------------------------
 * TypeLib_GetVariable --
 *	Gets the description for a variable (VARDESC) property, based on an 
 *	index.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	Throws HRESULT and char *.
 *-------------------------------------------------------------------------
 */
void TypeLib_GetVariable (ITypeInfo *pti, UINT index, TObjPtr &properties)
{
	USES_CONVERSION;

	ASSERT (pti != NULL && properties.isnotnull());
	VARDESC *	pDesc;
	HRESULT		hr;
	BSTR		name = NULL;
	char	*	szName = NULL;

	try {
		hr = pti->GetVarDesc(index, &pDesc);
		CHECKHR(hr);

		ASSERT (pDesc != NULL);

		if (!(pDesc->wVarFlags & VARFLAG_FHIDDEN)) // not a hidden variable
		{
			hr = pti->GetDocumentation(pDesc->memid, &name, NULL, NULL, NULL);
			CHECKHR(hr);

			szName = W2A(name);
			FreeBSTR(name);
			properties.lappend(szName);
		}
		pti->ReleaseVarDesc(pDesc);
	}

	catch (HRESULT hr) {
		if (pDesc != NULL) {
			pti->ReleaseVarDesc(pDesc);
			pDesc = NULL;
		}
		throw (hr);
	}

	catch (char *error) {
		if (pDesc != NULL) {
			pti->ReleaseVarDesc(pDesc);
			pDesc = NULL;
		}
		throw (error);
	}
}







/*
 *-------------------------------------------------------------------------
 * TypeLib_ProcessVariables --
 *	Appends to a tcl list object, a the set of VARDESC defined properties.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	Uses functions that throw HRESULT and char *
 *-------------------------------------------------------------------------
 */
void TypeLib_ProcessVariables (ITypeInfo *pti, TObjPtr &properties)
{
	ASSERT (pti != NULL && properties.isnotnull());
	TYPEATTR *pattr = NULL;
	HRESULT hr;
	UINT	count;

	hr = pti->GetTypeAttr(&pattr);
	CHECKHR(hr);
	count = pattr->cVars;
	ReleaseTypeAttr (pti, pattr);
	
	for (UINT index = 0; index < count; index++) {
		TypeLib_GetVariable (pti, index, properties);
	}
}




/*
 *-------------------------------------------------------------------------
 * TypeLib_ProcessFunctions --
 *	Scans the functions within the type and separates them into two lists:
 *	Methods functions, and property functions. Read/Write access to 
 *	properties are not determined here.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	Can throw HRESULT and char *.
 *-------------------------------------------------------------------------
 */
void	TypeLib_ProcessFunctions (ITypeInfo *pti, TObjPtr &methods, TObjPtr &properties)
{
	USES_CONVERSION;
	ASSERT (pti != NULL && properties.isnotnull() && methods.isnotnull());

	HRESULT				hr;
	TYPEATTR	*		pattr = NULL;
	WORD				count;
	FUNCDESC	*		pfd = NULL;
	BSTR				name;
	char		*		szname;
	THash<char, int>	proptbl;

	hr = pti->GetTypeAttr (&pattr);
	CHECKHR(hr);

	count = pattr->cFuncs;
	ReleaseTypeAttr (pti, pattr);

	
	try {
		for (WORD index = 0; index < count; index++)
		{
			hr = pti->GetFuncDesc (index, &pfd);
			CHECKHR(hr);
			// if the function shouldn't be shown, skip this iteration
			if ((pfd->wFuncFlags & FUNCFLAG_FRESTRICTED)) {
				pti->ReleaseFuncDesc (pfd); pfd = NULL;
				continue;
			}

			hr = pti->GetDocumentation(pfd->memid, &name, NULL, NULL, NULL);
			CHECKHR(hr);

			szname = W2A (name);
			FreeBSTR (name);
			if (pfd->invkind == INVOKE_FUNC) {
				methods.lappend(szname);
			} else {
				proptbl.set(szname, 0);
			}
			pti->ReleaseFuncDesc (pfd); pfd = NULL;
		}
		// now process the properties 
		for (THash<char,int>::iterator e = proptbl.begin(); e != proptbl.end(); e++)
			properties.lappend(e.key());
	}

	catch (char *error) {
		if (pfd != NULL) {
			pti->ReleaseFuncDesc (pfd); pfd = NULL;
		}
		throw (error);
	}
	catch (HRESULT hr) {
		if (pfd != NULL) {
			pti->ReleaseFuncDesc (pfd); pfd = NULL;
		}
		throw (hr);
	}
}






/*
 *-------------------------------------------------------------------------
 * TypeLib_DescribeTypeInfo --
 *	Describe the type info in terms of the types kind, its methods, 
 *	properties and inherited types.
 *
 * Result:
 *	Std Tcl return.
 *
 * Side effects:
 *	Can throw either HRESULT and char *. These probably can be removed 
 *	and returned directly in the interpreter. I've left them to be 
 *	picked up by the calling procedure, as that has its own exception
 *	handling code.
 *-------------------------------------------------------------------------
 */
int TypeLib_DescribeTypeInfo (Tcl_Interp *pInterp, ITypeInfo *pti)
{
	int cmdresult = TCL_ERROR;
	USES_CONVERSION;

	ASSERT (pti != NULL && pInterp != NULL);
	TYPEATTR *pta = NULL;
	HRESULT hr;
	TObjPtr presult,
			inherited,
			methods,
			properties;
	BSTR	bdoc = NULL;

	hr = pti->GetTypeAttr(&pta);
	CHECKHR(hr);

	try {
		if (pta->typekind == TKIND_ALIAS) {
			presult.create ();
			presult.lappend("typedef").lappend("").lappend("");
			
			//TypeLib_GetImplTypes (pti, inherited);
			TYPEDESC2Obj (pti, &(pta->tdescAlias), inherited);
			presult.lappend (inherited);
			cmdresult = TCL_OK;
		} 
		
		else {
			inherited.create();
			methods.create();
			properties.create();
			TypeLib_GetImplTypes (pti, inherited);
			TypeLib_ProcessFunctions (pti, methods, properties);
			TypeLib_ProcessVariables (pti, properties);

			presult.create();
			switch (pta->typekind)
			{
			case TKIND_ENUM:
				presult = "enum"; break;
			case TKIND_RECORD:
				presult = "struct"; break;
			case TKIND_MODULE:
				presult = "module"; break;
			case TKIND_INTERFACE:
				presult = "interface"; break;
			case TKIND_DISPATCH:
				presult = "dispatch"; break;
			case TKIND_COCLASS:
				presult = "class"; break;
			case TKIND_UNION:
				presult = "union"; break;
			default:
				presult = "???"; break;
			}

			
			presult.lappend(methods).lappend(properties).lappend(inherited);

			
			if (SUCCEEDED(pti->GetDocumentation (MEMBERID_NIL, NULL, &bdoc, NULL, NULL)) && bdoc != NULL)
			{
				presult.lappend (OLE2A(bdoc));
				SysFreeString (bdoc);
			}
			else
				presult.lappend ("");

			LPOLESTR lpsz;
			CHECKHR(StringFromCLSID (pta->guid, &lpsz));
			ASSERT (lpsz != NULL);
			if (lpsz != NULL) {
				presult.lappend(OLE2A (lpsz));
				CoTaskMemFree (lpsz); lpsz = NULL;
			}
		}
		Tcl_SetObjResult (pInterp, presult);
		cmdresult = TCL_OK;

		ReleaseTypeAttr (pti, pta);
	}
	catch (HRESULT hr) {
		ReleaseTypeAttr (pti, pta);
		throw (hr);
	}
	catch (char *error) {
		ReleaseTypeAttr (pti, pta);
		throw (error);
	}

	return cmdresult;
}




/*
 *-------------------------------------------------------------------------
 * DescPropertyFuncDesc --
 *	Helper function to provides a description in a tcl object, 
 *	of a accessor based property. The property name, hash, typeinfo,
 *	compiler interface, funcdesc are already provided. The function evaluates
 *	the read/write priviliges, and type of the property, before building the
 *	resultant list.
 *	
 * Result:
 *	None.
 *
 * Side effects:
 *	throws HRESULT or char*
 *-------------------------------------------------------------------------
 */
void DescPropertyFuncDesc (BSTR name, ULONG hash, ITypeInfo *pti, 
						   ITypeComp *pcmp, FUNCDESC *pfd, TObjPtr &pdesc)
{
	ASSERT (pti != NULL && pcmp != NULL);

	USES_CONVERSION;

	bool				bRead = false,
						bWrite = false;
	BSTR	*			fnames = NULL;
	char			*	szname = NULL;
	OptclBindPtr		obp;
	HRESULT				hr;
	
	UINT		totalread = 0;
	UINT		total = 0;
	TObjPtr		fdesc, param, type, optionparam, flags;

	try {
		// find out read/write access of this property
		bWrite = (pfd->invkind==INVOKE_PROPERTYPUT ||
				  pfd->invkind==INVOKE_PROPERTYPUTREF);

		// assertion: due to the order of computation,
		//			  if bWrite is TRUE, then bRead will be false
		bRead = !bWrite;

		if (!bWrite) {
			hr = pcmp->Bind (name, hash, INVOKE_PROPERTYPUT | INVOKE_PROPERTYPUTREF, 
							 &obp.m_pti, &obp.m_dk, &obp.m_bp);
			bWrite = SUCCEEDED(hr) && (obp.m_bp.lpfuncdesc->invkind==INVOKE_PROPERTYPUT ||
					  obp.m_bp.lpfuncdesc->invkind==INVOKE_PROPERTYPUTREF);
		}

		total = pfd->cParams + 1;
		fnames = new BSTR[total];
		hr = pti->GetNames(pfd->memid, fnames, total, &totalread);
		CHECKHR(hr);
		if (totalread != total)
			throw ("couldn't retrieve all the parameter names");

		pdesc.create();
		flags.create();
		if (bRead)
			flags.lappend ("read");
		if (bWrite)
			flags.lappend ("write");
		if (bRead) { // its a propertyget - use the return value of the function as the type
			TYPEDESC2Obj (pti, &(pfd->elemdescFunc.tdesc), type);
		} else { // its a propertyput only - use the first parameter
			TYPEDESC2Obj (pti, &(pfd->lprgelemdescParam->tdesc), type);
		}
		pdesc.lappend(flags).lappend(type).lappend(W2A(fnames[0])); 


		// now build up the parameters
		for (SHORT index = 0; index < pfd->cParams; index++)
		{
			ELEMDESC *elemdesc = pfd->lprgelemdescParam + index;
			
			flags.create();
			if (elemdesc->paramdesc.wParamFlags & PARAMFLAG_FIN)
				flags.lappend("in");
			if (elemdesc->paramdesc.wParamFlags & PARAMFLAG_FOUT)
				flags.lappend("out");
			if (elemdesc->paramdesc.wParamFlags & PARAMFLAG_FRETVAL)
				flags.lappend("retval");
			
			// type of parameter
			TYPEDESC2Obj(pti, &(elemdesc->tdesc), type);

			// setup the result
			param.create();
			param.lappend(flags).lappend(type).lappend(W2A(fnames[index+1]));

			// is it optional and does it have a default value
			if ((elemdesc->paramdesc.wParamFlags & PARAMFLAG_FHASDEFAULT)
				&& (elemdesc->paramdesc.wParamFlags & PARAMFLAG_FOPT)) {
				VariantToObj (&(elemdesc->paramdesc.pparamdescex->varDefaultValue), optionparam);
				param.lappend(optionparam);
			}
			else
			if ((pfd->cParams - index)<=pfd->cParamsOpt) 
				param.lappend ("?");
			
			pdesc.lappend(param);
		}

		FreeBSTRArray (fnames, totalread);
		delete fnames;
	}
	catch (char *error) {
		FreeBSTRArray (fnames, totalread);
		delete fnames;
		throw (error);
	}

	catch (HRESULT hr) {
		FreeBSTRArray (fnames, totalread);
		delete fnames;
		throw (hr);
	}
}





/*
 *-------------------------------------------------------------------------
 * TypeLib_DescribeTypeInfoElement --
 *	Called to describe an element of a type information pointer. Identifies
 *	it's role (currently only property or method), and retrieves a 
 *	description.
 *
 * Result:
 *	Std Tcl result.
 *
 * Side effects:
 *	Throws HRESULT and char *.
 *-------------------------------------------------------------------------
 */
int TypeLib_DescribeTypeInfoElement (Tcl_Interp *pInterp, ITypeInfo *pti, 
									 const char *elem)
{
	ASSERT (pInterp != NULL && pti != NULL && elem != NULL);

	USES_CONVERSION;

	int					cmdresult = TCL_ERROR;
	HRESULT				hr;
	OptclBindPtr		bp;
	ULONG				hash;
	LPOLESTR			name = A2OLE(elem);
	CComPtr<ITypeComp>	pcmp;
	TObjPtr				presult,
						pdesc;
	BSTR				bdoc;

	try {
		hr = pti->GetTypeComp (&pcmp);
		CHECKHR(hr);

		hash = LHashValOfName (LOCALE_SYSTEM_DEFAULT, name);
		hr = pcmp->Bind (name, hash, INVOKE_FUNC | INVOKE_PROPERTYGET,  &bp.m_pti, &bp.m_dk, &bp.m_bp);
		if (FAILED(hr) || (bp.m_dk == DESCKIND_NONE)) {
			bp.ReleaseBindPtr();
			hr = pcmp->Bind (name, hash, INVOKE_PROPERTYPUT | INVOKE_PROPERTYPUTREF, &bp.m_pti, &bp.m_dk, &bp.m_bp);
		}

		CHECKHR(hr);
		

		cmdresult = TCL_OK;
		switch (bp.m_dk) {
		case DESCKIND_FUNCDESC:
			// check access restrictions on the function
			if (bp.m_bp.lpfuncdesc->wFuncFlags & FUNCFLAG_FRESTRICTED) {
				Tcl_SetResult(pInterp, "you aren't allowed to view: '", TCL_STATIC);
				Tcl_AppendResult (pInterp, (char*)elem, "'", NULL);
				cmdresult = TCL_ERROR;
			} else {
				presult.create();
				if (bp.m_bp.lpfuncdesc->invkind == INVOKE_FUNC) {// its a standard function 
					FUNCDESC2Obj (bp.m_pti, bp.m_bp.lpfuncdesc, pdesc);
					presult.lappend ("method").lappend(pdesc);
				}
				else {// its an implicit variable with accessor function
					DescPropertyFuncDesc(name, hash, bp.m_pti, pcmp, bp.m_bp.lpfuncdesc, pdesc);
					presult.lappend ("property").lappend(pdesc);
				}
			}
			break;

		case DESCKIND_VARDESC:
			if ((bp.m_bp.lpvardesc->wVarFlags & VARFLAG_FRESTRICTED)) {
				Tcl_SetResult (pInterp, "you aren't allowed to view: '", TCL_STATIC);
				Tcl_AppendResult (pInterp, (char*)elem, "'", NULL);
				cmdresult = TCL_ERROR;
			} else {
				VARDESC2Obj (bp.m_pti, bp.m_bp.lpvardesc, pdesc);
				presult.create();
				presult.lappend ("property").lappend(pdesc);
			}
			break;

		case DESCKIND_TYPECOMP: // don't know how to handle these ones at the moment
			Tcl_SetResult (pInterp, "typecomp", TCL_STATIC);
			break;

		case DESCKIND_IMPLICITAPPOBJ: // don't know how to handle these ones at the moment
			Tcl_SetResult (pInterp, "appobj", TCL_STATIC);
			break;

		case DESCKIND_NONE:
		default:
			Tcl_SetResult (pInterp, "can't find a description for '", TCL_STATIC);
			Tcl_AppendResult (pInterp, (char*)elem, "'", NULL);
			cmdresult = TCL_ERROR;
		}
	}
	catch (HRESULT hr) {
		Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
		cmdresult = TCL_ERROR;
	}
	catch (char *err) {
		Tcl_SetResult (pInterp, err, TCL_VOLATILE);
		cmdresult = TCL_ERROR;
	}

	// get the documentation string 
	
	if (cmdresult == TCL_OK && (bp.m_dk == DESCKIND_FUNCDESC || bp.m_dk == DESCKIND_VARDESC)) {
		if (SUCCEEDED(
			bp.m_pti->GetDocumentation (bp.m_dk==DESCKIND_FUNCDESC?bp.m_bp.lpfuncdesc->memid:bp.m_bp.lpvardesc->memid,
										NULL, &bdoc, NULL, NULL)) && bdoc != NULL)
		{
			presult.lappend(OLE2A(bdoc));
			SysFreeString (bdoc);
		}
		else
		{
			presult.lappend ("");
		}
		Tcl_SetObjResult (pInterp, presult);
	}	

	return cmdresult;
}





/*
 *-------------------------------------------------------------------------
 * TypeLib_LoadedLibs --
 *	Lists the currently loaded libraries
 * Result:
 *	TCL_OK iff ok.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
TCL_CMDEF(TypeLib_LoadedLibs)
{
	if (objc != 1) {
		Tcl_WrongNumArgs (pInterp, 1, objv, "");
		return TCL_ERROR;
	}
	TObjPtr presult;
	presult.create(false);
	try {
		for (TypeLibsTbl::iterator i = g_libs.begin();i != g_libs.end(); i++)
			presult.lappend(i.key(), pInterp);
	}
	catch (char *error) {
		Tcl_SetResult (pInterp, error, TCL_VOLATILE);
		return TCL_ERROR;
	}
	Tcl_SetObjResult (pInterp, presult);
	return TCL_OK;
}



/*
 *-------------------------------------------------------------------------
 * TypeLib_LoadLib --
 *	Ensures that a given library is loaded. A library is described in terms
 *	of its filename.
 *
 * Result:
 *	TCL_OK iff successful.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
TCL_CMDEF(TypeLib_LoadLib)
{
	if (objc != 2) {
		Tcl_WrongNumArgs (pInterp, 1, objv, "library_path");
		return TCL_ERROR;
	}
	TObjPtr libname;
	libname.attach(objv[1], false);
	TypeLib * pLib = g_libs.LoadLib (pInterp, libname);
	if (pLib) {
		Tcl_SetResult (pInterp, pLib->m_progname, TCL_VOLATILE);
		return TCL_OK;
	} else
		return TCL_ERROR;
}



/*
 *-------------------------------------------------------------------------
 * TypeLib_UnloadLib --
 *	Unloads a loaded library, specified in its human readable description.
 *	Perhaps this could be extended to take multiple arguments.
 *
 * Result:
 *	Always TCL_OK
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
TCL_CMDEF(TypeLib_UnloadLib)
{
	if (objc != 2) {
		Tcl_WrongNumArgs (pInterp, 1, objv, "full_libname");
		return TCL_ERROR;
	}
	TObjPtr libname;
	libname.attach(objv[1], false);
	g_libs.UnloadLib (pInterp, libname);
	return TCL_OK;
}

/*
 *-------------------------------------------------------------------------
 * TypeLib_IsLibLoaded --
 *	Returns true in the interpreter if the library (specifed in the first
 *	parameter) is correct.
 * Result:
 *	TCL_OK iff # of params ok
 * Side effects:
 *	None
 *-------------------------------------------------------------------------
 */
TCL_CMDEF(TypeLib_IsLibLoaded)
{
	USES_CONVERSION;
	if (objc != 4) {
		Tcl_WrongNumArgs (pInterp, 1, objv, "lib_guid majorver minorver");
		return TCL_ERROR;
	}
	GUID guid;
	long maj, min;

	char * szguid = Tcl_GetStringFromObj (objv[1], NULL);
	ASSERT (szguid != NULL);
	if (FAILED(CLSIDFromString(A2OLE(szguid), &guid))) {
		Tcl_SetResult (pInterp, "string isn't a guid", TCL_STATIC);
		return TCL_ERROR;
	}
	
	if (Tcl_GetLongFromObj(pInterp, objv[2], &maj) == TCL_ERROR)
		return TCL_ERROR;

	if (Tcl_GetLongFromObj(pInterp, objv[3], &min) == TCL_ERROR)
		return TCL_ERROR;

	TypeLib * pLib = NULL;

	pLib = g_libs.TypeLibFromUID(guid, maj, min);
	Tcl_ResetResult (pInterp);
	if (pLib) 
		Tcl_SetObjResult(pInterp, pLib->m_progname);
	
	return TCL_OK;
}

/*
 *-------------------------------------------------------------------------
 * TypeLib_TypesInLib --
 *	Returns a list in the interpreter holding the name and typekind of each
 *	type described in the library referenced by the first parameter to this
 *	command.
 * Result:
 *	Std Tcl return results.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
TCL_CMDEF (TypeLib_TypesInLib)
{
	USES_CONVERSION;

	if (objc != 2) {
		Tcl_WrongNumArgs (pInterp, 1, objv, "programmatic_libname");
		return TCL_ERROR;
	}


	TypeLib *ptl = NULL;
	TObjPtr name, typedesc, types, tname;
	HRESULT hr;
	TYPEKIND tkind;
	CComPtr<ITypeInfo> pti;
	int retresult = TCL_ERROR;
	TYPEATTR *pta = NULL;
	ULONG flags;



	types.create();
	name.attach(objv[1]);
	if (!g_libs.find(name, &ptl)) {
		Tcl_SetResult (pInterp, "can't find library name: ", TCL_STATIC);
		Tcl_AppendResult (pInterp, (char*)name, NULL);
		return TCL_ERROR;
	}

	ASSERT (ptl != NULL && ptl->m_ptl != NULL);
	UINT count = ptl->m_ptl->GetTypeInfoCount();
	try {
		for (UINT index = 0; index < count; index++) {
			
			
			// get the type of the typeinfo
			hr = ptl->m_ptl->GetTypeInfoType (index, &tkind);
			CHECKHR(hr);

			// get the next typeifo
			pti = NULL; // free the last typeinfo
			hr = ptl->m_ptl->GetTypeInfo (index, &pti);
			CHECKHR(hr);

			ASSERT (pti != NULL);

			// check whether this is a restricted type
			hr = pti->GetTypeAttr (&pta);
			CHECKHR(hr);
			flags = pta->wTypeFlags;
			ReleaseTypeAttr(pti, pta);

			if (flags & TYPEFLAG_FRESTRICTED)
				continue; // it is so skip the rest of this iteration
			TypeLib_GetName(ptl->m_ptl, pti, tname);

			typedesc.create();
			typedesc.lappend (TYPEKIND2Str(tkind)).lappend(tname);
			types.lappend (typedesc);
		}
		Tcl_SetObjResult (pInterp, types);
		retresult = TCL_OK;
	}

	catch (char *error) {
		Tcl_SetResult (pInterp, error, TCL_VOLATILE);
	}
	catch (HRESULT hr) {
		Tcl_SetResult (pInterp, (char*)HRESULT2Str(hr), TCL_DYNAMIC);
	}

	return retresult;
}




HRESULT TypeLib_GetDefaultInterface (ITypeInfo *pti, bool bEventSource, ITypeInfo ** ppdefti) {
	ASSERT (pti != NULL && ppdefti != NULL);

	OptclTypeAttr attr;
	attr = pti;
	ASSERT (attr.m_pattr != NULL);
	if (attr->typekind != TKIND_COCLASS)
		return E_FAIL;
	HRESULT hr;
	WORD selected = -1;

	for (WORD index = 0; index < attr->cImplTypes; index++) {

		INT implflags;
		hr = pti->GetImplTypeFlags(index, &implflags);
		if (FAILED(hr)) return hr;
		
		if ( ((implflags & IMPLTYPEFLAG_FDEFAULT) == IMPLTYPEFLAG_FDEFAULT) &&
			 ((bEventSource && (implflags & IMPLTYPEFLAG_FSOURCE) == IMPLTYPEFLAG_FSOURCE) ||
			 (!bEventSource && (implflags & IMPLTYPEFLAG_FSOURCE) != (IMPLTYPEFLAG_FSOURCE)))
		   ) {
			break;
		}
	}
	if (index == attr->cImplTypes)
		return E_FAIL;

	CComPtr<ITypeInfo> pimpl;
	HREFTYPE hreftype;

	// retrieve the referenced typeinfo
	hr = pti->GetRefTypeOfImplType(index, &hreftype);
	if (FAILED(hr)) return hr;

	hr = pti->GetRefTypeInfo(hreftype, &pimpl);
	if (FAILED(hr)) return hr;
	OptclTypeAttr pimplattr;
	pimplattr = pimpl;

	// resolve typedefs 
	while (pimplattr->typekind == TKIND_ALIAS) {
		CComPtr<ITypeInfo> pref;
		hr = pimpl->GetRefTypeInfo(pimplattr->tdescAlias.hreftype, &pref);
		if (FAILED(hr)) return hr;
		pimpl = pref;
		pimplattr = pimpl;
	}

	// if this isn't an interface forget it
	if ((pimplattr->typekind != TKIND_DISPATCH) &&
		(pimplattr->typekind != TKIND_INTERFACE))
		return E_FAIL;

	// okay - return the typeinfo to the caller
	return pimpl.CopyTo(ppdefti);
}




/*
 *-------------------------------------------------------------------------
 * TypeLib_TypeInfo --
 *	Implements the typelib::typeinfo command.
 *	
 * Result:
 *	Std Tcl result.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
TCL_CMDEF(TypeLib_TypeInfo)
{
	USES_CONVERSION;
	if (objc != 2 && objc != 3) {
		Tcl_WrongNumArgs (pInterp, 1, objv, "typename ?type_element_name?");
		return TCL_ERROR;
	}

	TObjPtr tname;
	TypeLib *ptl = NULL;
	BSTR bsTypename = NULL;

	CComPtr <ITypeInfo> pti;
	CComPtr <ITypeComp> ptc;
	int cmdresult = TCL_ERROR;

	try {
		tname.attach(objv[1]);
		TypeLib_ResolveName (tname, &ptl, &pti);


		if (objc == 2) { // describing the entire type
			cmdresult = TypeLib_DescribeTypeInfo (pInterp, pti);
		} else { // describing a single element within the type
			TObjPtr item;
			item.attach (objv[2]);
			cmdresult = TypeLib_DescribeTypeInfoElement (pInterp, pti, item);
		}
	}

	catch (char *error) {
		if (error != NULL)
			Tcl_SetResult (pInterp, error, TCL_VOLATILE);
	}

	catch (HRESULT hr) {
		Tcl_SetResult (pInterp, (char*)HRESULT2Str(hr), TCL_DYNAMIC);
	}

	return cmdresult;
}







TCL_CMDEF(TypeLib_GetRegLibPath)
{
	USES_CONVERSION;
	if (objc != 4) {
		Tcl_WrongNumArgs (pInterp, 1, objv, "lib_id majver minver");
		return TCL_ERROR;
	}

	char * szGuid = Tcl_GetStringFromObj (objv[1], NULL);
	long maj, min;

	if (Tcl_GetLongFromObj(pInterp, objv[2], &maj) == TCL_ERROR)
		return TCL_ERROR;

	if (Tcl_GetLongFromObj(pInterp, objv[3], &min) == TCL_ERROR)
		return TCL_ERROR;

	GUID guid;
	if (FAILED(CLSIDFromString(A2OLE(szGuid), &guid))) {
		Tcl_SetResult (pInterp, "failed to convert to a guid: ", TCL_STATIC);
		Tcl_AppendResult (pInterp, szGuid, NULL);
		return TCL_ERROR;
	}

	CComBSTR path;
	HRESULT hr = QueryPathOfRegTypeLib(guid, maj, min, LOCALE_SYSTEM_DEFAULT, &path);
	if (FAILED(hr)) {
		Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
		return TCL_ERROR;
	}
	Tcl_SetResult (pInterp, W2A(path), TCL_VOLATILE);
	return TCL_OK;
}

TCL_CMDEF(TypeLib_GetLoadedLibPath)
{
	if (objc != 2) {
		Tcl_WrongNumArgs (pInterp, 1, objv, "progname");
		return TCL_ERROR;
	}
	char * szProgName = Tcl_GetStringFromObj(objv[1], NULL);
	ASSERT (szProgName);

	TypeLib * plib = NULL;
	g_libs.find(szProgName, &plib);
	if (plib==NULL) {
		Tcl_SetResult (pInterp, "couldn't find loaded library: ", TCL_STATIC);
		Tcl_AppendResult (pInterp, szProgName, NULL);
		return TCL_ERROR;
	}
	Tcl_SetObjResult (pInterp, plib->m_path);
	return TCL_OK;
}


TCL_CMDEF(TypeLib_GetDetails)
{
	USES_CONVERSION;
	if (objc != 2) {
		Tcl_WrongNumArgs (pInterp, 1, objv, "progname");
		return TCL_ERROR;
	}
	char * szProgName = Tcl_GetStringFromObj (objv[1], NULL);
	ASSERT(szProgName);
	TypeLib * plib = NULL;
	g_libs.find(szProgName, &plib);
	if (plib == NULL) {
		Tcl_SetResult (pInterp, "couldn't find loaded library: ", TCL_STATIC);
		Tcl_AppendResult (pInterp, szProgName, NULL);
		return TCL_ERROR;
	}
	TObjPtr obj;
	obj.create();
	LPOLESTR pstr;
	HRESULT hr;
	hr = StringFromCLSID(plib->m_libattr->guid, &pstr);
	if (FAILED(hr)) {
		Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
		return TCL_ERROR;
	}
	obj.lappend(OLE2A(pstr));
	CoTaskMemFree(pstr);
	obj.lappend(plib->m_libattr->wMajorVerNum);
	obj.lappend(plib->m_libattr->wMinorVerNum);
	obj.lappend(plib->m_path);
	obj.lappend(plib->m_fullname);
	Tcl_SetObjResult (pInterp, obj);
	return TCL_OK;
}


/*
 *-------------------------------------------------------------------------
 * TypeLib_ResolveName --
 *	Resolves a library name and type name to a typeinfo.
 *	if pptl is not NULL then the TypeLib structure for this type is provided
 *	as the result, also.
 *	May throw an HRESULT or char*.
 *
 * Result:
 *	None.
 * Side effects:
 *
 *-------------------------------------------------------------------------
 */

void TypeLib_ResolveName (const char * lib, const char * type, 
						  TypeLib **pptl, ITypeInfo **ppinfo)
{
	ASSERT (lib != NULL  && type != NULL && ppinfo != NULL);
	HRESULT hr;

	TypeLib * ptl = NULL;

	// bind to the library
	if (g_libs.find (lib, &ptl) == NULL) 
		throw ("failed to bind to library");

	ASSERT (ptl != NULL && ptl->m_ptl != NULL);
	if (ptl->m_ptc == NULL)
		throw("library doesn't provide a compiling interface");
	if (pptl != NULL)
		*pptl = ptl;

	// find the type info if required
	if (ppinfo != NULL && type != NULL) {
		hr = BindTypeInfo(ptl->m_ptc, type, ppinfo);
		CHECKHR(hr);
		if (*ppinfo  == NULL) 
			throw ("failed to bind to type");
	}
}


/*
 *-------------------------------------------------------------------------
 * TypeLib_ResolveName --
 *	Resolves a fully formed type name (ie lib.type) to its type info.
 *	if pptl is not NULL then the TypeLib structure for this type is provided
 *	as the result, also.
 *	Throws HRESULT or (char*) in case of error. Apologies for this error style -
 *	I know that its predominant in this file - I was simply experimenting :)
 *
 * Result:
 *	None.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void TypeLib_ResolveName (const char *name, TypeLib **pptl, ITypeInfo **ppinfo)
{
	ASSERT (name != NULL);
	char * lib = NULL,
		 * type = NULL;
	char * copy = new char [strlen (name) + 1];
	strcpy (copy, name);

	try {
		lib =  strtok (copy, ".");
		type = strtok (NULL, ".");
		if (type == NULL && ppinfo != NULL)
			throw ("string is not properly formatted");
		TypeLib_ResolveName (lib, type, pptl, ppinfo);
		delete_ptr (copy);
	}

	catch (HRESULT hr) {
		delete_ptr (copy);
		throw (hr);
	}

	catch (char * error) {
		delete_ptr (copy);
		throw (error);
	}
}




/*
 *-------------------------------------------------------------------------
 * TypeLib_ResolveConstant --
 *
 * Result:
 *
 * Side effects:
 *
 *-------------------------------------------------------------------------
 */
bool TypeLib_ResolveConstant (Tcl_Interp *pInterp, ITypeInfo *pti, 
							  const char *member, TObjPtr &pObj)
{
	ASSERT (pInterp != NULL && pti != NULL && member != NULL);

	USES_CONVERSION;
	CComPtr<ITypeComp> ptc;
	CComPtr<ITypeInfo> ptmpti;
	DESCKIND dk;
	BINDPTR bp; bp.lpvardesc = NULL;
	HRESULT hr;
	LPOLESTR cnst = A2OLE (member);

#ifdef _DEBUG
	// *** TypeInfo must be an enumeration
	TYPEATTR * pattr;
	hr = pti->GetTypeAttr (&pattr);
	CHECKHR(hr);
	ASSERT (pattr->typekind == TKIND_ENUM);
	pti->ReleaseTypeAttr (pattr);
#endif


	try {
		hr = pti->GetTypeComp (&ptc);
		CHECKHR(hr);
		
		hr = ptc->Bind (cnst, LHashValOfName (LOCALE_SYSTEM_DEFAULT, cnst), 
					DISPATCH_PROPERTYGET, &ptmpti, &dk, &bp);
		CHECKHR(hr);
		if (dk == DESCKIND_NONE)
			throw ("can't find constant");
		ASSERT (dk == DESCKIND_VARDESC || dk == DESCKIND_IMPLICITAPPOBJ);
		
		if (bp.lpvardesc->varkind != VAR_CONST)
			throw ("member is not a constant");
		ASSERT (bp.lpvardesc->lpvarValue != NULL);
		if (bp.lpvardesc->lpvarValue == NULL)
			throw ("constant didn't have a associated value!");
		var2obj (pInterp, *(bp.lpvardesc->lpvarValue), NULL, pObj);
		pti->ReleaseVarDesc (bp.lpvardesc);
		return true;
	}

	catch (char *e) {
		if (bp.lpvardesc != NULL)
			pti->ReleaseVarDesc (bp.lpvardesc);

		Tcl_SetResult (pInterp, (char*)member, TCL_VOLATILE);
		Tcl_AppendResult (pInterp, ": ", e, NULL);
	}

	catch (HRESULT hr) {
		if (bp.lpvardesc != NULL)
			pti->ReleaseVarDesc (bp.lpvardesc);
		Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
	}
	return false;
}





/*
 *-------------------------------------------------------------------------
 * TypeLib_ResolveConstant --
 *	Attempts to resolve the name of a constant to its value, to be stored
 *	in pObj. An optional type info constrains the binding.
 * Result:
 *	true iff successful. Else error in interpreter.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
bool TypeLib_ResolveConstant (Tcl_Interp *pInterp, char *szname, 
							  TObjPtr &pObj, ITypeInfo *pTypeInfo /* = NULL */)
{
	ASSERT (pInterp != NULL && szname != NULL);

	const char *token = ".";
	char *name;
	char *szfirst;
	TypeLib	* ptl;
	CComPtr<ITypeInfo> pti;

	try {
		if (pTypeInfo == NULL)
		{
			// we'll use the stack for our allocation - saves on clean-up code
			name = (char*)alloca (strlen(szname) + 1);
			if (name == NULL) throw (HRESULT(E_OUTOFMEMORY));

			strcpy (name, szname);
			SplitTypedString (name, &szfirst);
			if (szfirst == NULL)
				throw ("badly formed constant");

			// at this point, name points to the name of the type, and 
			// szfirst points to the name of the constant
	
			// retrieve the typelibrary, info and compiler interfaces
			TypeLib_ResolveName (name, &ptl, &pti);
			return TypeLib_ResolveConstant (pInterp, pti, szfirst, pObj);
		} else {
			return TypeLib_ResolveConstant (pInterp, pTypeInfo, szname, pObj);
		}
	
		
	}

	catch (char *e) {
		Tcl_SetResult (pInterp, szname, TCL_VOLATILE);
		Tcl_AppendResult (pInterp, ": ", e, NULL);
	}

	catch (HRESULT hr) {
		Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
	}

	return false;
}



///// TEST CODE /////


TCL_CMDEF(TypeLib_ResolveConstantTest)
{
	if (objc != 2 && objc != 3)
	{
		Tcl_WrongNumArgs (pInterp, 1, objv, "lib.type.member   or    lib.type member");
		return TCL_ERROR;
	}

	TObjPtr result;
	TObjPtr p1;
	TObjPtr p2;
	bool bOk;

	if (objc == 2) {
		p1.attach(objv[1]);
		bOk = TypeLib_ResolveConstant(pInterp, p1, result, NULL);
	} else {
		CComPtr<ITypeInfo> pti;
		TypeLib * ptl;

		p1.attach(objv[1]);
		p2.attach(objv[2]);
		
		TypeLib_ResolveName (p1, &ptl, &pti);
		bOk = TypeLib_ResolveConstant (pInterp, p2, result, pti);
	}
	if (bOk) {
		Tcl_SetObjResult (pInterp, result);
		return TCL_OK;
	} else
		return TCL_ERROR;
}


