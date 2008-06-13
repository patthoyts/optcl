/*
 *------------------------------------------------------------------------------
 *	optcl.cpp
 *	Tcl gateway functions are placed here. De/Initialisation 
 *	of the object map occurs here, together with registration of many of
 *	optcl's commands.
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
#include "resource.h"
#include "optclobj.h"
#include "objmap.h"
#include "dispparams.h"
#include "typelib.h"

//----------------------------------------------------------------
HINSTANCE			ghDll = NULL;
CComModule			_Module;
CComPtr<IMalloc>	g_pmalloc;
bool				g_bTkInit = false;
//----------------------------------------------------------------

// Function declarations
void Optcl_Exit (ClientData);


TCL_CMDEF(OptclNewCmd);
TCL_CMDEF(OptclLockCmd);
TCL_CMDEF(OptclUnlockCmd);
TCL_CMDEF(OptclClassCmd);
TCL_CMDEF(OptclInterfaceCmd);
TCL_CMDEF(OptclBindCmd);
TCL_CMDEF(OptclIsObjectCmd);
TCL_CMDEF(OptclInvokeLibFunction);


//----------------------------------------------------------------

/*
 *-------------------------------------------------------------------------
 * DllMain --
 *	Windows entry point - ensures that ATL's ax containement are 
 *	initialised.
 *
 * Result:
 *	TRUE.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
BOOL WINAPI DllMain (HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
#ifdef _DEBUG
	int tmpFlag;
#endif // _DEBUG

	switch (fdwReason) 
	{
	case DLL_PROCESS_ATTACH:
		ghDll = hinstDLL;
		_Module.Init (NULL, (HINSTANCE)hinstDLL);
		AtlAxWinInit();

		#ifdef _DEBUG
		// memory leak detection - only in the debug build
			tmpFlag = _CrtSetDbgFlag( _CRTDBG_REPORT_FLAG );
			tmpFlag |= _CRTDBG_LEAK_CHECK_DF;
			_CrtSetDbgFlag( tmpFlag );
		#endif // _DEBUG
		break;
	case DLL_PROCESS_DETACH:
		_Module.Term();
		break;
	}

	return TRUE;
}


/*
 *-------------------------------------------------------------------------
 * Optcl_Exit --
 *	Called by Tcl pending exit. Removes all optcl objects. Uninits OLE.
 * Result:
 *	None.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void Optcl_Exit (ClientData)
{
	// remove all the elements of the table
	g_objmap.DeleteAll ();
	g_pmalloc.Release();
	OleUninitialize();
}



/*
 *-------------------------------------------------------------------------
 * OptclNewCmd --
 *	Implements the optcl::new command. Format of this command currently is:
 *		?-start? ?-window path? ProgIdOrClsidOrDocument
 *	For the time being, documents require the -window option to be used 
 *	as this code relies on ATL containement to locate the document server.
 *	This constraint is not ensured by this code.
 *	This can easily be implemented for documents that are not to be contained.
 *
 * Result:
 *	Standard Tcl result.
 * Side effects:
 *	Depends on parameters.
 *-------------------------------------------------------------------------
 */
TCL_CMDEF(OptclNewCmd)
{
	OptclObj *pObj = NULL;
	TObjPtr id;
	Tcl_Obj ** old = (Tcl_Obj**)objv;

	char *path = NULL;
	bool start = false;
	static const char * err = "?-start? ?-window path? ProgIdOrClsidOrDocument";
	static const char * errcreate = "error in creating object";

	if (objc < 2 || objc > 5) {
		Tcl_WrongNumArgs (pInterp, 1, objv, (char*)err);
		return TCL_ERROR;
	}

	// do we have flags?
	// process each one
	while (objc >= 3)
	{
		TObjPtr element;
		element.attach(objv[1]);
		int len = strlen(element);
		if (strncmp (element, "-start", len) == 0) {
			start = true;
		}
		else if (strncmp (element, "-window", len) == 0) {
			if (--objc <= 0) {
				Tcl_SetResult (pInterp, "expected path after -window", TCL_STATIC);
				return TCL_ERROR;
			}
			objv++;
			element.attach(objv[1]);
				path = (char*)element;
		}
		else {
			Tcl_SetResult (pInterp, "unknown flag: ", TCL_STATIC);
			Tcl_AppendResult (pInterp, (char*)element, NULL);
			return TCL_ERROR;
		}
		if (--objc <= 0) {
			Tcl_WrongNumArgs (pInterp, 1, old, (char*) err);
			return TCL_ERROR;
		}
		objv++;
	}

	id.attach (objv[1]);

	try {
		// try creating the object
		pObj = g_objmap.Create (pInterp, id, path, start);
		if (pObj == NULL) 
			Tcl_SetResult (pInterp, (char*)errcreate, TCL_STATIC);
	}

	catch (HRESULT hr) {
		Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
	}

	catch (char *err) {
		Tcl_SetResult (pInterp, err, TCL_VOLATILE);
	}

	catch (...)
	{
		Tcl_SetResult (pInterp, (char*)errcreate, TCL_STATIC);
	}

	if (pObj != NULL) {
		Tcl_SetResult (pInterp, (char*)(const char*)(*pObj), TCL_VOLATILE);
		return TCL_OK;
	}
	else
		return TCL_ERROR;
}






/*
 *-------------------------------------------------------------------------
 * OptclLockCmd --
 *	Implements the optcl::lock command.
 * Result:
 *	Standard Tcl result
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
TCL_CMDEF(OptclLockCmd)
{
	if (objc != 2) {
		Tcl_WrongNumArgs (pInterp, 1, objv, "object");
		return TCL_ERROR;
	}
	TObjPtr name;
	name.attach(objv[1]);
	if (!g_objmap.Lock(name)) {
		return ObjectNotFound(pInterp, name);
	}
	return TCL_OK;
}





/*
 *-------------------------------------------------------------------------
 * OptclUnlockCmd --
 *	Implements the optcl::unlock command.
 *
 * Result:
 *	Standard Tcl result.
 *
 * Side effects:
 *	If the reference count of the object hits zero, then the object will be
 *	deleted, together with its Tcl command and its container window, if it 
 *	exists.
 *-------------------------------------------------------------------------
 */
TCL_CMDEF(OptclUnlockCmd)
{
	if (objc < 2) {
		Tcl_WrongNumArgs (pInterp, 1, objv, "object ...");
		return TCL_ERROR;
	}

	TObjPtr name;
	for (int i = 1; i < objc; i++) {
		name.attach(objv[1]);
		g_objmap.Unlock(name);
	}
	return TCL_OK;
}



/*
 *-------------------------------------------------------------------------
 * OptclInvokeLibFunction --
 *	Wild and useless attempt at calling ITypeInfo declared static DLL
 *	functions. Sigh!
 * Result:
 *
 * Side effects:
 *
 *-------------------------------------------------------------------------
 */
TCL_CMDEF(OptclInvokeLibFunction)
{
	USES_CONVERSION;
	DispParams				dp;
	LPOLESTR				olename;
	TObjPtr					name,
							presult;
	CComPtr<ITypeInfo>		pinfo;
	CComPtr<ITypeInfo>		pti;
	CComPtr<ITypeComp>		pcmp;
	HRESULT					hr;
	DESCKIND				dk; dk = DESCKIND_NONE;
	BINDPTR					bp; bp.lpfuncdesc = NULL;
	DISPID					dispid;
	EXCEPINFO				ei;
	UINT					ea = 0;
	VARIANT					varResult;
	bool					bOk = false;
	VOID				*	pFunc = NULL;

	if (objc < 3) {
		Tcl_WrongNumArgs (pInterp, 1, objv, "typename function args...");
		return TCL_ERROR;
	}

	try {
		// attempt to resolve the type
		name.attach(objv[1]);
		TypeLib_ResolveName (name, NULL, &pinfo);

		hr = pinfo->GetTypeComp (&pcmp);
		CHECKHR_TCL(hr, pInterp, TCL_ERROR);

		name.attach(objv[2]);
		olename = A2OLE(name);
		hr = pcmp->Bind (olename, LHashValOfName(LOCALE_SYSTEM_DEFAULT, olename), 
			INVOKE_FUNC, &pti, &dk, &bp);

		CHECKHR(hr);
		if (dk != DESCKIND_FUNCDESC || bp.lpfuncdesc->funckind != FUNC_STATIC) {
			Tcl_SetResult (pInterp, "static method not found: ", TCL_STATIC);
			Tcl_AppendResult (pInterp, (char*)name, NULL);
		} else {
			ASSERT (bp.lpfuncdesc != NULL);
			dispid = bp.lpfuncdesc->memid;
			hr = pinfo->AddressOfMember (dispid, INVOKE_FUNC, &pFunc);
			CHECKHR_TCL(hr, pInterp, TCL_ERROR);
			int params = objc - 3;

			// check for the last parameter being the return value and take
			// this into account when checking parameter counts
			int reqparams = bp.lpfuncdesc->cParams;
			if (reqparams > 0 && bp.lpfuncdesc->lprgelemdescParam[reqparams - 1].paramdesc.wParamFlags & PARAMFLAG_FRETVAL)
				--reqparams;


			if (params <= reqparams &&
				params >= (reqparams -bp.lpfuncdesc->cParamsOpt))
			{
				VariantInit (&varResult);
				
				// set up the dispatch arguments - must be in reverse order
				dp.Args (params);
				for (int i = params-1; i >= 0; i--)
				{
					bool		con_ok;
					LPVARIANT	pv;

					name.attach(objv[i+3]);
					// are we dealing with referenced parameter?
					if (bp.lpfuncdesc->lprgelemdescParam[i].tdesc.vt == VT_PTR) {
						ASSERT (bp.lpfuncdesc->lprgelemdescParam[i].tdesc.lptdesc != NULL);

						// allocate a variant to store the *value*
						pv = new VARIANT;
						VariantInit (pv);
						con_ok = obj2var_ti(pInterp, name, *pv, pti, &(bp.lpfuncdesc->lprgelemdescParam[i].tdesc));
						// we'll now set it as a reference for the dispatch parameters array
						// on destruction, the array will take care of clearing the variant
						dp.Set(params - i - 1, pv);
					}
					
					else 
						con_ok = obj2var_ti(pInterp, name, dp[params - i - 1], pti, &(bp.lpfuncdesc->lprgelemdescParam[i].tdesc));

					
					if (!con_ok)
					{
						ReleaseBindPtr (pti, dk, bp);
						return false; // error in type conversion
					}
				}
				hr = pinfo->Invoke (pFunc, dispid, DISPATCH_METHOD, &dp, &varResult, &ei, &ea);
				//hr = pDisp->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, 
				//	&dp, &varResult, &ei, &ea);

				if (hr == DISP_E_EXCEPTION)
					Tcl_SetResult (pInterp, ExceptInfo2Str (&ei), TCL_DYNAMIC);
				else if (hr == DISP_E_TYPEMISMATCH) {
					TDString td("type mismatch in parameter #");
					td << (long)(ea);
					Tcl_SetResult (pInterp, td, TCL_VOLATILE);
				}
				else
					CHECKHR_TCL(hr, pInterp, TCL_ERROR);
				if (FAILED(hr))
					return TCL_ERROR;
				if (bOk = var2obj(pInterp, varResult, NULL, presult))
					Tcl_SetObjResult (pInterp, presult);
				VariantClear(&varResult);
			}
			else
			{
				Tcl_SetResult (pInterp, "wrong # args", TCL_STATIC);
			}
		}
	}

	catch (HRESULT hr) {
		Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
	}	
	return bOk?TCL_OK:TCL_ERROR;
}

/*
 *-------------------------------------------------------------------------
 * OptclClassCmd --
 *	Implements the optcl::class command.
 *
 * Result:
 *	Standard Tcl result
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
TCL_CMDEF(OptclClassCmd)
{
	if (objc != 2) {
		Tcl_WrongNumArgs (pInterp, 1, objv, "object");
		return TCL_ERROR;
	}
	OptclObj *pObj = NULL;
	TObjPtr name;
	TObjPtr classname;

	name.attach (objv[1]);
	pObj = g_objmap.Find (name);
	if (pObj == NULL)
		return ObjectNotFound (pInterp, name);
	try {
		pObj->CoClassName(classname);
	}

	catch (HRESULT hr) {
		Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
		return TCL_ERROR;
	}
	Tcl_SetObjResult (pInterp, classname);
	return TCL_OK;
}




/*
 *-------------------------------------------------------------------------
 * OptclInterfaceCmd --
 *	Implements the optcl::interface command. Will either retrieve the 
 *	current active interface or set it:
 *		optcl::interface objid ?newinterface?
 *
 *	The new interface must be a proper typename. i.e. lib.type
 * Result:
 *	Standard Tcl result.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
TCL_CMDEF(OptclInterfaceCmd)
{
	if (objc != 2 && objc != 3) {
		Tcl_WrongNumArgs (pInterp, 1, objv, "object ?interface?");
		return TCL_ERROR;
	}

	OptclObj *pObj = NULL;
	TObjPtr name;
	TObjPtr intfname;

	name.attach (objv[1]);
	pObj = g_objmap.Find (name);
	if (pObj == NULL)
		return ObjectNotFound (pInterp, name);
	try {
		if (objc == 2) // get the current interface name
			pObj->InterfaceName(intfname);
		else // we are setting the interface
		{
			intfname.attach(objv[2]);
			pObj->SetInterfaceName(intfname);
		}
	}

	catch (HRESULT hr) {
		Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
		return TCL_ERROR;
	}
	catch (char *error) {
		Tcl_SetResult (pInterp, error, TCL_VOLATILE);
		return TCL_ERROR;
	}

	Tcl_SetObjResult (pInterp, intfname);
	return TCL_OK;
}




/*
 *-------------------------------------------------------------------------
 * OptclBindCmd --
 *	Implements the optcl::bind command. This enables the setting, unsetting
 *	and retrieving of a binding to an objects event, either on its default 
 *	interface (in which case the interface type is not required) or on 
 *	a non-default event interface. e.g.
 *		optcl::bind $obj NewDoc	OnNewDocTclhandler
 *		optcl::bind $obj NewDoc ==> OnNewDocTclhandler
 *		optcl::bind $obj ICustomInterface.Foo FooHandler
 *
 *	The tcl command is then called when the specified event is fired.
 *	The parameter list of the event is prepended with the identifier
 *	object that fired event. If a parameter of an event is an object, 
 *	it's lifetime is only within the duration of the execution of the 
 *	tcl handler. To allow for the object to persist after the handler has
 *	completed, the tcl script must call optcl::lock on the object.
 *
 * Result:
 *	Standard Tcl result.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
TCL_CMDEF(OptclBindCmd)
{
	if (objc != 3 && objc != 4)
	{
		Tcl_WrongNumArgs (pInterp, 1, objv, "object event_name ?tcl_command?");
		return TCL_ERROR;
	}

	OptclObj *pObj = NULL;
	TObjPtr name;
	TObjPtr value;
	bool bOk = false;

	name.attach (objv[1]);
	pObj = g_objmap.Find (name);
	if (pObj == NULL)
		return ObjectNotFound (pInterp, name);

	name.attach(objv[2]); // the event name
	if (objc == 3) // get the current binding (if any) for an event
		bOk = pObj->GetBinding (pInterp, name);
	else // we are setting the interface
	{
		value.attach(objv[3]);
		if (bOk = pObj->SetBinding(pInterp, name, value))
			Tcl_SetObjResult (pInterp, value);
	}
	return (bOk?TCL_OK:TCL_ERROR);
}



/*
 *-------------------------------------------------------------------------
 * OptclIsObjectCmd --
 *	Returns a boolean in the interpreter - true iff the only parameter
 *	for this command is an object.
 *
 * Result:
 *	TCL_OK always for the correct number of parameters.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
TCL_CMDEF(OptclIsObjectCmd)
{
	if (objc != 2) {
		Tcl_WrongNumArgs (pInterp, 1, objv, "object");
		return TCL_ERROR;
	}
	TObjPtr name(objv[1], false);
	TObjPtr found(false, false);
	if (g_objmap.Find (name))
	{
		found = true;
	}
	Tcl_SetObjResult (pInterp, found);
	return TCL_OK;
}


/*
 *-------------------------------------------------------------------------
 * Optcl_Init --
 *	Tcl's first entry point. Initialises ole, sets up the exit handler,
 *	invokes the startup script (stored in a windows resource) and setsup
 *	the optcl namespace. Finally, it initialises the type library system.
 *
 * Result:
 *	Standard Tcl result.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
int Optcl_Init (Tcl_Interp *pInterp)
{
	Tcl_CmdInfo *pinfo = NULL;

#ifdef USE_TCL_STUBS
#if (TCL_MAJOR_VERSION == 8) && (TCL_MINOR_VERSION >= 1)
	// initialise the Tcl stubs - failure is very bad
	if (Tcl_InitStubs (pInterp, "8.0", 0) == NULL)
		return TCL_ERROR;

	// if Tk is loaded then initialise the Tk stubs
	if (Tcl_Eval (pInterp, "package present Tk") != TCL_ERROR) {
		// initialise the Tk stubs - failure 
		if (Tk_InitStubs (pInterp, "8.0", 0) == NULL)
			return TCL_ERROR;
		g_bTkInit = true;
	}
#else
#error Wrong Tcl version for Stubs
#endif // (TCL_MAJOR_VERSION == 8) && (TCL_MINOR_VERSION >= 1)
#endif // USE_TCL_STUBS

	HRESULT hr;
	Tcl_PkgProvide(pInterp, "optcl", "3.0");

	OleInitialize(NULL);
	hr = CoGetMalloc(1, &g_pmalloc);
	CHECKHR_TCL(hr, pInterp, TCL_ERROR);

	Tcl_CreateExitHandler (Optcl_Exit, NULL);
	/*
	HRSRC hrsrc = FindResource (ghDll, MAKEINTRESOURCE(IDR_TYPELIB), _T("TCL_SCRIPT"));
	if (hrsrc == NULL) {
		Tcl_SetResult (pInterp, "failed to locate internal script", TCL_STATIC);
		return TCL_ERROR;
	}
	HGLOBAL hscript = LoadResource (ghDll, hrsrc);
	if (hscript == NULL) {
		Tcl_SetResult (pInterp, "failed to load internal script", TCL_STATIC);
		return TCL_ERROR;
	}

	ASSERT (hscript != NULL);
	char *szscript = (char*)LockResource (hscript);

	ASSERT (szscript != NULL);
	if (Tcl_GlobalEval (pInterp, szscript) == TCL_ERROR)
		return TCL_ERROR;
	*/

	Tcl_CreateObjCommand (pInterp, "optcl::new", OptclNewCmd, NULL, NULL);
	Tcl_CreateObjCommand (pInterp, "optcl::lock", OptclLockCmd, NULL, NULL);
	Tcl_CreateObjCommand (pInterp, "optcl::unlock", OptclUnlockCmd, NULL, NULL);
	Tcl_CreateObjCommand (pInterp, "optcl::class", OptclClassCmd, NULL, NULL);
	Tcl_CreateObjCommand (pInterp, "optcl::interface", OptclInterfaceCmd, NULL, NULL);
	Tcl_CreateObjCommand (pInterp, "optcl::bind", OptclBindCmd, NULL, NULL);
	Tcl_CreateObjCommand (pInterp, "optcl::module", OptclInvokeLibFunction, NULL, NULL);
	Tcl_CreateObjCommand (pInterp, "optcl::isobject", OptclIsObjectCmd, NULL, NULL);


	/// TESTS ///
	Tcl_CreateObjCommand (pInterp, "optcl::vartest", Obj2VarTest, NULL, NULL);
	
	return TypeLib_Init(pInterp);
}

