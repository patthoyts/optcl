/*
 *------------------------------------------------------------------------------
 *	eventbinding.cpp
 *	Declares classes used for implementing OpTcl's event binding.
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
#include "EventBinding.h"
#include "optcl.h"
#include "utility.h"
#include "objmap.h"
#include "typelib.h"
#include "optclbindptr.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////


/*
 *-------------------------------------------------------------------------
 * EventNotFound --
 *	Writes a standard error message to the interpreter, indicating
 *	that an event was not found.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void EventNotFound (Tcl_Interp *pInterp, const char * event)
{
	Tcl_SetResult (pInterp, "event not found: ", TCL_STATIC);
	Tcl_AppendResult (pInterp, (char*) event, NULL);
}


/*
 *-------------------------------------------------------------------------
 * EventBindings::EventBindings --
 *	Constructor - caches the parameters, and attempts to bind to ITypeComp
 *	interface. If not found, throw an HRESULT.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
EventBindings::EventBindings(OptclObj *pObj, REFGUID guid, ITypeInfo *pInfo) 
: m_ref(0), m_bindings(0), m_cookie(0)
{
	ASSERT (pInfo!= NULL);
	ASSERT (pObj != NULL);

	HRESULT hr;

	m_pti = pInfo;
	m_optclobj = pObj;
	m_guid = guid;

	hr = m_pti->GetTypeComp(&m_ptc);
	CHECKHR(hr);
	ASSERT (m_ptc != NULL);
}


/*
 *-------------------------------------------------------------------------
 * EventBindings::~EventBindings --
 *	Destructor - ensures that any event bindings within this object
 *	are deleted.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
EventBindings::~EventBindings()
{
	DeleteTbl();
}



/*
 *-------------------------------------------------------------------------
 * EventBindings::DeleteTbl --
 *	Iterates through the command table, and deletes each object
 *	before finally deleting the hash table.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void EventBindings::DeleteTbl()
{
	DispCmdTbl::iterator i;
	TObjPtr p;

	for (i = m_cmdtbl.begin(); i != m_cmdtbl.end(); i++)
	{
		BindingProps *p = *i;
		ASSERT (p != NULL);
		delete p;
	}
	m_cmdtbl.deltbl();
	m_bindings = 0;
}




// IUnknown Entries
/*
 *-------------------------------------------------------------------------
 * EventBindings::QueryInterface --
 *	Implements the IUnknown member. Successfull iff riid is for IUnknown,
 *	IDispatch, or the event interface that we're implementing.
 *
 * Result:
 *	Standard COM result.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
HRESULT STDMETHODCALLTYPE EventBindings::QueryInterface(REFIID riid, void ** ppvObject)
{
	HRESULT hr = S_OK;
	if (riid == IID_IUnknown)
		*ppvObject = (IUnknown*)this;
	else if (riid == IID_IDispatch || riid == m_guid)
		*ppvObject = (IDispatch*)this;
	else
		hr = E_NOINTERFACE;
	if (SUCCEEDED(hr))
		AddRef();

	return hr;
}





/*
 *-------------------------------------------------------------------------
 * EventBindings::AddRef --
 *	Implements the IUnknown member.
 * Result:
 *	Standard AddRef result.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
ULONG STDMETHODCALLTYPE EventBindings::AddRef( void)
{
	return InterlockedIncrement (&m_ref);
}





/*
 *-------------------------------------------------------------------------
 * EventBindings::Release --
 *	Implements the IUnknown member. In fact, this never deletes this object
 *	That responsibility is always with the optcl object. I am not sure
 *	if this approach is a good one. So beware! :o
 *
 * Result:
 *	Standard Release result.
 * Side effects:
 *	None, what so ever.
 *-------------------------------------------------------------------------
 */
ULONG STDMETHODCALLTYPE EventBindings::Release( void)
{
	// a dummy function
	if (InterlockedDecrement (&m_ref) <= 0)  
		m_ref = 0;
	return m_ref;
}





// IDispatch Entries
/*
 *-------------------------------------------------------------------------
 * EventBindings::GetTypeInfoCount --
 *	Implements the IDispatch member. 1 if we did get a type info. The
 *	check isn't necessary at all, but I've put it there just as a reminder.
 * Result:
 *	S_OK always.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
HRESULT STDMETHODCALLTYPE EventBindings::GetTypeInfoCount(UINT *pctinfo)
{
	ASSERT (m_pti != NULL);
	*pctinfo = (m_pti!=NULL)?1:0;
	return S_OK;
}






/*
 *-------------------------------------------------------------------------
 * EventBindings::GetTypeInfo --
 *	Implements the IDispatch member. Standard stuff.
 * Result:
 *	Standard GetTypeInfo result.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
HRESULT STDMETHODCALLTYPE 
EventBindings::GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo ** ppTInfo)
{
	ASSERT (lcid == LOCALE_SYSTEM_DEFAULT);
	HRESULT hr = S_OK;
	if (iTInfo != 0 || m_pti == NULL || ppTInfo == NULL)
		hr = DISP_E_BADINDEX;
	else {
		(*ppTInfo) = m_pti;
		(*ppTInfo)->AddRef();
	}
	return hr;
}





/*
 *-------------------------------------------------------------------------
 * EventBindings::GetIDsOfNames --
 *	Standard IDispatch member. Passes on the responsiblity to the type
 *	library.
 * Result:
 *	Standard com result.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
HRESULT STDMETHODCALLTYPE 
EventBindings::GetIDsOfNames(REFIID riid, LPOLESTR  *rgszNames, 
	          UINT cNames, LCID lcid, DISPID *rgDispId)
{
	HRESULT hr = S_OK;
	if (m_pti == NULL) 
		hr = DISP_E_UNKNOWNNAME;
	if (lcid != LOCALE_SYSTEM_DEFAULT)
		hr = DISP_E_UNKNOWNLCID;
	else
		hr = DispGetIDsOfNames (m_pti, rgszNames, cNames, rgDispId);
	return hr;
}





/*
 *-------------------------------------------------------------------------
 * EventBindings::Invoke --
 *	Called by the event source when an event is raised. Attempts to
 *	find event in the event table. If not bound yet, simply returns, 
 *	otherwise, asks the binding to be evaluated.
 * Result:
 *	S_OK iff succeeded. If error and an exception info struct is available, 
 *	then it is filled out.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
HRESULT STDMETHODCALLTYPE 
EventBindings::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
	   WORD wFlags, DISPPARAMS *pDispParams, LPVARIANT pVarResult, 
	   LPEXCEPINFO pExcepInfo, UINT *puArgErr)
{
	if (pDispParams == NULL)
		return E_FAIL;

	// look up the dispatch id in our table.
	BindingProps *bp = NULL;

	

	if (m_cmdtbl.find((DISPID*)dispIdMember, &bp) != NULL)
	{
		ASSERT (bp != NULL);
		int res = bp->Eval(m_optclobj, pDispParams, pVarResult, pExcepInfo);
		if (res == TCL_ERROR) {
			if (pExcepInfo == NULL)
				return E_FAIL;
			else
				return DISP_E_EXCEPTION;
		}
 	}
	return S_OK;
}



/*
 *-------------------------------------------------------------------------
 * EventBindings::Name2Dispid --
 *	Binds a string name to a dispatch id in the implemented event interface.
 *	
 * Result:
 *	true iff successful. pInterp will contain the error string if not.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
bool EventBindings::Name2Dispid (Tcl_Interp *pInterp, const char * name, DISPID &dispid)
{
	ASSERT (pInterp != NULL && name != NULL && m_ptc != NULL);

	USES_CONVERSION;
	LPOLESTR olename;
	HRESULT hr;
	OptclBindPtr obp;
	bool bOk = false;

	olename = A2OLE (name);
	
	try {
		hr = m_ptc->Bind (olename, LHashValOfName(LOCALE_SYSTEM_DEFAULT, olename), 
			INVOKE_FUNC, &obp.m_pti, &obp.m_dk, &obp.m_bp);
		CHECKHR(hr);
		
		if (obp.m_dk == DESCKIND_FUNCDESC) {
			ASSERT (obp.m_bp.lpfuncdesc != NULL);
			dispid = obp.m_bp.lpfuncdesc->memid;
			bOk = true;
		} else
			EventNotFound(pInterp, name);
	}
	catch (HRESULT hr) {
		Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_DYNAMIC);
	}

	return bOk;
}

/*
 *-------------------------------------------------------------------------
 * EventBindings::TotalBindings --
 *	Returns the total number of even bindings in this collection.
 * Result:
 *	The count.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
ULONG EventBindings::TotalBindings ()
{
	return m_bindings;
}


/*
 *-------------------------------------------------------------------------
 * EventBindings::SetBinding --
 *	Attempts to bind an event, give by name, to a tcl command (a tcl object)
 *
 * Result:
 *	true iff successful. Else, error string in interpreter.
 * Side effects:
 *	Changes the count of the number of bindings.
 *-------------------------------------------------------------------------
 */
bool EventBindings::SetBinding (Tcl_Interp *pInterp, const char * name, Tcl_Obj *pCommand)
{
	ASSERT (pInterp != NULL && name != NULL && pCommand != NULL);
	
	DISPID	dispid;

	if (!Name2Dispid (pInterp, name, dispid))
		return false;

	BindingProps *pbp = NULL;
	if (!m_cmdtbl.find ((DISPID*)(dispid), &pbp)) {
		pbp = new BindingProps (pInterp, pCommand);
		m_cmdtbl.set ((DISPID*)(dispid), pbp);
		m_bindings++;
	} else {
		ASSERT (pbp != NULL);
		pbp->m_pInterp = pInterp;
		pbp->m_pScript = pCommand;
	}
	return true;	
}



/*
 *-------------------------------------------------------------------------
 * EventBindings::GetBinding --
 *	Returns within the interpreter the tcl command bound to an event.
 * Result:
 *	true iff successful. Else, error string in interpreter.
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
bool EventBindings::GetBinding (Tcl_Interp *pInterp, const char * name)
{
	ASSERT (pInterp != NULL && name != NULL);
	DISPID dispid;

	if (!Name2Dispid (pInterp, name, dispid))
		return false;


	BindingProps *pbp = NULL;
	if (m_cmdtbl.find ((DISPID*)(dispid), &pbp)) {
		ASSERT (pbp != NULL);
		Tcl_SetObjResult (pInterp, pbp->m_pScript);
	}
	return true;
}


/*
 *-------------------------------------------------------------------------
 * EventBindings::DeleteBinding --
 *	Removes an event binding.
 * Result:
 *	false iff 'name'  is not the name of an existing event.
 * Side effects:
 *	Changes the count of the number of bindings.
 *-------------------------------------------------------------------------
 */
bool EventBindings::DeleteBinding (Tcl_Interp *pInterp, const char * name)
{
	ASSERT (pInterp != NULL && name != NULL);
	DISPID dispid;

	if (!Name2Dispid (pInterp, name, dispid))
		return false;
	BindingProps *pbp = NULL;
	if (m_cmdtbl.find ((DISPID*)(dispid), &pbp)) {
		ASSERT (pbp != NULL);
		delete pbp;
		m_cmdtbl.delete_entry ((DISPID*)(dispid));
		m_bindings--;
	}
	return true;
}


/*
 *-------------------------------------------------------------------------
 * BindingProps::Eval --
 *	The guts of the event handler. Pulls out the parameters (in reverse order)
 *	and invokes that on the registered tcl handler. Note that the command 
 *	line will look like:
 *		handler objid ?arg ...?
 *	where objid is the optcl identifier of the activex object that created
 *	the event.
 *
 * Result:
 *	Standard Tcl Result.
 *
 * Side effects:
 *	Depends on the tcl handler.
 *-------------------------------------------------------------------------
 */
int BindingProps::Eval (OptclObj *pObj, DISPPARAMS *pDispParams, LPVARIANT pVarResult, 
						LPEXCEPINFO pExcepInfo)
{
	ASSERT (m_pInterp != NULL && m_pScript.isnotnull());
	ASSERT (pDispParams != NULL);
	ASSERT (pObj != NULL);

	OptclObj ** ppObjs = NULL;
	
	TObjPtr cmd;
	UINT count;
	int result = TCL_ERROR;
	
	cmd.copy(m_pScript);
	cmd.lappend ((const char*)(*pObj)); // cast to a string

	ASSERT (pDispParams->cNamedArgs == 0);

	// potentially all the parameters could result in an object, so
	// allocate an array and set it all to nulls.
	if (pDispParams->cArgs > 0) {
		ppObjs = (OptclObj**)calloc(pDispParams->cArgs, sizeof (OptclObj*));
		if (ppObjs == NULL) {
			Tcl_SetResult (m_pInterp, "failed to allocate memory", TCL_STATIC);
			Tcl_BackgroundError (m_pInterp);
			return TCL_ERROR;
		}
	}

	// temporarily increase the reference count on the object
	// this way, if the event handler unlocks the objects, a possible
	// destruction doesn't occur until this event has been handled
	g_objmap.Lock (pObj);

	try {
		// convert the dispatch parameters into Tcl arguments
		for (count = 0; count < pDispParams->cArgs; count++)
		{
			TObjPtr param;
			if (!var2obj(m_pInterp, pDispParams->rgvarg[pDispParams->cArgs - count - 1], param, ppObjs+count))
				break;
			cmd.lappend(param, m_pInterp);
		}
	}

	// the error will already be stored in the interpreter
	catch (char *) {
	}

	
	// if we managed to convert all the parameters, invoke the function
	if (count == pDispParams->cArgs)
		result = Tcl_GlobalEvalObj (m_pInterp, cmd);


	
	// deallocate the objects
	for (count = 0; count < pDispParams->cArgs; count++)
	{
		if (ppObjs[count] != NULL) 
			g_objmap.Unlock (ppObjs[count]);
	}

	// release the object pointers
	if (ppObjs != NULL)
		free (ppObjs);

	if (result == TCL_ERROR)
	{
		// do we have a exception storage
		if (pExcepInfo != NULL)
		{
			// fill it in
			_bstr_t src(Tcl_GetStringResult(m_pInterp));
			_bstr_t name("OpTcl");
			pExcepInfo->wCode = 1001;
			pExcepInfo->wReserved = 0;
			pExcepInfo->bstrSource = name.copy();
			pExcepInfo->bstrDescription = src.copy();
			pExcepInfo->bstrHelpFile = NULL;
			pExcepInfo->pvReserved = NULL;
			pExcepInfo->pfnDeferredFillIn = NULL;
		}
		Tcl_BackgroundError (m_pInterp);
	}
	else
	{
		// get the Tcl result and store it in the result variant
		// currently we are limited to the basic types, until
		// I get the time to pull the typelib stuff to this point
		if (pVarResult != NULL)
		{
			TObjPtr pres(Tcl_GetObjResult (m_pInterp), false);
			VariantInit(pVarResult);
			obj2var (pres, *pVarResult);
		}
	}

	// finally unlock the object
	g_objmap.Unlock (pObj);
	return result;
}










