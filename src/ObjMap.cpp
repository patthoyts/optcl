/*
 *------------------------------------------------------------------------------
 *	objmap.cpp
 *	Implementation of the object table.
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
#include "optcl.h"
#include "utility.h"
#include "objmap.h"




// globals

// the one and only object map for this extension
// this class uses a Tcl hash table - this usually wouldn't be
// safe, except that this hash table is initialised (courtsey of THash<>)
// only on first uses (lazy). So it should be okay. Not sure how 
// this will behave in a multithreaded application

ObjMap	g_objmap;


//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////


ObjMap::ObjMap() : m_destructpending(false)
{

}

ObjMap::~ObjMap()
{
	
}


/*
 *-------------------------------------------------------------------------
 * ObjMap::DeleteAll --
 *	Deletes all objects in the system.
 * Result:
 *	None.
 * Side effects:
 *	Deletes all object commands from respective interpreters
 *-------------------------------------------------------------------------
 */
void ObjMap::DeleteAll ()
{
#ifdef _DEBUG
	ObjDump();
#endif // _DEBUG

	ObjNameMap::iterator i;
	for (i = m_namemap.begin(); i != m_namemap.end(); i++) {
		OptclObj *pobj = *i;
		ASSERT (pobj != NULL);
		DeleteCommand (pobj);
		delete pobj;
	}
	m_namemap.deltbl();
	m_unkmap.deltbl();
}

/*
 *-------------------------------------------------------------------------
 * ObjMap::ObjDump
 *	Dumps the current contents of the object map to the Debug Stream
 *
 * Result:
 *	None.
 *
 * Side Effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void ObjMap::ObjDump () 
{
	TRACE("BEGIN: OpTcl Object Dump\n");

	ObjNameMap::iterator i;
	for (i = m_namemap.begin(); i != m_namemap.end(); i++) {
		OptclObj *pobj = *i;
		ASSERT (pobj != NULL);
		TObjPtr interfacename;
		pobj->InterfaceName(interfacename);
		TRACE("\t%s %s %d\n", (char*)interfacename, pobj->m_name.c_str(), pobj->m_refcount);
	}	
	TRACE("END:   OpTcl Object Dump\n");
}

/*
 *-------------------------------------------------------------------------
 * ObjMap::Create --
 *	Creates an object for a particular interpreter. The type of the object 
 *	is identified by a string representing either a CLSID or ProgId. It would
 *	be neat if it also could be the name of a file on the local system or 
 *	at some URL. More on this later....
 *	If the object already exists in the object table, then we return that 
 *	object. In fact, this is a limitation of the system, as objects that have
 *	been registered in one interpreter cannot be accessed from another 
 *	interpreter.
 *
 * Result:
 *	A non-null pointer to the underlying Optcl object representation, 
 *	iff successful.
 *
 * Side effects:
 *	Creates also the Tcl command used to invoke this object.
 *
 *-------------------------------------------------------------------------
 */
OptclObj * ObjMap::Create (Tcl_Interp *pInterp, const char * id, const char * path, bool start)
{
	ASSERT (id != NULL);
	OptclObj *pObj = NULL,
			 *ptmp = NULL;

	pObj = new OptclObj ();
	if (!pObj->Create(pInterp, id, path, start)) {
		delete pObj;
		return NULL;
	}
	
	IUnknown **u = (IUnknown**)(IUnknown*)(*pObj);
	if (m_unkmap.find(u, &ptmp) != NULL) {
		ASSERT (ptmp != NULL);
		delete pObj;
		Lock (ptmp);
		return ptmp;
	}

	m_unkmap.set (u, pObj); 
	m_namemap.set (*pObj, pObj); // implicit const char * cast
	Lock(pObj);
	CreateCommand (pObj);
	return pObj;
}


/*
 *-------------------------------------------------------------------------
 * ObjMap::Add --
 *	Given an IUnknown pointer, this function ensures that the object table
 *	has a representation for it. If one cannot be found, then a new 
 *	representation is created, and the object command is created in the 
 *	specified interpreter.
 *
 * Result:
 *	Non-null pointer to the internal representation iff successful.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
OptclObj *	ObjMap::Add (Tcl_Interp *pInterp, LPUNKNOWN punk, ITypeInfo *pti)
{
	ASSERT (punk != NULL);
	CComPtr<IUnknown> t_unk;
	HRESULT hr;
	OptclObj *pObj = NULL;

	// get the objects pure IUnknown interface (punk can
	// point to any interface

	hr = punk->QueryInterface (IID_IUnknown, (void**)(&t_unk));
	CHECKHR(hr);
	IUnknown ** u = (IUnknown **)(IUnknown*)t_unk;

	if (m_unkmap.find(u, &pObj) == NULL) {
		pObj = new OptclObj();
		if (!pObj->Attach(pInterp, punk, pti))
		{
			delete pObj;
			pObj = NULL;
		}
		m_namemap.set(*pObj, pObj);
		m_unkmap.set(u, pObj);
		
		CreateCommand (pObj);
	}
	Lock(pObj);
	ASSERT (pObj != NULL);
	return pObj;
}



/*
 *-------------------------------------------------------------------------
 * ObjMap::Find --
 *	Given and IUnknown pointer, this function attempts to bind to an 
 *	existing representation within the object table.
 *
 * Result:
 *	A non-null pointer to the required Optcl object, iff successful.
 *
 * Side effects:
 *	None.
 *
 *-------------------------------------------------------------------------
 */
OptclObj *ObjMap::Find (LPUNKNOWN punk)
{
	ASSERT (punk != NULL);
	CComPtr<IUnknown> t_unk;
	HRESULT hr;
	OptclObj *pObj = NULL;

	// get the objects pure IUnknown interface (punk can
	// point to any interface

	hr = punk->QueryInterface (IID_IUnknown, (void**)(&t_unk));
	CHECKHR(hr);
	IUnknown **u = (IUnknown**)(IUnknown*)(t_unk);
	m_unkmap.find (u, &pObj);
	return pObj;
}



/*
 *-------------------------------------------------------------------------
 * ObjMap::Find --
 *	Finds an existing optcl object keyed on its name.
 *
 * Result:
 *	A non-null pointer to the required Optcl object, iff successful.
 *
 * Side effects:
 *	None.
 *
 *-------------------------------------------------------------------------
 */
OptclObj *ObjMap::Find (const char *name)
{
	ASSERT (name != NULL);
	OptclObj *pObj = NULL;
	m_namemap.find (name, &pObj);
	return pObj;
}





/*
 *-------------------------------------------------------------------------
 * ObjMap::DeleteCommand --
 *	Ensures that the object command associated with a valid Optcl object
 *	is quietly removed.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	po->m_cmdtoken is set to NULL.
 *-------------------------------------------------------------------------
 */
void ObjMap::DeleteCommand (OptclObj *po)
{
	ASSERT (po != NULL);
	
	if (po->m_cmdtoken == NULL) 
		return;
	
	
	const char * cmdname = Tcl_GetCommandName (po->m_pInterp, po->m_cmdtoken);
	if (cmdname == NULL)
		return;
	Tcl_CmdInfo cmdinf;

	if (Tcl_GetCommandInfo (po->m_pInterp, cmdname, &cmdinf) == 0)
		return;
	
	// modify the command info of this command so that the callback is now disabled
	cmdinf.deleteProc = NULL;
	cmdinf.deleteData = NULL;

	Tcl_SetCommandInfo (po->m_pInterp, cmdname, &cmdinf);
	Tcl_DeleteCommand (po->m_pInterp, cmdname);
	po->m_cmdtoken = NULL;
}


/*
 *-------------------------------------------------------------------------
 * ObjMap::Delete --
 *	Deletes an Optcl object, ensuring the removal of its object command,
 *	and its entries in the object table.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *
 *-------------------------------------------------------------------------
 */
void ObjMap::Delete (OptclObj *pObj)
{
	ASSERT (pObj != NULL);
	TRACE("Deleting: ");
	TRACE_OPTCLOBJ(pObj);
	// first ensure that we delete the objects command
	DeleteCommand(pObj);	
	m_namemap.delete_entry (*pObj);
	m_unkmap.delete_entry ((IUnknown**)(IUnknown*)(*pObj));
	delete pObj;
}



/*
 *-------------------------------------------------------------------------
 * ObjMap::Delete --
 *	Deletes an optcl object keyed on its name. Ensures the removal of the 
 *	object command, as well as its reference in the object table.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *
 *-------------------------------------------------------------------------
 */
void ObjMap::Delete (const char * name)
{
	ASSERT (name != NULL);
	OptclObj *pObj = NULL;

	if (m_namemap.find (name, &pObj) == NULL)
		return;
	ASSERT (pObj != NULL);
	ASSERT (strcmp(name, *pObj) == 0);
	Delete (pObj);
}



/*
 *-------------------------------------------------------------------------
 * ObjMap::Lock --
 *	Increments the reference count on an optcl object.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *
 *-------------------------------------------------------------------------
 */
void ObjMap::Lock (OptclObj *po)
{
	ASSERT (po != NULL);
	++po->m_refcount;
	TRACE_OPTCLOBJ(po);
}


/*
 *-------------------------------------------------------------------------
 * ObjMap::Unlock --
 *	Decrements the reference count on an optcl object. If zero, the object
 *	is deleted. These functions could do with thread safety in the future.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *
 *-------------------------------------------------------------------------
 */
void ObjMap::Unlock(OptclObj *po)
{
	ASSERT (po != NULL);
	--(po->m_refcount);
	TRACE_OPTCLOBJ(po);
	if (po->m_refcount == 0)
		Delete (po);
}




/*
 *-------------------------------------------------------------------------
 * ObjMap::Lock --
 *	Increments the reference count on an optcl object, keyed on its name
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *
 *-------------------------------------------------------------------------
 */
bool ObjMap::Lock (const char *name)
{
	ASSERT (name != NULL);
	OptclObj *pObj = NULL;
	if (m_namemap.find (name, &pObj) == NULL)
		return false;
	Lock (pObj);
	return true;
}






/*
 *-------------------------------------------------------------------------
 * ObjMap::Unlock --
 *	Decrements the reference count of an optcl object, keyed on its name. 
 *	If zero, the object is deleted.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *
 *-------------------------------------------------------------------------
 */
bool ObjMap::Unlock(const char *name)
{
	ASSERT (name != NULL);
	OptclObj *pObj = NULL;
	if (m_namemap.find (name, &pObj) == NULL)
		return false;
	Unlock(pObj);
	return true;
}



/*
 *-------------------------------------------------------------------------
 * ObjMap::CreateCommand --
 *	Creates the command that is to be associated with an optcl object.
 *	The command is created within the interpreter referenced by the object.
 *	The command token is stored within the object.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *
 *-------------------------------------------------------------------------
 */
void ObjMap::CreateCommand(OptclObj * pObj)
{
	ASSERT (pObj != NULL);
	pObj->m_cmdtoken = 
		Tcl_CreateObjCommand (pObj->m_pInterp, (char*)(const char*)(*pObj), 
				ObjMap::OnCmd, (ClientData)pObj, ObjMap::OnCmdDelete);
}


/*
 *-------------------------------------------------------------------------
 * ObjMap::OnCmd --
 *	Function called from tcl whenever an optcl object command is invoked.
 *	The ClientData is the pointer to the object.
 *
 * Result:
 *	Std Tcl results.
 *
 * Side effects:
 *	Anything, depending on the invocation
 *
 *-------------------------------------------------------------------------
 */
TCL_CMDEF(ObjMap::OnCmd)
{
	OptclObj *po = (OptclObj*)cd; // cast the client data to the underlying 
								  // object
	ASSERT (po != NULL);
	return (po->InvokeCmd (pInterp, objc-1, objv+1))?TCL_OK:TCL_ERROR;
}


/*
 *-------------------------------------------------------------------------
 * ObjMap::OnCmdDelete --
 *	Called when (and only when) a script deletes an object command. The
 *	referenced optcl object is also destroyed.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *
 *-------------------------------------------------------------------------
 */
void ObjMap::OnCmdDelete (ClientData cd)
{
	OptclObj *po = (OptclObj*) cd;
	ASSERT (po != NULL);
	g_objmap.Delete(po);
}





