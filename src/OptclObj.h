/*
 *------------------------------------------------------------------------------
 *	optclobj.cpp
 *	Declares the functionality for the internal representation of 
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

#if !defined(AFX_OPTCLOBJ_H__8A11BC04_616B_11D4_8004_0040055861F2__INCLUDED_)
#define AFX_OPTCLOBJ_H__8A11BC04_616B_11D4_8004_0040055861F2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

// forward declarations of used classes
#include "container.h"
#include <string>
#include "optcltypeattr.h"

class ObjMap;
class EventBindings;
class OptclBindPtr;
class DispParams;

typedef THash<GUID, EventBindings*> EventBindingsTbl;


class OptclObj {
friend ObjMap;
friend CContainer;


public:
	OptclObj ();
	virtual ~OptclObj ();

	bool	Create (Tcl_Interp *pInterp, const char *strid, const char *windowpath, bool start);
	bool	Attach (Tcl_Interp *pInterp, LPUNKNOWN punk, ITypeInfo *pti = NULL);

	operator LPUNKNOWN();
	operator const char * ();
	
	void	CoClassName (TObjPtr &pObj);
	void	InterfaceName (TObjPtr &pObj);
	void	SetInterfaceName (TObjPtr &pObj);
	HRESULT SetInterfaceFromType (ITypeInfo *pinfo);

	bool	InvokeCmd (Tcl_Interp *pInterp, int objc, Tcl_Obj *CONST objv[]);

	bool	ResolvePropertyObject (Tcl_Interp *pInterp, const char *sname, 
								   IUnknown **ppunk, ITypeInfo **ppinfo, ITypeComp **ppcmp);

	bool	GetBinding (Tcl_Interp *pInterp, char *name);
	bool	SetBinding (Tcl_Interp *pInterp, char *name, Tcl_Obj *command);
	
	bool	GetState (Tcl_Interp *pInterp);
	


protected:	// methods
	void	CreateName (LPUNKNOWN punk);
	void	InitialiseUnknown (LPUNKNOWN punk);
	void	InitialiseClassInfo (LPUNKNOWN punk);
	void	InitialisePointers (LPUNKNOWN punk, ITypeLib *pLib = NULL, ITypeInfo *pinfo = NULL);
	void	CreateCommand();
	HRESULT InitialisePointersFromCoClass ();
	HRESULT	GetTypeAttr();
	void	ReleaseTypeAttr();
	void	ReleaseBindingTable();



	bool	BuildParamsWithBindPtr (Tcl_Interp *pInterp, int objc, Tcl_Obj *CONST objv[], 
									OptclBindPtr & bp,  DispParams & dp);
	bool	RetrieveOutParams (Tcl_Interp *pInterp, int objc, Tcl_Obj *CONST objv[],
								  OptclBindPtr & bp,  DispParams & dp);
	
	bool	InvokeNoTypeInfVariant (Tcl_Interp *pInterp, long ik, int objc, Tcl_Obj *CONST objv[], 
									IDispatch *pDisp, VARIANT &varResult);
	bool	InvokeNoTypeInf (Tcl_Interp *pInterp, long ik, int objc, Tcl_Obj *CONST objv[], 
							 IDispatch *pDisp);

	bool	InvokeWithTypeInfVariant (Tcl_Interp *pInterp, long invokekind,
								  int objc, Tcl_Obj *CONST objv[], 
								  IUnknown *pUnk, ITypeInfo *pti, ITypeComp *pCmp, 
								  VARIANT &varResult, ITypeInfo **ppResultInfo = NULL);
	bool	InvokeWithTypeInf (Tcl_Interp *pInterp, long ik, int objc, Tcl_Obj *CONST objv[], 
							   IUnknown *pUnk, ITypeInfo *pti, ITypeComp *pcmp);

	bool	CheckInterface (Tcl_Interp *pInterp);

	bool	SetProp (Tcl_Interp *pInterp, int paircount, Tcl_Obj * CONST namevalues[], 
					 IUnknown *punk, ITypeInfo *pti, ITypeComp *ptc);

	bool	GetProp (Tcl_Interp *pInterp, Tcl_Obj *name, IUnknown *punk, ITypeInfo *pti, ITypeComp *ptc);
	bool	GetIndexedVariant (Tcl_Interp *pInterp, Tcl_Obj *name, 
			  IUnknown *punk, ITypeInfo *pti, ITypeComp *ptc, VARIANT &varResult, ITypeInfo **ppResultInfo);

	bool	GetPropVariantDispatch (Tcl_Interp *pInterp, const char*name, 
									IDispatch * pcurrent, VARIANT &varResult);

	bool	FindEventInterface (Tcl_Interp *pInterp, const char * lib, const char * type,
								ITypeInfo **ppinfo, GUID * pguid);

	bool	FindDefaultEventInterface (Tcl_Interp *pInterp, ITypeInfo **ppinfo, GUID *pguid);

	void	ContainerWantsToDie ();
protected:	// properties
	CComPtr<IUnknown>		m_pcurrent;	// Current interface
	CComPtr<IUnknown>		m_punk;		// the 'true' IUnknown; reference purposes only
	CComPtr<ITypeLib>		m_ptl;		// the type library for this object
	CComPtr<ITypeInfo>		m_pti;		// the type interface for the current interface
	CComPtr<ITypeComp>		m_ptc;		// the type info's compiler interface
	CComPtr<ITypeInfo>		m_pti_class;// the type interface for the this coclass
public:
	OptclTypeAttr 			m_pta;		// the type attribute for the current typeinfo
protected:
	std::string				m_name;
	unsigned long			m_refcount;	// reference count of this optcl object
	Tcl_Interp		*		m_pInterp;	// interpreter that created this object
	Tcl_Command				m_cmdtoken;	// command token of the tcl command within the above interpreter
	EventBindingsTbl		m_bindings; // bindings for event interfaces of this object
	CContainer				m_container;// container
	bool					m_destroypending; // true during a final delete operation
};


#endif // !defined(AFX_OPTCLOBJ_H__8A11BC04_616B_11D4_8004_0040055861F2__INCLUDED_)
