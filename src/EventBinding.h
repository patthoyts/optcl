/*
 *------------------------------------------------------------------------------
 *	eventbinding.h
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

#if !defined(AFX_EVENTBINDING_H__818C3160_57FC_11D3_86E8_0000B482A708__INCLUDED_)
#define AFX_EVENTBINDING_H__818C3160_57FC_11D3_86E8_0000B482A708__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

struct BindingProps;
class OptclObj;


typedef THash<DISPID, BindingProps*> DispCmdTbl;



struct BindingProps
{
	TObjPtr			m_pScript;
	Tcl_Interp	*	m_pInterp;

	BindingProps (Tcl_Interp *pInterp, Tcl_Obj * pScript) 
	{
		ASSERT (pInterp != NULL && pScript != NULL);
		m_pInterp = pInterp;
		m_pScript = pScript;
	}

	int Eval (OptclObj *pObj, DISPPARAMS *pDispParams, LPVARIANT pVarResult, 
		LPEXCEPINFO pExcepInfo);
};



class EventBindings : public IDispatch
{
public:
	friend OptclObj;

	EventBindings(OptclObj *pObj, REFGUID guid, ITypeInfo *pInfo);
	virtual ~EventBindings();


	bool SetBinding (Tcl_Interp *pInterp, const char * name, Tcl_Obj *pCommand);
	bool GetBinding (Tcl_Interp *pInterp, const char * name);
	bool DeleteBinding (Tcl_Interp *pInterp, const char * name);
	
	ULONG TotalBindings ();

	// IUnknown Entries
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void ** ppvObject);
    ULONG STDMETHODCALLTYPE AddRef( void);
    ULONG STDMETHODCALLTYPE Release( void);

	// IDispatch Entries
	HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo);

	HRESULT STDMETHODCALLTYPE 
	GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo ** ppTInfo);

	HRESULT STDMETHODCALLTYPE 
	GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames, 
	              UINT cNames, LCID lcid, DISPID *rgDispId);

	HRESULT STDMETHODCALLTYPE 
	Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
	       WORD wFlags, DISPPARAMS *pDispParams, LPVARIANT pVarResult, 
		   LPEXCEPINFO pExcepInfo, UINT *puArgErr);

protected:
	void	DeleteTbl();
	bool	Name2Dispid (Tcl_Interp *pInterp, const char * name, DISPID &dispid);

protected:
	LONG					m_ref;		// COM reference count for this event binding
	ULONG					m_bindings;	// total number of bindings in this event object
	CComPtr<ITypeInfo>		m_pti;		// the type information that we are going to be binding
	CComPtr<ITypeComp>		m_ptc;		// fast access to members
	DispCmdTbl				m_cmdtbl;	// mapping of dispatch ids to Tcl commands
	OptclObj		*		m_optclobj; // the parent object of this binding
	DWORD					m_cookie;	// cookie used for event advise
	GUID					m_guid;		// the id for the event interface
};

#endif // !defined(AFX_EVENTBINDING_H__818C3160_57FC_11D3_86E8_0000B482A708__INCLUDED_)
