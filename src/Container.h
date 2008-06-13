/*
 *------------------------------------------------------------------------------
 *	container.cpp
 *	Declaration of the CContainer class, providing functionality for
 *	a Tk activex container widget.
 *	1999-01-26 created
 *	1999-08-25 modified for use in Optcl
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

#if !defined(AFX_CONTAINER_H__C07038C0_9445_11D2_86E7_0000B482A708__INCLUDED_)
#define AFX_CONTAINER_H__C07038C0_9445_11D2_86E7_0000B482A708__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


class OptclObj;


class CContainer  
{
public: // constructors
	CContainer(OptclObj *pObj);
	virtual ~CContainer();

public: // non-static methods
	IUnknown * Create (Tcl_Interp *pInterp, Tk_Window tkParent, const char * path, const char *id);

public: // static methods
	static void ContainerEventProc (ClientData cd, XEvent *pEvent);
	static void DeleteContainer (char *pObject);
	static LRESULT CALLBACK WidgetSubclassProc (HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
protected: // static methods
	static int WidgetCmd (ClientData cd, Tcl_Interp *pInterp, int objc, Tcl_Obj *CONST objv[]);
	
protected: // non-static methods
	int ConfigInfo (Tcl_Interp *pInterp);
	int ConfigInfo (Tcl_Interp *pInterp, char *pProperty);
	int ConfigInfo (Tcl_Interp *pInterp, int objc, Tcl_Obj *CONST pArgs[]);
	int SetProperty (Tcl_Interp *pInterp, Tcl_Obj *pProperty, Tcl_Obj *pValue, bool &bChanged);
	bool CreateControl (Tcl_Interp *pInterp, const char *id);
	void InitFromObject ();
	int CreateTkWindow(Tk_Window tkParent, char *path);

protected: // members variables
	Tk_Window		m_tkWindow;
	Tcl_Interp	*	m_pInterp;
	Display		*	m_tkDisplay;
	Tcl_Command		m_widgetCmd;
	HWND			m_hTkWnd;
	DWORD			m_height;
	DWORD			m_width;
	LONG			m_windowproc;
	bool			m_bDestroyPending;
	OptclObj	*	m_optclobj;

	// Com pointers
	CComPtr<IUnknown>			m_pUnk;		// pointer to the contained object
	CComPtr<IUnknown>			m_pUnkHost;	// pointer to the host IUnknown
	
	// QI ptrs that have the IDD-templatised versions
	CComQIPtr<IOleObject>								m_pObj;
	CComQIPtr<IOleInPlaceObject>						m_pInPlaceObj;
	CComQIPtr<IOleWindow>								m_pOleWnd;
	CComQIPtr<IOleControlSite>							m_pSite;
	CComQIPtr<IAxWinHostWindow, &IID_IAxWinHostWindow>	m_pHost;	// pointer to the host (client site) object
protected: // static member variables
	
	static const char *	m_propname;
};

#endif // !defined(AFX_CONTAINER_H__C07038C0_9445_11D2_86E7_0000B482A708__INCLUDED_)
