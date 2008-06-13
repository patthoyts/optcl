/*
 *------------------------------------------------------------------------------
 *	container.cpp
 *	Implementation of the CContainer class, providing functionality for
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



#include "stdafx.h"
#include "tbase.h"
#include "optcl.h"
#include "utility.h"
#include "Container.h"
#include "optclobj.h"



const char *	CContainer::m_propname = "container";

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////
/*
 *-------------------------------------------------------------------------
 * CContainer::CContainer() --
 *	Constructor
 * Result:
 *	None
 * Side effects:
 *	Members set to default values - initial height and width information
 *	stored here.
 *-------------------------------------------------------------------------
 */
CContainer::CContainer(OptclObj *parent) :
m_tkWindow(NULL),
m_widgetCmd(NULL),
m_pInterp(NULL),
m_height(200),
m_width(200),
m_windowproc(NULL),
m_bDestroyPending(false),
m_optclobj(parent)
{
	ASSERT (parent != NULL);
}

/*
 *-------------------------------------------------------------------------
 * CContainer::~CContainer() --
 *	Destructor
 * Result:
 *	None
 * Side effects:
 *	Tk window is requested to be destroyed. 
 *	COM resources release (except for the control container which is release
 *	when the Tk window is actually destroyed.
 *-------------------------------------------------------------------------
 */

CContainer::~CContainer()
{
	// close down the references to the object
	m_bDestroyPending = true;
	m_pUnk.Release();
	m_pObj.Release();
	m_pSite.Release();
	m_pInPlaceObj.Release();
	m_pOleWnd.Release();
	m_pUnkHost.Release();

	if (m_widgetCmd != NULL) {
		if (!Tcl_InterpDeleted(m_pInterp)) 
			Tcl_DeleteCommandFromToken (m_pInterp, m_widgetCmd);
		m_widgetCmd = NULL;
	}

	if (m_tkWindow != NULL) 
	{
		// remove the subclass
		SetWindowLong (m_hTkWnd, GWL_WNDPROC, (LONG)m_windowproc);
		RemoveProp (m_hTkWnd, m_propname);
		Tk_DestroyWindow (m_tkWindow);
		m_tkWindow = NULL;
	}
	
}

/*
 *-------------------------------------------------------------------------
 * CContainer::ContainerEventProc --
 *	Called by Tk to process events
 * Result:
 *	None
 * Side effects:
 *	Lifetime and size of widget affected - focus model needs working on!
 *-------------------------------------------------------------------------
 */
void CContainer::ContainerEventProc (ClientData cd, XEvent *pEvent)
{
	CContainer *pContainer = (CContainer *)cd;
	SIZEL s, hm;
	RECT r;

	switch (pEvent->type)
	{
	case Expose:
		// Nothing required as the AxAtl window 
		// should receive its own exposure event
		break;
	case FocusIn:
		if (pContainer->m_pSite)
			pContainer->m_pSite->OnFocus(TRUE);
		/*
		hControl = ::GetWindow (pContainer->m_hTkWnd, GW_CHILD);
		if (hControl)
			::SetFocus(hControl);
		*/

		break;
	case ConfigureNotify:
		s.cx = Tk_Width(pContainer->m_tkWindow);
		s.cy = Tk_Height(pContainer->m_tkWindow);
		r.left = r.top = 0;
		r.right = s.cx;
		r.bottom = s.cy;
		

		AtlPixelToHiMetric(&s, &hm);
		if (pContainer->m_pObj)
			pContainer->m_pObj->SetExtent(DVASPECT_CONTENT, &hm);
		if (pContainer->m_pInPlaceObj)
			pContainer->m_pInPlaceObj->SetObjectRects (&r, &r);

		break;
	case DestroyNotify:
		if (!pContainer->m_bDestroyPending) {
			Tcl_EventuallyFree(cd, DeleteContainer);
			pContainer->m_tkWindow = NULL;
		}
		break;


	default:
		break;
	}

}

/*
 *-------------------------------------------------------------------------
 * CContainer::DeleteContainer --
 *	Called by ContainerEventProc, when the Tk_Window is about to be 
 *	destroyed by scripting.
 * Result:
 *	None
 * Side effects:
 *	Memory deallocated
 *-------------------------------------------------------------------------
 */
void CContainer::DeleteContainer (char *pObject)
{
	CContainer *pContainer = (CContainer*)pObject;
	pContainer->m_optclobj->ContainerWantsToDie();
}



/*
 *-------------------------------------------------------------------------
 * CContainer::Create --
 *	Called by the related object in order to create the window.
 *	tkParent in a parent of the window to be created. The string
 *	pointed to by 'id' is the clsid/progid/documentpath of this object.
 *
 * Result:
 *	NULL iff failed to be created (pInterp will store descriptive result)
 *
 * Side effects:
 *	Depends on object being created.
 *-------------------------------------------------------------------------
 */
IUnknown * CContainer::Create (Tcl_Interp *pInterp, Tk_Window tkParent, 
							   const char * widgetpath, const char *id)
{
	m_pInterp = pInterp;
	char *path;
	if (TCL_ERROR == CreateTkWindow (tkParent, (char*)widgetpath))
		return NULL;

	path = Tk_PathName(m_tkWindow);
	
	Tcl_VarEval (pInterp, "winfo id ", path, (char*)NULL);

	int iParent;
	Tcl_GetIntFromObj (pInterp, Tcl_GetObjResult (pInterp), &iParent);
	m_hTkWnd = (HWND) iParent;
	SetProp (m_hTkWnd, m_propname, (HANDLE)this);

	if (!CreateControl(pInterp, id))
		return NULL;

	InitFromObject ();
	

	// subclass this window (once again, since ATL has already hooked it), in order
	// to correctly handle mouse messages and destruction
	m_windowproc = SetWindowLong (m_hTkWnd, GWL_WNDPROC, (LONG)WidgetSubclassProc);

	// Set up the height and width accordingly
	Tk_GeometryRequest (m_tkWindow, m_width, m_height);


	m_widgetCmd = Tcl_CreateObjCommand (m_pInterp, path, WidgetCmd, 
										(ClientData)this, NULL);

	Tcl_SetResult (m_pInterp, path, TCL_STATIC);
	return m_pUnk;
}




/*
 *-------------------------------------------------------------------------
 * CContainer::WidgetCmd --
 *	Static class method that is called by Tcl when invoking the widget
 *	command.
 * Result:
 *	TCL_OK if command execute ok; else TCL_ERROR
 * Side effects:
 *	Dependant on the subcommand
 *-------------------------------------------------------------------------
 */
int CContainer::WidgetCmd (	ClientData cd, Tcl_Interp *pInterp, int objc, 
							Tcl_Obj *CONST objv[] )
{
	if (objc < 2) {
		Tcl_AppendResult (pInterp, "wrong # args: should be \"", 
			Tcl_GetStringFromObj (objv[0], NULL), " option ?arg arg ...?\"", (char*)NULL);
		return TCL_ERROR;
	}
	char *szCommand = Tcl_GetStringFromObj(objv[1], NULL);
	int nLength = strlen(szCommand);
	CContainer *pWidget = (CContainer*)cd;

	if (strncmp (szCommand, "configure", nLength) == 0) {
		switch (objc) {
		case 2:
			return pWidget->ConfigInfo (pInterp);
			break;
		case 3:
			return pWidget->ConfigInfo (pInterp, Tcl_GetStringFromObj(objv[2], NULL));
			break;
		default:		
			return pWidget->ConfigInfo (pInterp, objc - 2, objv + 2); 
			break;
		}
	}

	if (strncmp (szCommand, "cget", nLength) == 0) {
		if (objc == 3) {
			return pWidget->ConfigInfo (pInterp, Tcl_GetStringFromObj(objv[2], NULL));
		} else {
			Tcl_AppendResult (pInterp, "wrong # args: should be \"", 
				Tcl_GetStringFromObj (objv[0], NULL), " cget arg\"", (char*)NULL);
			return TCL_ERROR;
		}
	}
	
	Tcl_AppendResult (pInterp, "urecognised command: ", szCommand, (char*)NULL);
	return TCL_ERROR;
}

/*
 *-------------------------------------------------------------------------
 * CContainer::ConfigInfo (Tcl_Interp *pInterp) --
 *	Overloaded method that returns the value of all configuration options
 *	for the widget
 * Result:
 *	TCL_OK
 * Side effects:
 *	New Tcl result
 *-------------------------------------------------------------------------
 */
int CContainer::ConfigInfo (Tcl_Interp *pInterp)
{
	Tcl_DString dstring;
	Tcl_DStringInit(&dstring);

	Tcl_ResetResult(pInterp);
	ConfigInfo (pInterp, "-width");
	Tcl_DStringAppendElement(&dstring, Tcl_GetStringResult (pInterp));

	Tcl_ResetResult(pInterp);
	ConfigInfo (pInterp, "-height");
	Tcl_DStringAppendElement(&dstring, Tcl_GetStringResult (pInterp));

	Tcl_SetResult(pInterp, dstring.string, TCL_VOLATILE);
	Tcl_DStringFree(&dstring);
	return TCL_OK;
}

/*
 *-------------------------------------------------------------------------
 * CContainer::ConfigInfo (Tcl_Interp *pInterp, char *pProperty) --
 *	Overloaded method that provides the value of a given configuration
 *	option.
 * Result:
 *	TCL_OK iff configuration option exists; else TCL_ERROR
 * Side effects:
 *	New Tcl result.
 *-------------------------------------------------------------------------
 */
int CContainer::ConfigInfo (Tcl_Interp *pInterp, char *pProperty)
{
	bool bFound = false;
	if (strcmp(pProperty, "-width")==0) {
		bFound = true;
		Tcl_SetObjResult (pInterp, Tcl_NewIntObj (m_width));
	}

	else if (strcmp(pProperty, "-height")==0) {
		bFound = true;
		Tcl_SetObjResult (pInterp, Tcl_NewIntObj (m_height));
	}
	
	else if (strcmp(pProperty, "-takefocus")==0) {
		bFound = true;
		Tcl_SetObjResult (pInterp, Tcl_NewBooleanObj(1)); 
	}

	if (!bFound) {
		Tcl_ResetResult (pInterp);
		Tcl_AppendResult (pInterp, "unknown option \"", pProperty, (char*)NULL);
		return TCL_ERROR;
	}
	return TCL_OK;
}

/*
 *-------------------------------------------------------------------------
 * CContainer::ConfigInfo (Tcl_Interp *pInterp, int objc, Tcl_Obj *CONST pArgs[]) --
 *	Overloaded method; used to set the value of a number of options
 * Result:
 *	TCL_OK iff all specified options are set ok; else TCL_ERROR
 * Side effects:
 *	Change in options may have an effect on the size and viewing of the 
 *	widget
 *-------------------------------------------------------------------------
 */
int CContainer::ConfigInfo (Tcl_Interp *pInterp, int objc, Tcl_Obj *CONST pArgs[])
{
	if (objc % 2 == 0) 
	{
		bool bChanged = false;
		for (int i = 0; i < objc; i += 2) {
			if (SetProperty (pInterp, pArgs[i], pArgs[i+1], bChanged) != TCL_OK)
				return TCL_ERROR;
		}
		if (bChanged) 
		{
			Tk_GeometryRequest(m_tkWindow, m_width, m_height);
		}
		return TCL_OK;
	}
	else  // # of values != # of options
	{
		char *szLast = Tcl_GetStringFromObj(pArgs[objc-1], NULL);
		Tcl_ResetResult(pInterp);
		Tcl_AppendResult (pInterp, "unknown option \"", szLast, "\"", (char*)NULL);
		return TCL_ERROR;
	}
}

/*
 *-------------------------------------------------------------------------
 * CContainer::SetProperty --
 *	Sets a single option with a new value.
 * Result:
 *	TCL_OK iff option set ok; else TCL_ERROR
 * Side effects:
 *	Size and other viewing factors may change for widget
 *-------------------------------------------------------------------------
 */
int CContainer::SetProperty (Tcl_Interp *pInterp, Tcl_Obj *pProperty, Tcl_Obj *pValue, bool &bChanged)
{
	char *szProperty = Tcl_GetStringFromObj (pProperty, NULL);
	int value;

	if (strcmp(szProperty, "-width")==0) {
		if (Tcl_GetIntFromObj (pInterp, pValue, &value) == TCL_ERROR)
			return TCL_ERROR;
		m_width = abs(value);
		bChanged = true;
	}

	else if (strcmp(szProperty, "-height")==0) {
		if (Tcl_GetIntFromObj (pInterp, pValue, &value) == TCL_ERROR)
			return TCL_ERROR;
		m_height = abs(value);
		bChanged = true;
	}

	else {
		Tcl_ResetResult (pInterp);
		Tcl_AppendResult (pInterp, "unknown option \"", szProperty, (char*)NULL);
		return TCL_ERROR;
	}

	return TCL_OK;
}

/*
 *-------------------------------------------------------------------------
 * CContainer::WidgetSubclassProc --
 *	This is used to subclass the main window to handle proper forwarding 
 *	of mouse capture, release of the control container, and the release of 
 *	OLE resources.
 * Result:
 *	The return of the subclassed window procedure
 * Side effects:
 *	Mouse capture affected - COM interfaces destroyed, OLE uninitialised by one
 *-------------------------------------------------------------------------
 */
LRESULT CALLBACK CContainer::WidgetSubclassProc (HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CContainer *pContainer = (CContainer*) GetProp (hwnd, m_propname);
	if (pContainer == NULL) return 0;
	WORD fwEvent = LOWORD (wParam);
	WORD idChild = HIWORD (wParam);
	HWND hCurrentFocus = GetFocus();

	switch (uMsg) {
	case WM_MOUSEACTIVATE:
		return MA_ACTIVATE;
		break;
	
	case WM_NCCREATE:
		pContainer->m_pHost.Release();
		break;
	}

	return CallWindowProc ((WNDPROC)(pContainer->m_windowproc), hwnd, uMsg, wParam, lParam);
}


/*
 *-------------------------------------------------------------------------
 * CContainer::InitFromObject --
 *	Initialises interface pointers to the underlying control, and site
 * Result:
 *	None
 * Side effects:
 *	COM memory allocation of vtables etc.
 *-------------------------------------------------------------------------
 */
void CContainer::InitFromObject ()
{
	AtlAxGetControl(m_hTkWnd, &m_pUnk);
	m_pObj = m_pUnk;
	m_pInPlaceObj = m_pUnk;
	m_pOleWnd = m_pUnk;
	m_pSite = m_pUnkHost;
}


/*
 *-------------------------------------------------------------------------
 * CContainer::CreateTkWindow --
 *	Called by the Create member function to create the Tk window
 * Result:
 *	returns TCL_OK iff window created
 * Side effects:
 *
 *-------------------------------------------------------------------------
 */
int CContainer::CreateTkWindow(Tk_Window tkParent, char *path)
{
	// create the window, specifying that it is a child
	m_tkWindow = Tk_CreateWindowFromPath(m_pInterp, tkParent, path, NULL);
	if (m_tkWindow == NULL)
		return TCL_ERROR;

	Tk_SetClass(m_tkWindow, "Container");
	m_tkDisplay = Tk_Display(m_tkWindow);
		
	Tk_CreateEventHandler (m_tkWindow,	
						   StructureNotifyMask | 
						   ExposureMask | 
						   FocusChangeMask,
						   ContainerEventProc, (ClientData)this);
	
	Tk_MakeWindowExist(m_tkWindow);
	return TCL_OK;
}



/*
 *-------------------------------------------------------------------------
 * CContainer::CreateControl --
 *	Using Atl, creates the control. Ensures that the children of the 
 *	control can be navigated with the keyboard.
 *
 * Result:
 *	true iff succeeded. pInterp will hold a description of the error.
 *
 * Side effects:
 *	Depends on the object being created
 *-------------------------------------------------------------------------
 */
bool CContainer::CreateControl (Tcl_Interp *pInterp, const char *id)
{
	USES_CONVERSION;
	HRESULT hr = E_FAIL;
	HWND	hWndChild;
	LPOLESTR oleid = A2OLE(id);

	hr = AtlAxCreateControl (oleid, m_hTkWnd, NULL, &m_pUnkHost);
	if (FAILED(hr)) 
		goto error;

	m_pHost = m_pUnkHost;

	if (m_pHost == NULL) {
		hr = E_NOINTERFACE;
		goto error;
	}


	// check for control parent style if control has a window
	hWndChild = ::GetWindow(m_hTkWnd, GW_CHILD);
	if(hWndChild != NULL)
	{
		if(::GetWindowLong(hWndChild, GWL_EXSTYLE) & WS_EX_CONTROLPARENT)
		{
			DWORD dwExStyle = ::GetWindowLong(m_hTkWnd, GWL_EXSTYLE);
			dwExStyle |= WS_EX_CONTROLPARENT;
			::SetWindowLong(m_hTkWnd, GWL_EXSTYLE, dwExStyle);
		}
	}
	return true;
error:
	Tcl_SetResult (pInterp, HRESULT2Str(hr), TCL_STATIC);
	return false;
}



