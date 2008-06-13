/*
 *------------------------------------------------------------------------------
 *	objmap.h
 *	Definition of the object table.
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

#if !defined(AFX_OBJMAP_H__8A11BC00_616B_11D4_8004_0040055861F2__INCLUDED_)
#define AFX_OBJMAP_H__8A11BC00_616B_11D4_8004_0040055861F2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "optcl.h"
#include "OptclObj.h"

typedef THash<char, OptclObj*>		ObjNameMap;
typedef THash<IUnknown*, OptclObj*>	ObjUnkMap;

class ObjMap {
	friend OptclObj;
protected:
	ObjNameMap	m_namemap;
	ObjUnkMap	m_unkmap;
	bool		m_destructpending;

public: // constructor / destructor
	ObjMap ();
	virtual ~ObjMap ();

	OptclObj *	Create (Tcl_Interp *pInterp, const char * id, const char * path, bool start);
	OptclObj *	Add (Tcl_Interp *pInterp, LPUNKNOWN punk, ITypeInfo *pti = NULL);
	OptclObj *	Find (LPUNKNOWN punk);
	OptclObj *	Find (const char *name);

	void		Delete (const char * name);
	void		DeleteAll ();

	bool		Lock (const char *name);
	bool		Unlock(const char *name);

	void 		Lock (OptclObj *);
	void		Unlock(OptclObj *);
	

public:		// statics
	static TCL_CMDEF(OnCmd);
	static void OnCmdDelete (ClientData cd);
	
protected:
	void		Delete (OptclObj *);
	void		CreateCommand (OptclObj *);
	void		DeleteCommand (OptclObj *);
	void		ObjDump ();
};


// Global Variable Declaration!!!

extern ObjMap	g_objmap; // one object map per application

#endif // !defined(AFX_OBJMAP_H__8A11BC00_616B_11D4_8004_0040055861F2__INCLUDED_)
