/*
 *------------------------------------------------------------------------------
 *	typelib.h
 *	Declares a collection of function for accessing typelibraries. 
 *	Currently this only includes browsing facilities. 
 *	In the future, this may contain typelib building functionality.
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


#ifndef _TYPELIB_H_62518A80_624A_11d4_8004_0040055861F2
#define _TYPELIB_H_62518A80_624A_11d4_8004_0040055861F2

#include "tbase.h"

// TypeLib provides the structure that holds the main pointer to the library ITypeLib 
// interface, together with its compiler interface
struct TypeLib {
	CComPtr<ITypeLib>	m_ptl; 
	CComPtr<ITypeComp>	m_ptc;

	TypeLib (ITypeLib *ptl, ITypeComp *ptc) {
		m_ptl = ptl;
		m_ptc = ptc;
	}
};



// TypeLibsTbl - a hash table mapping library programmatic name to a TypeLib structure
// Internally it also holds a mapping from the a libraries human readable name to
// the same structure
class TypeLibsTbl : public THash<char, TypeLib*>
{
public:
				TypeLibsTbl ();
	virtual		~TypeLibsTbl ();
	void		DeleteAll ();
	ITypeLib*	LoadLib (Tcl_Interp *pInterp, const char *fullname);
	void		UnloadLib (Tcl_Interp *pInterp, const char *fullname);
	bool		IsLibLoaded (const char *fullname);
	TypeLib*	EnsureCached (ITypeLib  *pLib);
	TypeLib*	EnsureCached (ITypeInfo *pInfo);
protected: // methods
	TypeLib*	Cache (const char *szname, const char *szfullname, ITypeLib *ptl, ITypeComp *ptc);

protected: // properties
	THash <char, Tcl_HashEntry*>	m_loadedlibs; // by name
};

// globals
extern TypeLibsTbl g_libs;


void	TypeLib_GetName (ITypeLib *, ITypeInfo *, TObjPtr &pname);
void	TypeLib_ResolveName (const char *name, TypeLib **pptl, ITypeInfo **ppinfo);
void	TypeLib_ResolveName (const char * lib, const char * type, TypeLib **pptl, ITypeInfo **ppinfo);
void	ReleaseBindPtr (ITypeInfo *pti, DESCKIND dk, BINDPTR &ptr);
bool	TypeLib_ResolveConstant (Tcl_Interp *pInterp, char *fullformatname, 
								 TObjPtr &pObj, ITypeInfo *pInfo = NULL);
bool	TypeLib_ResolveConstant (Tcl_Interp *pInterp, ITypeInfo *pti, 
								 const char *member, TObjPtr &pObj);


#endif // _TYPELIB_H_62518A80_624A_11d4_8004_0040055861F2