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
	TLIBATTR	*		m_libattr;

	TObjPtr				m_progname, m_fullname, m_path;


	TypeLib () {
		m_progname.create();
		m_fullname.create();
		m_path.create();
		m_libattr = NULL;
	}

	~TypeLib () {
		if (m_libattr != NULL) {
			ASSERT (m_ptl != NULL);
			m_ptl->ReleaseTLibAttr(m_libattr);
		}	
	}

	HRESULT Init (ITypeLib *ptl, ITypeComp *ptc, const char * progname, 
				  const char * fullname, const char * path) {
		ASSERT (progname != NULL && fullname != NULL);

		m_ptl = ptl;
		m_ptc = ptc;

		m_progname = progname;
		m_fullname = fullname;
		if (path)
			m_path = path;
		else
			m_path = "???";
		return ptl->GetLibAttr(&m_libattr);
	}
};


struct TypeLibUniqueID {
	TypeLibUniqueID (const GUID & guid, WORD maj, WORD min) {
		m_guid = guid;
		m_majorver = maj;
		m_minorver = min;
	}

	GUID	m_guid;
	WORD	m_majorver;
	WORD	m_minorver;
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

	TypeLib*	LoadLib (Tcl_Interp *pInterp, const char * fullpath);
	void		UnloadLib (Tcl_Interp *pInterp, const char * progname);
	TypeLib*	TypeLibFromUID (const GUID & guid, WORD maj, WORD min);

	TypeLib*	EnsureCached (ITypeLib  *pLib);
	TypeLib*	EnsureCached (ITypeInfo *pInfo);

	char*		GetFullName (char * szProgName);
	GUID*		GetGUID (char * szProgName);

protected: // methods
	TypeLib*	Cache (Tcl_Interp * pInterp, ITypeLib *ptl, const char * path = NULL);
	HRESULT		GenerateNames (TObjPtr &progname, TObjPtr &username, ITypeLib *pLib);

protected: // properties
	THash <TypeLibUniqueID, Tcl_HashEntry*>		m_loadedlibs;	// by unique and full descriptor
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
HRESULT TypeLib_GetDefaultInterface (ITypeInfo *pti, bool bEventSource, ITypeInfo ** ppdefti);

#endif // _TYPELIB_H_62518A80_624A_11d4_8004_0040055861F2