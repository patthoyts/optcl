/*
 *------------------------------------------------------------------------------
 *	optcltypeattr.h
 *	Definition of the OptclTypeAttr class, a wrapper for the TYPEATTR 
 *	pointer type.
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

#if !defined(AFX_OPTCLTYPEATTR_H__5826EED2_5FA7_11D3_86E8_0000B482A708__INCLUDED_)
#define AFX_OPTCLTYPEATTR_H__5826EED2_5FA7_11D3_86E8_0000B482A708__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

class OptclTypeAttr  
{
public:
	CComPtr<ITypeInfo>	m_pti;
	TYPEATTR	*		m_pattr;

public:
	OptclTypeAttr();
	virtual ~OptclTypeAttr();
	HRESULT GetTypeAttr ();
	void	ReleaseTypeAttr ();
	OptclTypeAttr & operator= (ITypeInfo *pti);
	TYPEATTR * operator -> ();
};

#endif // !defined(AFX_OPTCLTYPEATTR_H__5826EED2_5FA7_11D3_86E8_0000B482A708__INCLUDED_)
