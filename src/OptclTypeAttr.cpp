/*
 *------------------------------------------------------------------------------
 *	optcltypeattr.cpp
 *	Implementation of the OptclTypeAttr class, a wrapper for the TYPEATTR 
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

#include "stdafx.h"
#include "tbase.h"
#include "optcl.h"
#include "utility.h"
#include "OptclTypeAttr.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

OptclTypeAttr::OptclTypeAttr() : m_pattr(NULL)
{

}

OptclTypeAttr::~OptclTypeAttr()
{
	ReleaseTypeAttr();
}


HRESULT OptclTypeAttr::GetTypeAttr ()
{
	HRESULT hr = S_OK;
	// only get if we haven't already
	if (m_pattr == NULL) {
		ASSERT (m_pti != NULL);
		hr = m_pti->GetTypeAttr (&m_pattr);
	}
	return hr;
}


void OptclTypeAttr::ReleaseTypeAttr ()
{
	if (m_pattr != NULL) 
	{
		ASSERT (m_pti != NULL);
		m_pti->ReleaseTypeAttr(m_pattr);;
		m_pattr = NULL;
	}
}


OptclTypeAttr & OptclTypeAttr::operator= (ITypeInfo *pti)
{
	ReleaseTypeAttr();
	m_pti = pti;
	if (m_pti != NULL)
		GetTypeAttr();
	return *this;
}


TYPEATTR * OptclTypeAttr::operator -> ()
{
	ASSERT (m_pattr != NULL);
	return m_pattr;
}

