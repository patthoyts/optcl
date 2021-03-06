/*
 *------------------------------------------------------------------------------
 *	optclbindptr.cpp
 *	Implements the class used wrapping a BINDPTR, DESCKIND and ITypeInfo
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
#include "typelib.h"
#include "OptclBindPtr.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

OptclBindPtr::OptclBindPtr()
{
	m_bp.lpfuncdesc = NULL;
	m_dk = DESCKIND_NONE;
}

OptclBindPtr::~OptclBindPtr()
{
	ReleaseBindPtr();
}


void OptclBindPtr::ReleaseBindPtr ()
{	
	::ReleaseBindPtr(m_pti, m_dk, m_bp);
	m_dk = DESCKIND_NONE;
}




