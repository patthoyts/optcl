/*
 *------------------------------------------------------------------------------
 *	optclbindptr.h
 *	Defines the class used wrapping a BINDPTR, DESCKIND and ITypeInfo
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
 */#if !defined(AFX_OPTCLBINDPTR_H__2682D1C3_5EDC_11D3_86E8_0000B482A708__INCLUDED_)
#define AFX_OPTCLBINDPTR_H__2682D1C3_5EDC_11D3_86E8_0000B482A708__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000


// wrapper class for a BINDPTR, DESCKIND, and ITypeInfo
class OptclBindPtr  
{
public:
	BINDPTR				m_bp;
	DESCKIND			m_dk;
	CComPtr<ITypeInfo>	m_pti;

public:
	OptclBindPtr();
	virtual ~OptclBindPtr();
	void ReleaseBindPtr ();

	// inline functions
	MEMBERID	OptclBindPtr::memid ()
	{
		ASSERT (m_bp.lpfuncdesc != NULL);

		switch (m_dk) {
		case DESCKIND_FUNCDESC:
			return m_bp.lpfuncdesc->memid;

		case DESCKIND_IMPLICITAPPOBJ:
		case DESCKIND_VARDESC:
			return m_bp.lpvardesc->memid;
		default:
			ASSERT (FALSE);
			return DISPID_UNKNOWN;
		}
	}


	short		OptclBindPtr::cParams()
	{
		ASSERT (m_bp.lpfuncdesc != NULL);
		switch (m_dk) {
		case DESCKIND_FUNCDESC:
			return m_bp.lpfuncdesc->cParams;
		case DESCKIND_IMPLICITAPPOBJ:
		case DESCKIND_VARDESC:
			return 1;
		default:
			ASSERT (FALSE);
			return 0;
		}
	}

	short		OptclBindPtr::cParamsOpt()
	{
		ASSERT (m_bp.lpfuncdesc != NULL);
		switch (m_dk) {
		case DESCKIND_FUNCDESC:
			return m_bp.lpfuncdesc->cParamsOpt;
		case DESCKIND_IMPLICITAPPOBJ:
		case DESCKIND_VARDESC:
			return 1;
		default:
			ASSERT (FALSE);
			return 0;
		}
	}

	ELEMDESC *	OptclBindPtr::param(short param)
	{
		ASSERT (m_bp.lpfuncdesc != NULL);
		ASSERT (param < cParams());

		switch (m_dk) {
		case DESCKIND_FUNCDESC:
			return (m_bp.lpfuncdesc->lprgelemdescParam + param);
		case DESCKIND_IMPLICITAPPOBJ:
		case DESCKIND_VARDESC:
			return (&m_bp.lpvardesc->elemdescVar);
		default:
			ASSERT (FALSE);
			return 0;
		}

	}

	ELEMDESC *  OptclBindPtr::result()
	{
		ASSERT (m_bp.lpfuncdesc != NULL);

		switch (m_dk) {
		case DESCKIND_FUNCDESC:
			return (&m_bp.lpfuncdesc->elemdescFunc);
		case DESCKIND_IMPLICITAPPOBJ:
		case DESCKIND_VARDESC:
			return (&m_bp.lpvardesc->elemdescVar);
		default:
			ASSERT (FALSE);
			return 0;
		}
	}
};

#endif // !defined(AFX_OPTCLBINDPTR_H__2682D1C3_5EDC_11D3_86E8_0000B482A708__INCLUDED_)
