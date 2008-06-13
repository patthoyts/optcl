/*
 *------------------------------------------------------------------------------
 *	dispparams.h
 *	Declaration of the DispParams class, a wrapper for the DISPPARAMS
 *	Automation type.
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

#if !defined(AFX_DISPPARAMS_H__BF3EF6CA_73B0_11D4_8004_0040055861F2__INCLUDED_)
#define AFX_DISPPARAMS_H__BF3EF6CA_73B0_11D4_8004_0040055861F2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000



class DispParams : public DISPPARAMS  
{
public:
	DispParams();
	virtual ~DispParams();

	void Release ();
	void Args (UINT args, UINT named = 0);
	void SetDISPID (UINT named, DISPID id);
	VARIANTARG &operator[] (UINT arg);


	/*
	 *-------------------------------------------------------------------------
	 * Set --
	 *	Template function that sets the value of a variant at index 'index'
	 *	to a template-type value. This works because the type T should have
	 *	an appropriate casting operator to fit those of VC6's _variant_t
	 * Result:
	 *	None.
	 * Side effects:
	 *	None.
	 *-------------------------------------------------------------------------
	 */
	template <class T> void Set (UINT index, T value)
	{
		_variant_t t;
		VARIANTARG *pref = &(operator[](index));
		t.Attach (*pref);
		t = value;
		*pref = t.Detach();
	}

	void Set (UINT index, VARIANT * pv);
};

#endif // !defined(AFX_DISPPARAMS_H__BF3EF6CA_73B0_11D4_8004_0040055861F2__INCLUDED_)

