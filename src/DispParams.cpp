/*
 *------------------------------------------------------------------------------
 *	dispparams.cpp
 *	Implementation of the DispParams class, a wrapper for the DISPPARAMS
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

#include "stdafx.h"
#include "tbase.h"
#include "optcl.h"
#include "DispParams.h"
#include "utility.h"




//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

/*
 *-------------------------------------------------------------------------
 * DispParams::DispParams --
 *	Constructor - nulls out everything.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
DispParams::DispParams()
{
	rgvarg = NULL;
	rgdispidNamedArgs = NULL;
	cArgs = 0;
	cNamedArgs = 0;
}

/*
 *-------------------------------------------------------------------------
 * DispParams::~DispParams --
 *	Destructor - releases internal data.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
DispParams::~DispParams()
{
	Release();
}

/*
 *-------------------------------------------------------------------------
 * DispParams::Release --
 *	Releases all allocated variants. Nulls out the DISPPARAMS structure.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void DispParams::Release ()
{
	// start releasing the variants
	for (UINT i = 0; i < cArgs; i++)
	{
		ASSERT (rgvarg != NULL);
		OptclVariantClear (rgvarg+i);
	}

	delete_ptr (rgvarg);
	delete_ptr (rgdispidNamedArgs);
	cArgs = 0;
	cNamedArgs = 0;
}



/*
 *-------------------------------------------------------------------------
 * DispParams::Args --
 *	Sets up the number of arguments, both name and unnamed.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	Allocates enough memory for the dispatch id's of the named arguments.
 *
 *-------------------------------------------------------------------------
 */
void DispParams::Args (UINT args, UINT named)
{
	UINT i;

	Release();
	if (args > 0)
	{
		rgvarg = new VARIANTARG[args];
		for (i = 0; i < args; i++)
			VariantInit(rgvarg+i);
		cArgs = args;
	}

	if (named > 0)
	{
		rgdispidNamedArgs = new DISPID[named];
		for (i = 0; i < named; i++)
			rgdispidNamedArgs[i] = DISPID_UNKNOWN;
		cNamedArgs = named;
	}
}

/*
 *-------------------------------------------------------------------------
 * DispParams::SetDISPID --
 *	Sets the dispatch id of a named arguments. The argument is accessed using
 *	the index 'named'.
 *
 * Result:
 *	None.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
void DispParams::SetDISPID (UINT named, DISPID id)
{
	ASSERT (named < cNamedArgs);
	rgdispidNamedArgs[named] = id;
}


/*
 *-------------------------------------------------------------------------
 * DispParams::operator[] --
 *	operator to get direct access to an argument at a certain index.
 *
 * Result:
 *	A reference to the argument.
 *
 * Side effects:
 *	None.
 *-------------------------------------------------------------------------
 */
VARIANTARG &DispParams::operator[] (UINT arg)
{
	ASSERT (arg < cArgs);
	return rgvarg[arg];
}




/*
 *-------------------------------------------------------------------------
 * DispParams::Set --
 *
 * Result:
 *
 * Side effects:
 *
 *-------------------------------------------------------------------------
 */
void DispParams::Set (UINT index, VARIANT * pv)
{
	ASSERT (pv != NULL);
	ASSERT (index < cArgs);

	V_VARIANTREF(rgvarg + index) = pv;
	V_VT(rgvarg + index) = VT_VARIANT|VT_BYREF;
}
