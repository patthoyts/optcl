

/*
 *-------------------------------------------------------------------------
 * ComRecordInfoImpl.h
 *	Declares a IRecordInfo, that unlike the one shipped by MS, isn't 
 *	reliant on the presence of a GUID for any structure.
 *
 *	Copyright (C) 2000 Farzad Pezeshkpour
 * Email:	fuzz@sys.uea.ac.uk
 * Date:	6th April 2000
 *
 * How-To:	1) Add both this file and ComRecordInfoImpl.cpp to your project
 *			2) Include this file where-ever you wish to access a structure
 *			   using IRecordInfo.
 *			3) Call GetRecordInfoFromTypeInfo2 instead of 
 *			   GetRecordInfoFromTypeInfo to retrieve an IRecordInfo.
 * Licence:
 *	This library is free software; you can redistribute it and/or
 *	modify it under the terms of the GNU Lesser General Public
 *	License as published by the Free Software Foundation; either
 *	version 2.1 of the License, or (at your option) any later version.
 *
 *	This library is distributed in the hope that it will be useful,
 *	but WITHOUT ANY WARRANTY; without even the implied warranty of
 *	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 *	Lesser General Public License for more details.
 *
 *	You should have received a copy of the GNU Lesser General Public
 *	License along with this library; if not, write to the Free Software
 *	Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
 *
 *-------------------------------------------------------------------------
 */




#if !defined(AFX_COMRECORDINFOIMPL_H__B3BDEDA0_FB84_11D3_9D8A_DFFCB467E034__INCLUDED_)
#define AFX_COMRECORDINFOIMPL_H__B3BDEDA0_FB84_11D3_9D8A_DFFCB467E034__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/*
 *-------------------------------------------------------------------------
 * GetRecordInfoFromTypeInfo2 --
 *	This is a replacement for GetRecordInfoFromTypeInfo. It returns an 
 *	instance of the new IRecordInfo.
 *
 * Result:
 *	Standard COM result. S_OK iff all ok.
 *
 * Side Effects:
 *	Memory allocated for the new object implementing IRecordInfo.
 *-------------------------------------------------------------------------
 */
HRESULT GetRecordInfoFromTypeInfo2 (ITypeInfo *pti, IRecordInfo **ppri);

#endif // !defined(AFX_COMRECORDINFOIMPL_H__B3BDEDA0_FB84_11D3_9D8A_DFFCB467E034__INCLUDED_)
