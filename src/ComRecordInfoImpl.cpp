

/*
 *-------------------------------------------------------------------------
 * ComRecordInfoImpl.cpp
 *	Implements an IRecordInfo, that unlike the one shipped by MS, isn't 
 *	reliant on the presence of a GUID for any structure.
 * Copyright (C) 2000 Farzad Pezeshkpour
 *
 * Email:	fuzz@sys.uea.ac.uk
 * Date:	6th April 2000
 *
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
 *-------------------------------------------------------------------------
 */



#include "stdafx.h"
#include <atlbase.h>
#include <atlcom.h>
#include "ComRecordInfoImpl.h"

/*
 *-------------------------------------------------------------------------
 * Class CComRecordInfoImpl --
 *	Declaration of the class that implements the new IRecord Info.
 *-------------------------------------------------------------------------
 */
class CComRecordInfoImpl : 
public CComObjectRoot, public IRecordInfo
{
public:
	BEGIN_COM_MAP(CComRecordInfoImpl)
		COM_INTERFACE_ENTRY(IRecordInfo)
	END_COM_MAP()


	CComRecordInfoImpl();
	virtual ~CComRecordInfoImpl();

	HRESULT SetTypeInfo (ITypeInfo *pti);
	void FinalRelease ();

	STDMETHOD(RecordInit)(PVOID pvNew);
    STDMETHOD(RecordClear)(PVOID pvExisting);
    STDMETHOD(RecordCopy)(PVOID pvExisting, PVOID pvNew);
    STDMETHOD(GetGuid)(GUID  *pguid);
	STDMETHOD(GetName)(BSTR  *pbstrName);
    STDMETHOD(GetSize)(ULONG  *pcbSize);
    STDMETHOD(GetTypeInfo)(ITypeInfo  * *ppTypeInfo);
    STDMETHOD(GetField)(PVOID pvData, LPCOLESTR szFieldName, VARIANT  *pvarField);
    STDMETHOD(GetFieldNoCopy)(PVOID pvData, LPCOLESTR szFieldName, VARIANT  *pvarField, PVOID  *ppvDataCArray);
    STDMETHOD(PutField)(ULONG wFlags, PVOID pvData, LPCOLESTR szFieldName, VARIANT  *pvarField);
    STDMETHOD(PutFieldNoCopy)(ULONG wFlags, PVOID pvData, LPCOLESTR szFieldName, VARIANT  *pvarField);
    STDMETHOD(GetFieldNames)(ULONG  *pcNames, BSTR  *rgBstrNames);
    BOOL STDMETHODCALLTYPE IsMatchingType(IRecordInfo  *pRecordInfo);
    PVOID STDMETHODCALLTYPE RecordCreate(void);
    STDMETHOD(RecordCreateCopy)(PVOID pvSource, PVOID  *ppvDest);
    STDMETHOD(RecordDestroy)(PVOID pvRecord);

protected:
	STDMETHODIMP GetFieldNoCopy(PVOID pvData, VARDESC *pvd, VARIANT  *pvarField, PVOID  *ppvDataCArray);
	STDMETHODIMP PutFieldNoCopy(ULONG wFlags, PVOID pvData, VARDESC *pvd, VARIANT  *pvarField);
protected:
	void	ReleaseTypeAttr ();

protected:
	CComPtr<ITypeInfo>	m_pti; // type info we're implementing
	TYPEATTR *			m_pta; // type attribute for the type
	CComBSTR			m_name; // name of the this record type 
};



/*
 *-------------------------------------------------------------------------
 * Class: CVarDesc
 *	Implements a wrapper for the VARDESC data type, and its retrieval from
 *	an ITypeInfo interface pointer.
 *-------------------------------------------------------------------------
 */
class CVarDesc {
protected:
	CComPtr<ITypeInfo> m_pti;	// reference to the ITypeInfo parent of the VARDESC
public:
	VARDESC *	m_pvd;			// the vardesc itself
public:
	// constructor / destructor
	CVarDesc () : m_pvd(NULL) {}

	virtual ~CVarDesc () {
		Release();
	}


	// operator overloads to make this object look more like a VARDESC...
	
	// pointer de-reference
	VARDESC * operator-> () {
		ATLASSERT (m_pvd != NULL);
		return m_pvd;
	}

	// castin operator
	operator VARDESC* () {
		ATLASSERT (m_pvd != NULL);
		return m_pvd;		
	}

	/*
	 *-------------------------------------------------------------------------
	 * Release --
	 *	Releases the VARDESC if it has been allocated.
	 *	Releases reference to the ITypeInfo.
	 *
	 * Result:
	 *	None.
	 *
	 * Side Effects:
	 *	None.	
	 *-------------------------------------------------------------------------
	 */
	void Release () {
		if (m_pvd != NULL) {
			ATLASSERT(m_pti != NULL);
			m_pti->ReleaseVarDesc(m_pvd);
			m_pti.Release();
			m_pvd = NULL;
		}
	}


	/*
	 *-------------------------------------------------------------------------
	 * Set --
	 *	Sets the VARDESC based on an index into the ITypeInfo parameter.
	 *
	 * Result:
	 *	S_OK iff succeeded.
	 *
	 * Side Effects:
	 *	Any previous VARDESC is released.	
	 *-------------------------------------------------------------------------
	 */
	HRESULT Set (ITypeInfo *pti, ULONG index) {
		Release();
		m_pti = pti;
		HRESULT hr;
		hr = m_pti->GetVarDesc (index, &m_pvd);
		return hr;
	}


	/*
	 *-------------------------------------------------------------------------
	 * Set --
	 *	Sets the VARDESC based on the variable name within the ITypeInfo parameter.
	 *	
	 * Result:
	 *	S_OK iff succeeded.
	 *
	 * Side Effects:
	 *	Any previous VARDESC is released.	
	 *-------------------------------------------------------------------------
	 */
	HRESULT Set (ITypeInfo *pti, LPCOLESTR name) {
		CComPtr<ITypeComp> ptc;
		HRESULT hr;
		hr = pti->GetTypeComp (&ptc);
		if (FAILED(hr))
			return hr;
		CComPtr<ITypeInfo> pti2;
		DESCKIND dk;
		BINDPTR bp;
		hr = ptc->Bind ((LPOLESTR)name, 0, INVOKE_PROPERTYPUT | INVOKE_PROPERTYPUTREF, &pti2, &dk, &bp);
		if (FAILED(hr))
			return hr;
		if (dk != DESCKIND_VARDESC) {
			ReleaseBindPtr(dk, bp);
			return E_FAIL;
		} else {
			Release();
			m_pvd = bp.lpvardesc;
			m_pti = pti;
			return S_OK;
		}
	}


private:
	/*
	 *-------------------------------------------------------------------------
	 * ReleaseBindPtr --
	 *	Releases the bind ptr according to its type.
	 *
	 * Result:
	 *	None.
	 *
	 * Side Effects:
	 *	None.	
	 *-------------------------------------------------------------------------
	 */
	void ReleaseBindPtr (DESCKIND dk, BINDPTR bp) {
		if (bp.lptcomp == NULL)
			return;

		switch (dk) {
		case DESCKIND_FUNCDESC:
			m_pti->ReleaseFuncDesc(bp.lpfuncdesc);
			break;			
		case DESCKIND_TYPECOMP:
			bp.lptcomp->Release();
			break;
		default:
			ATLASSERT(FALSE);
			break;
		}
	}
};


//------------------ IRecordInfo Implementation ---------------------------


/*
 *-------------------------------------------------------------------------
 * GetRecordInfoFromTypeInfo2 --
 *	Creates a valid IRecordInfo interface for the give ITypeInfo interface.
 *	The only criteria is that the type info must be of the type TKIND_RECORD
 *	The type info does not have to provide a GUID.
 *
 * Result:
 *	S_OK iff successful.
 *
 * Side Effects:
 *	A CComRecordInfo object is created on the heap.
 *
 *-------------------------------------------------------------------------
 */
HRESULT GetRecordInfoFromTypeInfo2 (ITypeInfo *pti, IRecordInfo **ppri)
{
	ATLASSERT (pti != NULL && ppri != NULL);
	CComObject<CComRecordInfoImpl> *pri = NULL;
	CComPtr<IRecordInfo> ptmpri;
	HRESULT hr = CComObject<CComRecordInfoImpl>::CreateInstance (&pri);
	if (FAILED (hr))
		return hr;
	hr = pri->QueryInterface (&ptmpri);
	if (FAILED(hr))
		return hr;
	hr = pri->SetTypeInfo (pti);
	if (FAILED (hr))
		return hr;
	return ptmpri.CopyTo(ppri);
}





//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CComRecordInfoImpl::CComRecordInfoImpl() : m_pta(NULL)
{

}

CComRecordInfoImpl::~CComRecordInfoImpl()
{

}

/*
 *-------------------------------------------------------------------------
 * FinalRelease --
 *	Called by the ATL framework when the object is about to be destroyed.
 *
 * Result:
 *	None.
 *
 * Side Effects:
 *	Releases the TYPEATTR for this Record Info.	
 *-------------------------------------------------------------------------
 */
void CComRecordInfoImpl::FinalRelease () {
	ReleaseTypeAttr();
}


/*
 *-------------------------------------------------------------------------
 * SetTypeInfo --
 *	Sets the current TypeInfo that this RecordInfo is implementing.
 *
 * Result:
 *	S_OK iff successful.
 *
 * Side Effects:
 *	Releases any previous type info reference and attributes.
 *
 *-------------------------------------------------------------------------
 */
HRESULT CComRecordInfoImpl::SetTypeInfo (ITypeInfo *pti)
{
	TYPEATTR *pta = NULL;
	// retrieve the type attribute for the 
	try {
		if (FAILED(pti->GetTypeAttr(&pta)))
			throw false;
		if (pta->typekind != TKIND_RECORD)
			throw false;
		ReleaseTypeAttr();
		m_pti = pti;
		m_pta = pta;
		pti->GetDocumentation(-1, &m_name, NULL, NULL, NULL);
		return S_OK;
	} catch (...) {
		if (pta != NULL)
			pti->ReleaseTypeAttr(pta);
		return E_INVALIDARG;
	}
}

/*
 *-------------------------------------------------------------------------
 * ReleaseTypeAttr --
 *	Releases the TYPEATTR if any.
 *
 * Result:
 *	None.
 *
 * Side Effects:
 *	None.	
 *-------------------------------------------------------------------------
 */
void CComRecordInfoImpl::ReleaseTypeAttr ()
{
	ATLASSERT (m_pta == NULL || m_pti != NULL);

	if (m_pta != NULL && m_pti != NULL) {
		m_pti->ReleaseTypeAttr(m_pta);
		m_pta = NULL;
	}
}






/*
 *-------------------------------------------------------------------------
 * RecordInit --
 *	Initiliases the contents of a created record structure. All existing
 *	values are ignored.
 *
 * Result:
 *	S_OK iff successfull.
 *
 * Side Effects:
 *	None.
 *
 *-------------------------------------------------------------------------
 */
STDMETHODIMP CComRecordInfoImpl::RecordInit(PVOID pvNew)
{
	HRESULT hr;
	for (WORD iVar = 0; iVar < m_pta->cVars; iVar++) {

		CVarDesc vd;
		PVOID pvField;
		CComPtr<ITypeInfo> pRefInfo;
		CComPtr<IRecordInfo> pRefRecInfo;


		hr = vd.Set(m_pti, iVar);
		if (FAILED(hr))
			return hr;
		ATLASSERT ( (vd->elemdescVar.tdesc.vt & VT_BYREF) == NULL);


		pvField = (BYTE*)pvNew + vd->oInst;
		
		
		switch (vd->elemdescVar.tdesc.vt) {
		case VT_USERDEFINED:
			hr = m_pti->GetRefTypeInfo(vd->elemdescVar.tdesc.hreftype, &pRefInfo);
			if (FAILED(hr)) return hr;

			hr = GetRecordInfoFromTypeInfo2 (pRefInfo, &pRefRecInfo);
			if (FAILED(hr)) return hr;

			hr = pRefRecInfo->RecordInit(pvField);
			if (FAILED(hr))
				return hr;
			break;

		case VT_BSTR:
			// is this correct?
			*((BSTR*)pvField) = SysAllocString (L"");
			break;

		case VT_DATE:
			*((DATE*)pvField) = 0;
			break;

		case VT_CY:
			((CY*)pvField)->int64 = 0;
			break;

		// generic 8bit data types 
		case VT_I1:
		case VT_UI1:
			*((BYTE*)pvField) = 0;
			break;

		// generic 16bit data types 
		case VT_I2:
		case VT_UI2:
			*((SHORT*)pvField) = 0;
			break;

		// generic 32 bit data types
		case VT_I4:
		case VT_UI4:
		case VT_R4:
		case VT_UNKNOWN:
		case VT_DISPATCH:
		case VT_ERROR:
			*((ULONG*)pvField) = 0;
			break;

		// platform specific: INT
		case VT_INT:
		case VT_UINT:
			*((INT*)pvField) = 0;
			break;

		// boolean
		case VT_BOOL:
			*((VARIANT_BOOL*)pvField) = VARIANT_FALSE;
			break;
		
		// double
		case VT_R8:
			*((DOUBLE*)pvField) = double(0);
			break;

		default:
			// is it an array?
			if (vd->elemdescVar.tdesc.vt & VT_ARRAY) {
				*((SAFEARRAY**)pvField) = NULL;
			}
		}
		
	}
	return S_OK;
}




/*
 *-------------------------------------------------------------------------
 * CComRecordInfoImpl::RecordClear --
 *	Iterates through the existing record, clearing all referenced resources, 
 *	and setting to zero.
 *
 * Result:
 *	Standard COM result.
 *
 * Side Effects:
 *	None.	
 *-------------------------------------------------------------------------
 */
STDMETHODIMP CComRecordInfoImpl::RecordClear(PVOID pvExisting)
{
	HRESULT hr;
	for (WORD iVar = 0; iVar < m_pta->cVars; iVar++) {

		CVarDesc vd;
		PVOID pvField;
		CComPtr<ITypeInfo> pRefInfo;
		CComPtr<IRecordInfo> pRefRecInfo;


		hr = vd.Set(m_pti, iVar);
		if (FAILED(hr))
			return hr;
		ATLASSERT ( (vd->elemdescVar.tdesc.vt & VT_BYREF) == NULL);

		pvField = (BYTE*)pvExisting + vd->oInst;
		
		if (vd->elemdescVar.tdesc.vt & VT_ARRAY) {
			SafeArrayDestroy (*((SAFEARRAY**)pvField));
			*((SAFEARRAY**)pvField) = NULL;
		} else {
			switch (vd->elemdescVar.tdesc.vt) {
			case VT_USERDEFINED:
				hr = m_pti->GetRefTypeInfo(vd->elemdescVar.tdesc.hreftype, &pRefInfo);
				if (FAILED(hr)) return hr;

				hr = GetRecordInfoFromTypeInfo2 (pRefInfo, &pRefRecInfo);
				if (FAILED(hr)) return hr;

				hr = pRefRecInfo->RecordClear(pvField);
				if (FAILED(hr))
					return hr;
				break;
			/* strings */
			case VT_BSTR:
				SysFreeString(*( (BSTR*)pvField ));
				*( (BSTR*)pvField ) = NULL;
				break;
			/* interface types */
			case VT_DISPATCH:
			case VT_UNKNOWN:
				(*((IUnknown**)pvField))->Release();
				(*((IUnknown**)pvField)) = NULL;
				break;
			}
		}
	}
	return S_OK;
}


/*
 *-------------------------------------------------------------------------
 * CComRecordInfoImpl::RecordCopy --
 *	Makes a copy of the existing record to the new record.
 *
 * Result:
 *	Standard COM result.
 *
 * Side Effects:
 *	Performs a deep copy on all references.	
 *-------------------------------------------------------------------------
 */
STDMETHODIMP CComRecordInfoImpl::RecordCopy(PVOID pvExisting, PVOID pvNew)
{
	HRESULT hr;
	for (WORD iVar = 0; iVar < m_pta->cVars; iVar++) {
		PVOID pvSrc, pvDst;
		CVarDesc vd;
		CComPtr<ITypeInfo> refInfo;
		CComPtr<IRecordInfo> refrecInfo;

		hr = vd.Set (m_pti, iVar);
		if (FAILED(hr)) return hr;

		pvSrc = (BYTE*)pvExisting + vd->oInst;
		pvDst = (BYTE*)pvNew + vd->oInst;

		ATLASSERT ( (vd->elemdescVar.tdesc.vt & VT_BYREF) == 0);
		if (vd->elemdescVar.tdesc.vt & VT_ARRAY != 0) {
			hr = SafeArrayCopyData (*((SAFEARRAY**)pvSrc), *((SAFEARRAY**)pvDst));
			if (FAILED(hr)) return hr;
		} else {
			switch (vd->elemdescVar.tdesc.vt) {
			// interfaces ...
			case VT_UNKNOWN:
			case VT_DISPATCH:
				*((IUnknown**)pvDst) = *((IUnknown**)pvSrc);
				(*((IUnknown**)pvDst))->AddRef();
				break;
			// string
			case VT_BSTR:
				*((BSTR*)pvDst) = SysAllocString (*((BSTR*)pvSrc));
				break;
			// 8 bit copy
			case VT_I1:
			case VT_UI1:
				*((BYTE*)pvDst) = *((BYTE*)pvSrc);
				break;
			// 16 bit copy
			case VT_I2:
			case VT_UI2:
				*((SHORT*)pvDst) = *((SHORT*)pvSrc);
				break;
			// 32 bit copy
			case VT_I4:
			case VT_UI4:
			case VT_R4:
			case VT_ERROR:
				*((ULONG*)pvDst) = *((ULONG*)pvSrc);
				break;
			// doubles (64 bit)
			case VT_R8:
				*((DOUBLE*)pvDst) = *((DOUBLE*)pvSrc);
				break;
			// currency
			case VT_CY:
				*((CY*)pvDst) = *((CY*)pvSrc);
				break;
			// date
			case VT_DATE:
				*((DATE*)pvDst) = *((DATE*)pvSrc);
				break;
			// boolean
			case VT_BOOL:
				*((VARIANT_BOOL*)pvDst) = *((VARIANT_BOOL*)pvSrc);
				break;
			// decimal
			case VT_DECIMAL:
				*((DECIMAL*)pvDst) = *((DECIMAL*)pvSrc);
				break;
			// TypeLib defined
			case VT_USERDEFINED:
				hr = m_pti->GetRefTypeInfo(vd->elemdescVar.tdesc.hreftype, &refInfo);
				if (FAILED(hr)) return hr;
				hr = GetRecordInfoFromTypeInfo2 (m_pti, &refrecInfo);
				if (FAILED(hr)) return hr;
				hr = refrecInfo->RecordCopy (pvSrc, pvDst);
				if (FAILED(hr)) return hr;
				break;
			default:
				break;
			}
		}
	}
	return S_OK;
}


/*
 *-------------------------------------------------------------------------
 * CComRecordInfoImpl::GetGuid --
 *	Retrieve GUID of struct. Can possibly be IID_NULL.
 *
 * Result:
 *	S_OK
 *
 * Side Effects:
 *	None.
 *-------------------------------------------------------------------------
 */
STDMETHODIMP CComRecordInfoImpl::GetGuid(GUID  *pguid)
{
	*pguid = m_pta->guid;
	return S_OK;
}


/*
 *-------------------------------------------------------------------------
 * CComRecordInfoImpl::GetName --
 *	Retrieve the name of the structure.
 *
 * Result:
 *	S_OK;
 *
 * Side Effects:
 *	None.	
 *-------------------------------------------------------------------------
 */
STDMETHODIMP CComRecordInfoImpl::GetName(BSTR  *pbstrName)
{
	*pbstrName = m_name.Copy();
	return (pbstrName!=NULL?S_OK:E_FAIL);
}


/*
 *-------------------------------------------------------------------------
 * CComRecordInfoImpl::GetSize --
 *	Retrieve the size, in bytes of the structure.
 *
 * Result:
 *	None.
 *
 * Side Effects:
 *	None.	
 *-------------------------------------------------------------------------
 */
STDMETHODIMP CComRecordInfoImpl::GetSize(ULONG  *pcbSize)
{
	ATLASSERT (m_pta != NULL);
	*pcbSize = m_pta->cbSizeInstance;
	return S_OK;
}


/*
 *-------------------------------------------------------------------------
 * CComRecordInfoImpl::GetTypeInfo --
 *	Retrieve ITypeInfo for this structure.
 *
 * Result:
 *	S_OK iff all ok.
 *
 * Side Effects:
 *	None.	
 *-------------------------------------------------------------------------
 */
STDMETHODIMP CComRecordInfoImpl::GetTypeInfo(ITypeInfo **ppTypeInfo)
{
	ATLASSERT(m_pti != NULL);
	return m_pti.CopyTo(ppTypeInfo);
}


/*
 *-------------------------------------------------------------------------
 * CComRecordInfoImpl::GetField --
 *	Retrieve the value of a given field within a structure of this type
 *	The value of the field is returned as a copy of the original.
 * Result:
 *	
 * Side Effects:
 *	
 *-------------------------------------------------------------------------
 */
STDMETHODIMP CComRecordInfoImpl::GetField(PVOID pvData, LPCOLESTR szFieldName, VARIANT  *pvarField)
{
	VARIANT refVar;
	PVOID pvFieldData;
	HRESULT hr;

	VariantInit (&refVar);
	VariantClear(pvarField);

	hr = GetFieldNoCopy (pvData, szFieldName, &refVar, &pvFieldData);
	if (FAILED(hr))
		return hr;
	hr = VariantCopyInd(pvarField, &refVar);
	return hr;
}


/*
 *-------------------------------------------------------------------------
 * CComRecordInfoImpl::GetFieldNoCopy --
 *	Retrieve a direct reference to the field's value using a VARDESC to identify the
 *	field. The caller must not free the returned variant.
 *
 * Result:
 *	S_OK iff ok.
 *
 * Side Effects:
 *	None.	
 *-------------------------------------------------------------------------
 */
STDMETHODIMP CComRecordInfoImpl::GetFieldNoCopy(PVOID pvData, VARDESC *pvd, VARIANT  *pvarField, PVOID  *ppvDataCArray) {
	HRESULT hr;
	hr = VariantClear (pvarField);
	if (FAILED(hr)) return hr;

	// retrieve a pointer to the field data
	PVOID pfield;
	pfield = ( ((BYTE*)pvData) + pvd->oInst);
	*ppvDataCArray = pfield;

	// now crack the field type ...

	// first some assertions ...
	// not by-reference (COM Automation / Variant Declaration rules)
	ATLASSERT ( (pvd->elemdescVar.tdesc.vt & VT_BYREF) == 0);

	if (pvd->elemdescVar.tdesc.vt == VT_USERDEFINED) {
		// resolve the referenced type
		CComPtr<ITypeInfo> pRefInfo;
		CComPtr<IRecordInfo> pRefRecInfo;
		hr = m_pti->GetRefTypeInfo (pvd->elemdescVar.tdesc.hreftype, &pRefInfo);
		if (FAILED(hr))
			return hr;
		hr = GetRecordInfoFromTypeInfo2 (pRefInfo, &pRefRecInfo);
		if (FAILED(hr))
			return hr;

		// set the field reference and its record info
		pvarField->pvRecord = pfield;
		hr = pRefRecInfo.CopyTo(&(pvarField->pRecInfo));
		if (FAILED(hr))
			return hr;
		pvarField->vt = VT_RECORD;
	} else {
		// in all other cases, we just set the pointer to the field member
		pvarField->byref = pfield;
		// the vartype of the resulting parameter will be a reference to the type of the field
		pvarField->vt = (pvd->elemdescVar.tdesc.vt | VT_BYREF);
	}
	return S_OK;

}


/*
 *-------------------------------------------------------------------------
 * CComRecordInfoImpl::GetFieldNoCopy --
 *	Retrieve the value of a field as a reference, given the name of the field.
 *
 * Result:
 *	S_OK iff ok.
 *
 * Side Effects:
 *	None.	
 *-------------------------------------------------------------------------
 */
STDMETHODIMP CComRecordInfoImpl::GetFieldNoCopy(PVOID pvData, LPCOLESTR szFieldName, VARIANT  *pvarField, PVOID  *ppvDataCArray)
{
	HRESULT hr;
	CVarDesc vd;

	hr = vd.Set(m_pti, szFieldName);
	if (FAILED(hr)) return hr;

	hr = VariantClear (pvarField);
	if (FAILED(hr)) return hr;
	return GetFieldNoCopy (pvData, vd, pvarField, ppvDataCArray);
}


/*
 *-------------------------------------------------------------------------
 * CComRecordInfoImpl::PutField --
 *	Places a copy of the variant to the field, applying any type coercion
 *	as required. Rules for INVOKE_PROPERTYPUT are handled at a deeper 
 *	level of call.
 *
 * Result:
 *	S_OK iff all ok.
 *
 * Side Effects:
 *	None.	
 *-------------------------------------------------------------------------
 */
STDMETHODIMP CComRecordInfoImpl::PutField(ULONG wFlags, PVOID pvData, LPCOLESTR szFieldName, VARIANT  *pvarField)
{
	CVarDesc vd;
	HRESULT hr;
	hr = vd.Set(m_pti, szFieldName);
	if (FAILED(hr)) return hr;

	VARIANT varCopy;
	VariantInit (&varCopy);
	hr = VariantCopy (&varCopy, pvarField);
	if (FAILED(hr)) return hr;
	return PutFieldNoCopy (wFlags, pvData, vd, &varCopy);
}

/*
 *-------------------------------------------------------------------------
 * CComRecordInfoImpl::PutFieldNoCopy --
 *	Given the VARDESC for a field, this function places the value in 
 *	pvarField to the field, without allocating any new resources.
 *	I'm not too sure about the INVOKE_PROPERTYPUT implementation
 *	which I've tried to follow from the MSDN documentation. As
 *	far as I can make out, the field must be of type VT_DISPATCH
 *	(or do I have to explicitly check for derivation from IDispatch?)
 *	The value is either of type VT_DISPATCH (in which case it's default
 *	property is used as the actual value), or any other valid variant
 *	sub-type. The actual value will be set to the default property of
 *	the field.
 *	
 * Result:
 *	Standard COM result - S_OK iff all OK.
 *
 * Side Effects:
 *	None.	
 *-------------------------------------------------------------------------
 */
STDMETHODIMP CComRecordInfoImpl::PutFieldNoCopy(ULONG wFlags, PVOID pvData, VARDESC *pvd, VARIANT  *pvarField)
{
	PVOID field = (BYTE*)pvData + pvd->oInst;
	HRESULT hr;

	// perform the conversion ...
	
	if (wFlags == INVOKE_PROPERTYPUT) {

		// if the field isn't a dispatch object or is null then we fail
		if (pvd->elemdescVar.tdesc.vt != VT_DISPATCH)
			return E_FAIL;

		IDispatch * pdisp = *((IDispatch**)field);
		if (pdisp == NULL)
			return E_FAIL;

		CComVariant varResult;
		DISPPARAMS dp;
		DISPID dispidNamed = DISPID_PROPERTYPUT;
		dp.cArgs = 1;
		dp.cNamedArgs = 1;
		dp.rgdispidNamedArgs = &dispidNamed;
		dp.rgvarg = pvarField;
		hr = pdisp->Invoke (DISPID_VALUE, IID_NULL, 0, DISPID_PROPERTYPUT, &dp, &varResult, NULL, NULL);
		return hr;
	} else {
		// do a straight conversion
		hr = VariantChangeType (pvarField, pvarField, NULL, pvd->elemdescVar.tdesc.vt);
		if (FAILED(hr))
			return hr;

		// now perform a shallow copy
		if (pvd->elemdescVar.tdesc.vt & VT_ARRAY != 0) {
			*((SAFEARRAY**)field) = pvarField->parray;
		} else {
			switch (pvd->elemdescVar.tdesc.vt) {
			// interfaces ...
			case VT_UNKNOWN:
			case VT_DISPATCH:
				*((IUnknown**)field) = pvarField->punkVal;
				break;
			// string
			case VT_BSTR:
				*((BSTR*)field) = pvarField->bstrVal;
				break;
			// 8 bit copy
			case VT_I1:
			case VT_UI1:
				*((BYTE*)field) = pvarField->bVal;
				break;
			// 16 bit copy
			case VT_I2:
			case VT_UI2:
				*((SHORT*)field) = pvarField->iVal;
				break;
			// 32 bit copy
			case VT_I4:
			case VT_UI4:
			case VT_R4:
			case VT_ERROR:
				*((ULONG*)field) = pvarField->ulVal;
				break;
			// doubles (64 bit)
			case VT_R8:
				*((DOUBLE*)field) = pvarField->dblVal;
				break;
			// currency
			case VT_CY:
				*((CY*)field) = pvarField->cyVal;
				break;
			// date
			case VT_DATE:
				*((DATE*)field) = pvarField->date;
				break;
			// boolean
			case VT_BOOL:
				*((VARIANT_BOOL*)field) = pvarField->boolVal;
				break;
			// decimal
			case VT_DECIMAL:
				*((DECIMAL*)field) = pvarField->decVal;
				break;
			// TypeLib defined
			case VT_USERDEFINED:
				*((PVOID*)field) = pvarField->pvRecord;
				break;
			default:
				break;
			}
		}
		return S_OK;
	}
	
}


/*
 *-------------------------------------------------------------------------
 * CComRecordInfoImpl::PutFieldNoCopy --
 *	As the VARDESC variation above, but using the field name instead.
 *
 * Result:
 *	S_O iff all ok.
 *
 * Side Effects:
 *	None.	
 *-------------------------------------------------------------------------
 */
STDMETHODIMP CComRecordInfoImpl::PutFieldNoCopy(ULONG wFlags, PVOID pvData, LPCOLESTR szFieldName, VARIANT  *pvarField)
{
	CVarDesc vd;
	HRESULT hr;
	hr = vd.Set(m_pti, szFieldName);
	if (FAILED(hr)) return hr;
	return PutFieldNoCopy (wFlags, pvData, vd, pvarField);
}


/*
 *-------------------------------------------------------------------------
 * CComRecordInfoImpl::GetFieldNames --
 *	Retrieves an array of fields names.
 *
 * Result:
 *	S_OK iff all ok.
 *
 * Side Effects:
 *	None.	
 *-------------------------------------------------------------------------
 */
STDMETHODIMP CComRecordInfoImpl::GetFieldNames(ULONG  *pcNames, BSTR  *rgBstrNames)
{
	ULONG index = 0;
	if (pcNames == NULL)
		return E_INVALIDARG;
	if (rgBstrNames == NULL) {
		*pcNames = m_pta->cVars;
		return S_OK;
	}

	if (*pcNames > m_pta->cVars)
		*pcNames = m_pta->cVars;

	try {
		for (index = 0; index < *pcNames; index++) {
			CVarDesc vd;
			HRESULT hr;
			hr = vd.Set (m_pti, index);
			if (FAILED(hr))
				throw (hr);
			
			UINT dummy = 1;
			hr = m_pti->GetNames (vd->memid, rgBstrNames+index, 1, &dummy);
			if (FAILED(hr))
				throw(hr);
		}
	} catch (HRESULT hr) {
		while (index > 0) 
			SysFreeString (rgBstrNames[--index]);
		return hr;
	}
	return S_OK;
}


/*
 *-------------------------------------------------------------------------
 * CComRecordInfoImpl::IsMatchingType --
 *	Checks for equivalence of this record type and the one referenced by
 *	the only parameter. Because we can't guarantee the use of GUIDs
 *	I've settled for matching on the type and library name.
 *
 * Result:
 *	TRUE iff the record structures match.
 *
 * Side Effects:
 *	None.	
 *-------------------------------------------------------------------------
 */
BOOL STDMETHODCALLTYPE CComRecordInfoImpl::IsMatchingType(IRecordInfo  *pRecordInfo)
{
	BOOL result = FALSE;
	CComBSTR bstrOtherName;
	HRESULT hr;

	hr = pRecordInfo->GetName(&bstrOtherName);
	if (FAILED(hr)) return FALSE;

	if (wcscmp(bstrOtherName, m_name) == 0) {
		CComPtr<ITypeInfo> pOtherInfo;
		CComPtr<ITypeLib> pOurLib, pOtherLib;
		UINT dummy;
		TLIBATTR * pOurAttr = NULL, *pOtherAttr = NULL;

		hr = pRecordInfo->GetTypeInfo(&pOtherInfo);
		if (FAILED (hr)) return FALSE;

		hr = pOtherInfo->GetContainingTypeLib(&pOtherLib, &dummy);
		if (FAILED(hr)) return FALSE;

		hr = m_pti->GetContainingTypeLib(&pOurLib, &dummy);
		if (FAILED(hr)) return FALSE;

		hr = pOurLib->GetLibAttr (&pOurAttr);
		hr = pOtherLib->GetLibAttr (&pOtherAttr);
		if (pOurAttr != NULL && pOtherAttr != NULL)
			result = (pOurAttr->guid == pOtherAttr->guid);
		if (pOurAttr != NULL)
			pOtherLib->ReleaseTLibAttr (pOurAttr);
		if (pOtherAttr != NULL)
			pOtherLib->ReleaseTLibAttr (pOtherAttr);
	}
	return result;
}


/*
 *-------------------------------------------------------------------------
 * STDMETHODCALLTYPE CComRecordInfoImpl::RecordCreate --
 *	Allocates (using the task memory allocator) a new record, and
 *	initialises it.
 *
 * Result:
 *	Pointer to the record structure iff successfull; else NULL.
 *
 * Side Effects:
 *	None.	
 *-------------------------------------------------------------------------
 */
PVOID STDMETHODCALLTYPE CComRecordInfoImpl::RecordCreate( void)
{
	PVOID prec = CoTaskMemAlloc(m_pta->cbSizeInstance);
	if (FAILED(RecordInit(prec))) {
		CoTaskMemFree(prec);
		prec = NULL;
	}
	return prec;
}


/*
 *-------------------------------------------------------------------------
 * CComRecordInfoImpl::RecordCreateCopy --
 *	Creates a copy of the passed record structure.
 *
 * Result:
 *	S_OK iff successfull.
 *
 * Side Effects:
 *	None.	
 *-------------------------------------------------------------------------
 */
STDMETHODIMP CComRecordInfoImpl::RecordCreateCopy(PVOID pvSource, PVOID  *ppvDest)
{
	*ppvDest = RecordCreate();
	if (*ppvDest == NULL)
		return E_FAIL;
	return RecordCopy (pvSource, *ppvDest);
}


/*
 *-------------------------------------------------------------------------
 * CComRecordInfoImpl::RecordDestroy --
 *	Clears the given record and releases the memory associated with it.
 *
 * Result:
 *	S_OK iff all OK.
 *
 * Side Effects:
 *	None.	
 *-------------------------------------------------------------------------
 */
STDMETHODIMP CComRecordInfoImpl::RecordDestroy(PVOID pvRecord)
{
	HRESULT hr;
	if (pvRecord) {
		hr = RecordClear(pvRecord);
		CoTaskMemFree(pvRecord);
	}
	return hr;
}

