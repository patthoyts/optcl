/*
 *------------------------------------------------------------------------------
 *	tbase.h
 *	C++ Wrapper classes for common Tcl types.
 *
 * Updated: 1999.03.08 - Removed a few bugs from TObjPtr
 * Updated: 1999.07.11 - Added isnull and isnotnull to TObjPtr
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


#ifndef _3CC705E0_BA28_11d2_8003_0040055861F2_
#define _3CC705E0_BA28_11d2_8003_0040055861F2_


#include <tcl.h>
#include <stdlib.h>
#include <string.h>


#ifndef ASSERT
#	ifdef _DEBUG
#		include <crtdbg.h> 
#		define ASSERT(x) _ASSERTE(x)
#	else
#		define ASSERT(x)
#	endif
#endif




class TObjPtr
{
protected:
	Tcl_Obj *m_po; 
	bool	 m_ba; 
public:
	TObjPtr() : m_po(NULL), m_ba(true)
	{
	}

	TObjPtr (int i, Tcl_Obj *const objs[], bool bauto = true) : m_po (NULL), m_ba(bauto)
	{
		ASSERT (i >= 0 && (i == 0 || objs!=NULL));
		m_po = Tcl_NewListObj (i, objs);
		if (m_ba)
			incr();
	}

	TObjPtr(Tcl_Obj *ptr, bool bauto =true) : m_po(ptr), m_ba(bauto)
	{
		if (m_ba)
			incr();
	}

	TObjPtr(const TObjPtr &src, bool bauto=false) : m_po(NULL), m_ba(bauto)
	{
		copy (src, bauto);
	}

	TObjPtr(const char *string, bool bauto=true) : m_po(NULL), m_ba(bauto)
	{
		m_po = Tcl_NewStringObj ((char*)string, -1);
		if (m_po==NULL)
			throw ("failed to create string object");
		if (m_ba)
			incr();
	}

	TObjPtr(const long l, bool bauto=true) : m_po(NULL), m_ba(bauto)
	{
		m_po = Tcl_NewLongObj (l);
		if (m_po==NULL)
			throw ("failed to create long tcl object");
		if (m_ba)
			incr();
	}


	TObjPtr(const int i, bool bauto=true) : m_po(NULL), m_ba(bauto)
	{
		m_po = Tcl_NewIntObj (i);
		if (m_po==NULL)
			throw ("failed to create int tcl object");
		if (m_ba)
			incr();
	}

	TObjPtr(const bool b, bool bauto=true) : m_po(NULL), m_ba(bauto)
	{
		m_po = Tcl_NewBooleanObj (b);
		if (m_po==NULL)
			throw ("failed to create long tcl object");
		if (m_ba)
			incr();
	}

	TObjPtr(const double d, bool bauto=true) : m_po(NULL), m_ba(bauto)
	{
		m_po = Tcl_NewDoubleObj (d);
		if (m_po==NULL)
			throw ("failed to create double object");
		if (m_ba)
			incr();
	}


	virtual ~TObjPtr()
	{
		if (m_ba!=NULL && m_po != NULL) {
			if (m_po->refCount == 0)
				incr();
			decr();
		}
		m_po = NULL;
	}

	Tcl_Obj* create (bool bauto=true)
	{
		if (m_ba!=NULL && m_po != NULL) {
			if (m_po->refCount == 0)
				incr();
			decr();
		}
		m_ba = bauto;
		m_po = Tcl_NewObj ();
		if (m_ba)
			incr();
		return m_po;
	}

	
	void	incr()
	{
		if (m_po)
			Tcl_IncrRefCount(m_po);
	}

	void	decr()
	{
		if (m_po)
			Tcl_DecrRefCount(m_po);
	}

	bool	isnull ()
	{
		return (m_po == NULL);
	}

	bool	isnotnull()
	{
		return (m_po != NULL);
	}

	void attach (Tcl_Obj *ptr, bool bauto=false)
	{
		if (m_ba)
			decr();
		m_po = ptr;
		m_ba = bauto;
	}

	Tcl_Obj *detach ()
	{
		Tcl_Obj *p = m_po;
		m_po = NULL;
		return p;
	}

	void	copy (const Tcl_Obj *src, bool bauto = true)
	{
		ASSERT (src!=NULL);
		if (m_ba)
			decr();

		m_po = Tcl_DuplicateObj((Tcl_Obj*)src);
		ASSERT (m_po);
		m_ba = bauto;
		if (m_ba)
			incr();
	}

	int	llength (Tcl_Interp *pInterp = NULL)
	{
		ASSERT (m_po!=NULL);
		int length;
		if (TCL_OK != Tcl_ListObjLength (pInterp, m_po, &length)) {
			if (pInterp != NULL)
				throw (Tcl_GetStringResult (pInterp));
			else
				throw ("failed to get length of list");
		}
		return length;
	}


	TObjPtr lindex (int index, Tcl_Interp *pInterp = NULL)
	{
		ASSERT (m_po);
		Tcl_Obj *pObj = NULL;
		if (TCL_OK != Tcl_ListObjIndex (pInterp, m_po, index, &pObj)) {
			if (pInterp != NULL)
				throw (Tcl_GetStringResult (pInterp));
			else
				throw ("failed to get list item");
		}
		return TObjPtr(pObj, false);
	}


	TObjPtr& lappend (TObjPtr &pObj, Tcl_Interp *pInterp = NULL)
	{
		return lappend ((Tcl_Obj*)pObj, pInterp);
	}

	TObjPtr& lappend (Tcl_Obj *pObj, Tcl_Interp *pInterp = NULL)
	{
		ASSERT (pObj!=NULL && m_po!=NULL);
		if (TCL_OK != Tcl_ListObjAppendElement (pInterp, m_po, pObj)) {
			if (pInterp != NULL)
				throw (Tcl_GetStringResult (pInterp));
			else
				throw ("failed to add element to list");
		}
		return *this;
	}

	TObjPtr& lappend(const char *string, Tcl_Interp *pInterp = NULL)
	{
		ASSERT (string!=NULL && m_po != NULL);
		return lappend (TObjPtr(string), pInterp);
	}

	TObjPtr& lappend(const int i, Tcl_Interp *pInterp = NULL)
	{
		ASSERT (m_po != NULL);
		return lappend (TObjPtr(i), pInterp);
	}

	TObjPtr& lappend(const long l, Tcl_Interp *pInterp = NULL)
	{
		ASSERT (m_po != NULL);
		return lappend (TObjPtr (l), pInterp);
	}

	TObjPtr& lappend(const double d, Tcl_Interp *pInterp = NULL)
	{
		ASSERT (m_po != NULL);
		return lappend (TObjPtr (d), pInterp);
	}

	TObjPtr& lappend(const bool b, Tcl_Interp *pInterp = NULL)
	{
		ASSERT (m_po != NULL);
		return lappend (TObjPtr (b), pInterp);
	}


	operator int() 
	{
		ASSERT (m_po != NULL);
		int n;
		if (TCL_OK != Tcl_GetIntFromObj (NULL, m_po, &n)) 
			// perform a cast
			n = (int) double (*this);
		return n;
	}


	operator long() 
	{
		long n;
		ASSERT (m_po != NULL);
		if (TCL_OK != Tcl_GetLongFromObj (NULL, m_po, &n))
			// perform a cast
			n = (long) double (*this);
		return n;
	}

	
	operator bool()
	{
		int b;
		ASSERT (m_po != NULL);
		if (TCL_OK != Tcl_GetBooleanFromObj (NULL, m_po, &b))
			throw ("failed to convert object to bool");
		return (b!=0);
	}

	operator double ()
	{
		double d;
		ASSERT (m_po != NULL);
		if (TCL_OK != Tcl_GetDoubleFromObj (NULL, m_po, &d))
			throw ("failed to convert object to double");
		return d;
	}

	operator char*() const
	{
		if (m_po == NULL) return NULL;
		return Tcl_GetStringFromObj(m_po, NULL);
	}

	operator Tcl_Obj*() const
	{
		return m_po;
	}


	TObjPtr &operator= (Tcl_Obj *ptr)
	{
		attach(ptr, true); // automatically sets reference management
		if (m_po != NULL)
			incr();
		return *this;
	}

	TObjPtr &operator= (const char *string)
	{
		ASSERT(string!=NULL && m_po != NULL);
		Tcl_SetStringObj (m_po, (char*)string, -1);
		return *this;
	}

	TObjPtr &operator= (const long l)
	{
		ASSERT (m_po != NULL);
		Tcl_SetLongObj  (m_po, l);
		return *this;
	}

	TObjPtr &operator= (const int i)
	{
		ASSERT (m_po != NULL);
		Tcl_SetIntObj (m_po, i);
		return *this;
	}

	TObjPtr &operator= (const bool b)
	{
		ASSERT (m_po != NULL);
		Tcl_SetBooleanObj (m_po, b?1:0);
		return *this;
	}

	TObjPtr &operator= (const double d)
	{
		ASSERT (m_po != NULL);
		Tcl_SetDoubleObj (m_po, d);
		return *this;
	}


	bool operator== (Tcl_Obj *ptr)
	{
		return (ptr == m_po);
	}


	TObjPtr &operator+= (const char *string)
	{
		ASSERT (string && m_po);
		Tcl_AppendToObj(m_po, (char*)string, -1);
		return *this;
	}

	TObjPtr& operator+= (Tcl_Obj *pObj)
	{
		ASSERT (m_po != NULL);
		return lappend (pObj);
	}

	TObjPtr &operator+= (TObjPtr &pObj)
	{
		ASSERT (m_po != NULL);
		return lappend (pObj);
	}

	TObjPtr &operator+= (int i)
	{
		ASSERT (m_po != NULL);
		(*this) = int(*this) + i;
		return *this;
	}

	TObjPtr &operator+= (long l)
	{
		ASSERT (m_po != NULL);
		(*this) = long(*this) + l;
		return *this;
	}


	TObjPtr &operator+= (double d)
	{
		ASSERT (m_po != NULL);
		(*this) = double(*this) + d;
		return *this;
	}

	TObjPtr &operator-= (Tcl_Obj *pObj)
	{
		ASSERT (m_po != NULL && pObj != NULL);
		Tcl_Obj ** objv;
		int objc;
		char *sObj = Tcl_GetStringFromObj (pObj, NULL),
			 *sTemp;

		if (sObj == NULL)
			return *this;

		Tcl_ListObjGetElements (NULL, m_po, &objc, &objv);	
		for (int i = 0; i < objc; i++)
		{
			if (objv[i] != NULL) {
				sTemp = Tcl_GetStringFromObj (objv[i], NULL);
				if (sTemp != NULL && strcmp (sObj, sTemp) == 0)
					Tcl_ListObjReplace ( NULL, m_po, i, 1, 0, NULL);
			}
		}
		return *this;
	}

	TObjPtr &operator-= (TObjPtr &obj)
	{
		return operator-=((Tcl_Obj*)obj);
	}

	TObjPtr &operator-= (int i)
	{
		return operator=(int(*this) - i);
	}

	TObjPtr &operator-= (long l)
	{
		return operator=(long(*this) - l);
	}

	TObjPtr &operator-= (double d)
	{
		return operator=(double(*this) - d);
	}

	TObjPtr &operator *= (double d)
	{
		return operator=(double(*this) * d);
	}

	TObjPtr &operator *= (int i)
	{
		return operator=(int(*this) * i);
	}

	TObjPtr &operator *= (long l)
	{
		return operator=(long(*this) * l);
	}


	TObjPtr &operator /= (double d)
	{
		return operator=(double(*this) / d);
	}

	TObjPtr &operator /= (int i)
	{
		return operator=(int(*this) / i);
	}

	TObjPtr &operator /= (long l)
	{
		return operator=(long(*this) / l);
	}

	Tcl_Obj **operator &()
	{
		return &m_po;
	}


	Tcl_Obj *operator ->() const
	{
		return m_po;
	}

	bool operator!= (Tcl_Obj *p)
	{
		return (m_po != p);
	}
	
};







template <class K, class V>
class THashIterator 
{
protected:
	Tcl_HashTable *m_pt;
	Tcl_HashEntry *m_pe;
	Tcl_HashSearch m_s;

public:
	THashIterator () : m_pt(NULL),
					   m_pe(NULL)
	{}

	THashIterator (Tcl_HashTable *pTable) :
	m_pt(pTable)
	{
		ASSERT (m_pt!=NULL);
		m_pe = Tcl_FirstHashEntry (m_pt, &m_s);
	}

	THashIterator (THashIterator<K,V> &src)
	{
		*this = src;
	}

	virtual ~THashIterator ()
	{}

	V operator * ()
	{
		if (m_pe == NULL)
			throw ("null hash iterator");
		return (V)Tcl_GetHashValue (m_pe);
	}

	THashIterator &operator ++ ()
	{
		if (m_pe != NULL)
			m_pe = Tcl_NextHashEntry (&m_s);
		return *this;
	}


	THashIterator &operator ++ (int)
	{
		if (m_pe != NULL)
			m_pe = Tcl_NextHashEntry (&m_s);
		return *this;
	}

	operator Tcl_HashEntry* () 
	{
		return m_pe;
	}

	bool operator!= (Tcl_HashEntry *pEntry)
	{
		return m_pe != pEntry;
	}

	bool operator== (Tcl_HashEntry *pEntry)
	{
		return m_pe == pEntry;
	}

	K* key () {
		ASSERT (m_pt != NULL);
		if (m_pe == NULL)
			throw ("null hash iterator");
		return (K*)Tcl_GetHashKey (m_pt, m_pe);
	}

	THashIterator<K,V> &operator = (THashIterator<K,V> &i)
	{
		m_pt = i.m_pt;
		m_pe = i.m_pe;
		m_s = i.m_s;
		return *this;
	}
};




template <class K, class V, int Size=sizeof(K)/sizeof(int)>
class THash 
{
public: 
	typedef THashIterator<K,V> iterator;
protected:
	int				m_keytype;
	bool			m_bCreated;
	Tcl_HashTable	m_tbl;
public:
	THash ():
	  m_keytype(Size),
	  m_bCreated(false)
	{
	}

	~THash ()
	{
		deltbl();
	}



	iterator begin ()
	{
		createtbl();
		iterator i(&m_tbl);
		return i;
	}

	iterator end ()
	{
		createtbl();
		iterator i;
		return i;
	}


	Tcl_HashEntry *find (const K *key, V *value = NULL)
	{
		Tcl_HashEntry *p = NULL;
		createtbl();
		
		p = Tcl_FindHashEntry (&m_tbl, (char*)key);
		if (value != NULL && p!=NULL)
			*value = (V)Tcl_GetHashValue (p);
		return p;
	}


	bool delete_entry (const K *key)
	{
		Tcl_HashEntry *p = find (key);
		if (p!=NULL) 
			Tcl_DeleteHashEntry (p);
		return (p!=NULL);
	}


	Tcl_HashEntry * create_entry (const K *key, int *created = NULL)
	{
		ASSERT (key != NULL);
		createtbl();

		int c;
		Tcl_HashEntry *p;

		if (created == NULL)
			p = Tcl_CreateHashEntry (&m_tbl, (char*)key, &c);
		else
			p = Tcl_CreateHashEntry (&m_tbl, (char*)key, created);
		return p;
	}

	Tcl_HashEntry * set (const K *key, const V &value)
	{
		ASSERT (key != NULL);
		createtbl();

		Tcl_HashEntry *p = create_entry (key);
		if (p!=NULL)
			Tcl_SetHashValue (p, (ClientData)value);
		return p;
	}

	K *key (const Tcl_HashEntry *p)
	{
		ASSERT (p!=NULL);
		if (!m_bCreated) return NULL;
		return (K*)Tcl_GetHashKey (&m_tbl, p);
	}

	operator Tcl_HashTable*()
	{
		return &m_tbl;
	}
	
	void deltbl ()
	{
		if (m_bCreated) {
			Tcl_DeleteHashTable (&m_tbl);
			m_bCreated = false;
		}
	}

	void createtbl ()
	{
		if (!m_bCreated) {
			Tcl_InitHashTable (&m_tbl, m_keytype);
			m_bCreated = true;
		}
	}

};



class TDString {
protected:
	Tcl_DString ds;
public:
	TDString ()
	{
		Tcl_DStringInit(&ds);
	}

	TDString (const char *init)
	{
		Tcl_DStringInit(&ds);
		append(init);
	}

	~TDString ()
	{
		Tcl_DStringFree(&ds);
	}

	TDString& set (const char *string = "")
	{
		ASSERT (string != NULL);
		Tcl_DStringFree (&ds);
		Tcl_DStringInit(&ds);
		append(string);
		return *this;
	}

	char *append (const char *string, int length = -1)
	{
		ASSERT (string != NULL);
		return Tcl_DStringAppend (&ds, (char*)string, length);
	}

	TDString& operator<< (const char *string)
	{
		ASSERT (string!=NULL);
		append(string);
		return *this;
	}

	TDString& operator<< (const long val)
	{
		TObjPtr p(val);
		append((char*)p);
		return *this;
	}

	TDString& operator<< (const int val)
	{
		TObjPtr p(val);
		append((char*)p);
		return *this;
	}

	TDString& operator<< (const double fval)
	{
		TObjPtr d(fval);
		append((char*)d);
		return *this;
	}

	operator const char*()
	{
		return value();
	}

	// type unsafe, as the string still belongs to this object
	operator char*()
	{
		return (char*)(value());
	}

	TDString& operator= (TDString & src)
	{
		set (src.value());
		return *this;
	}

	char *append_element(char *string)
	{
		ASSERT (string != NULL);
		return Tcl_DStringAppendElement (&ds, (char*)string);
	}

	void start_sublist ()
	{
		Tcl_DStringStartSublist (&ds);
	}

	void end_sublist ()
	{
		Tcl_DStringEndSublist (&ds);
	}

	int length ()
	{
		return Tcl_DStringLength (&ds);
	}

	const char *value ()
	{
		return Tcl_DStringValue (&ds);
	}

	void set_result (Tcl_Interp *pInterp)
	{
		Tcl_DStringResult (pInterp, &ds);
	}

	void get_result (Tcl_Interp *pInterp)
	{
		Tcl_DStringGetResult (pInterp, &ds);
	}
};



#endif // _3CC705E0_BA28_11d2_8003_0040055861F2_
