// comvector.h: Wrapper around a single-dim SAFEARRAY and a lock on a SAFEARRAY
/////////////////////////////////////////////////////////////////////////////
// Copyright (c) 1999, Chris Sells. All rights reserved. No warranties extended.
// Comments to csells@sellsbrothers.com.
/////////////////////////////////////////////////////////////////////////////
// CComVector is a wrapper around a single-dim SAFEARRAY.
// CComVectorData is a wrapper around a lock on a SAFEARRAY, acquired via
// SafeArrayAccessData.
//
// By splitting the two concepts into two classes, each can be managed
// separately, therefore avoiding the overhead and inconvenience of
// SafeArrayGetElement and SafeArrayPutElement.
//
// By specializing on one dimension, the language mapping is easy.
/////////////////////////////////////////////////////////////////////////////
// History:
//  5/13/99:    Initial release.
/////////////////////////////////////////////////////////////////////////////
// Usage:
/*
STDMETHODIMP CFoo::UseSafeArray(SAFEARRAY** ppsa) // [in]
{
    CComVectorData<BSTR> rg(*ppsa);
    if( !rg ) return E_UNEXPECTED;

    for( int i = 0; i < rg.Length(); ++i )
    {
        TCHAR   sz[32];
        wsprintf(sz, __T("rg[%d]= %S\n"), i, (rg[i] ? rg[i] : OLESTR("")));
        OutputDebugString(sz);
    }

    // Assume data member: 'CComVector<BSTR> m_rg;'
    m_rg = *ppsa;   // Cache a copy

    return S_OK;
}

STDMETHODIMP CFoo::GetSafeArray(SAFEARRAY** ppsa) // [in, out]
{
    CComVector<BSTR> v(100);
    if( !v ) return E_OUTOFMEMORY;

    CComVectorData<BSTR> rg(v);
    if( !rg ) return E_UNEXPECTED;

    for( int i = 0; i < 100; ++i )
    {
        char    sz[16];
        rg[i] = A2BSTR(itoa(i, sz, 10));
    }

    return v.DetachTo(ppsa);
}
*/

#pragma once
#ifndef INC_COMVECTOR
#define INC_COMVECTOR

#include <crtdbg.h>

#ifndef HR
#define HR(_ex) { HRESULT _hr = _ex; if( FAILED(_hr) ) return _hr; }
#endif

/////////////////////////////////////////////////////////////////////////////
// CComVector wraps a SAFEARRAY*

class CComVectorBase
{
public:
    CComVectorBase(SAFEARRAY* psa) : m_psa(0)
    {
        if( psa ) Copy(psa);
    }

    CComVectorBase(const CComVectorBase& v) : m_psa(0)
    {
        if( v.m_psa ) Copy(v.m_psa);
    }

    ~CComVectorBase()
    {
        Destroy();
    }

    HRESULT Copy(const SAFEARRAY* psa)
    {
        Destroy();

        if( psa )
        {
            _ASSERTE(SafeArrayGetDim(const_cast<SAFEARRAY*>(psa)) == 1);
            HR(SafeArrayCopy(const_cast<SAFEARRAY*>(psa), &m_psa));
        }

        return S_OK;
    }

    HRESULT Destroy()
    {
        if( !m_psa ) return S_OK;

        HRESULT hr = SafeArrayDestroy(m_psa);
        m_psa = 0;
        return hr;
    }

    void Attach(SAFEARRAY* psa)
    {
        Destroy();
        m_psa = psa;
    }

    SAFEARRAY* Detach()
    {
        SAFEARRAY*  psa = m_psa;
        m_psa = 0;
        return psa;
    }

    HRESULT DetachTo(/*[in, out]*/ SAFEARRAY** ppsa)
    {
        if( !ppsa ) HR(E_POINTER);

        // If the [in] array is already created, destroy it
        if( *ppsa ) HR(SafeArrayDestroy(*ppsa));

        *ppsa = m_psa;
        m_psa = 0;
        return S_OK;
    }

    HRESULT CopyTo(/*[in, out]*/ SAFEARRAY** ppsa)
    {
        if( !ppsa ) HR(E_POINTER);

        // If the [in] array is already created, destroy it
        if( *ppsa ) HR(SafeArrayDestroy(*ppsa));

        HR(SafeArrayCopy(m_psa, ppsa));
        return S_OK;
    }

    operator SAFEARRAY*()
    {
        return m_psa;
    }

    SAFEARRAY** operator&()
    {
        _ASSERTE(!m_psa);
        return &m_psa;
    }

    long Length()
    {
        return Length(m_psa);
    }

public:
    static
    long Length(const SAFEARRAY* psa)
    {
        if( !psa ) return 0;
        _ASSERTE(SafeArrayGetDim(const_cast<SAFEARRAY*>(psa)) == 1);

        long    ub; SafeArrayGetUBound(const_cast<SAFEARRAY*>(psa), 1, &ub);
        long    lb; SafeArrayGetLBound(const_cast<SAFEARRAY*>(psa), 1, &lb);
        return ub - lb + 1;
    }

protected:
    template <typename T> VARTYPE VarType(T*);

    template<> VARTYPE VarType(LONG*) { return VT_I4; }
    template<> VARTYPE VarType(BYTE*) { return VT_UI1; }
    template<> VARTYPE VarType(SHORT*) { return VT_I2; }
    template<> VARTYPE VarType(FLOAT*) { return VT_R4; }
    template<> VARTYPE VarType(DOUBLE*) { return VT_R8; }
    template<> VARTYPE VarType(VARIANT_BOOL*) { return VT_BOOL; }
    template<> VARTYPE VarType(SCODE*) { return VT_ERROR; }
    template<> VARTYPE VarType(CY*) { return VT_CY; }
    template<> VARTYPE VarType(DATE*) { return VT_DATE; }
    template<> VARTYPE VarType(BSTR*) { return VT_BSTR; }
    template<> VARTYPE VarType(IUnknown **) { return VT_UNKNOWN; }
    template<> VARTYPE VarType(IDispatch **) { return VT_DISPATCH; }
    template<> VARTYPE VarType(SAFEARRAY **) { return VT_ARRAY; }
    template<> VARTYPE VarType(VARIANT *) { return VT_VARIANT; }
    template<> VARTYPE VarType(CHAR*) { return VT_I1; }
    template<> VARTYPE VarType(USHORT*) { return VT_UI2; }
    template<> VARTYPE VarType(ULONG*) { return VT_UI4; }
    template<> VARTYPE VarType(INT*) { return VT_INT; }
    template<> VARTYPE VarType(UINT*) { return VT_UINT; }

    template<> VARTYPE VarType(BYTE **) { return VT_BYREF|VT_UI1; }
    template<> VARTYPE VarType(SHORT **) { return VT_BYREF|VT_I2; }
    template<> VARTYPE VarType(LONG **) { return VT_BYREF|VT_I4; }
    template<> VARTYPE VarType(FLOAT **) { return VT_BYREF|VT_R4; }
    template<> VARTYPE VarType(DOUBLE **) { return VT_BYREF|VT_R8; }
    template<> VARTYPE VarType(VARIANT_BOOL **) { return VT_BYREF|VT_BOOL; }
    template<> VARTYPE VarType(SCODE **) { return VT_BYREF|VT_ERROR; }
    template<> VARTYPE VarType(CY **) { return VT_BYREF|VT_CY; }
    template<> VARTYPE VarType(DATE **) { return VT_BYREF|VT_DATE; }
    template<> VARTYPE VarType(BSTR **) { return VT_BYREF|VT_BSTR; }
    template<> VARTYPE VarType(IUnknown ***) { return VT_BYREF|VT_UNKNOWN; }
    template<> VARTYPE VarType(IDispatch ***) { return VT_BYREF|VT_DISPATCH; }
    template<> VARTYPE VarType(SAFEARRAY ***) { return VT_BYREF|VT_ARRAY; }
    template<> VARTYPE VarType(VARIANT **) { return VT_BYREF|VT_VARIANT; }
    template<> VARTYPE VarType(PVOID*) { return VT_BYREF; }
    template<> VARTYPE VarType(DECIMAL **) { return VT_BYREF|VT_DECIMAL; }
    template<> VARTYPE VarType(CHAR **) { return VT_BYREF|VT_I1; }
    template<> VARTYPE VarType(USHORT **) { return VT_BYREF|VT_UI2; }
    template<> VARTYPE VarType(ULONG **) { return VT_BYREF|VT_UI4; }
    template<> VARTYPE VarType(INT **) { return VT_BYREF|VT_INT; }
    template<> VARTYPE VarType(UINT **) { return VT_BYREF|VT_UINT; }

protected:
    SAFEARRAY*  m_psa;
};

template <typename T>
class CComVector : public CComVectorBase
{
public:
    CComVector(SAFEARRAY* psa = 0) : CComVectorBase(psa)
    {
    }

    CComVector(const CComVector& v) : CComVectorBase(v)
    {
    }

    CComVector(long celt, long nLowerBound = 0, bool bZeroMemory = true) : CComVectorBase(0)
    {
        Create(celt, nLowerBound, bZeroMemory);
    }

    // TODO: Hook up Copy policy classes
    /*
    CComVector(const T* rg, long celt) : m_psa(0)
    {
        Copy(rg, celt);
    }
    */

    CComVector& operator=(const SAFEARRAY* psa)
    {
        if( psa != m_psa ) Copy(psa);
        return *this;
    }

    CComVector& operator=(const CComVector& v)
    {
        if( &v != this ) Copy(v.m_psa);
        return *this;
    }

    HRESULT Create(long celt, long nLowerBound = 0, bool bZeroMemory = true)
    {
        SAFEARRAYBOUND  sab = { celt, nLowerBound };
        m_psa = SafeArrayCreate(VarType((T*)0), 1, &sab);
        if( !m_psa ) HR(E_OUTOFMEMORY);

        if( bZeroMemory )
        {
            T*      rg = 0;
            HR(SafeArrayAccessData(m_psa, (void**)&rg));
            ZeroMemory(rg, sizeof(T) * celt);
            SafeArrayUnaccessData(m_psa);
        }

        return S_OK;
    }
};

/////////////////////////////////////////////////////////////////////////////
// CComVectorData represents a lock on SAFEARRAY data

class CComVectorDataBase
{
public:
    CComVectorDataBase(SAFEARRAY* psa) : m_psa(0), m_pv(0), m_celt(0)
    {
        if( psa ) AccessData(psa);
    }

    ~CComVectorDataBase()
    {
        UnaccessData();
    }

    HRESULT AccessData(SAFEARRAY* psa)
    {
        if( !psa || (SafeArrayGetDim(psa) != 1) ) return E_INVALIDARG;
        UnaccessData();

        HR(SafeArrayAccessData(psa, &m_pv));
        m_psa = psa;
        m_celt = CComVectorBase::Length(m_psa);
        return S_OK;
    }

    HRESULT UnaccessData()
    {
        if( m_psa && m_pv ) HR(SafeArrayUnaccessData(m_psa));
        m_psa = 0;
        m_pv = 0;
        m_celt = 0;
        return S_OK;
    }

    long Length()
    {
        return m_celt;
    }

    operator bool()
    {
        return (m_pv ? true : false);
    }
/*
    operator const T*()
    {
        return reinterpret_cast<T*>(m_pv);
    }
*/

protected:
    SAFEARRAY*  m_psa;
    void*       m_pv;
    long        m_celt;
};

template <typename T>
class CComVectorData : public CComVectorDataBase
{
public:
    CComVectorData(SAFEARRAY* psa = 0) : CComVectorDataBase(psa)
    {
    }

    T& operator[](long n)
    {
#ifdef _DEBUG
        if( n < 0 || n >= m_celt )
        {
            _ASSERTE(false && "Accessing data out of bounds");
        }
#endif
        return reinterpret_cast<T*>(m_pv)[n];
    }
};

#endif  // INC_COMVECTOR
