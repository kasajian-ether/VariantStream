#pragma once

#include "StreamSupport.h"

template< class T, VARTYPE VAR_T>
class COneDimNumericArrayTest
{
public:

    //------------------------------------------------------------------------------
    // Creates an array of 10 numbers, the numbers 0, 1, 2, etc.
    //------------------------------------------------------------------------------

    static HRESULT GetArrayOfTenNumbers( SAFEARRAY*& safearray )
    {
        int const arraySize = 10;
        CComVector<T> a(arraySize);

        CComVectorData<T> rg(a);
        if ( !rg )
            HR( E_UNEXPECTED );

        for( int i = 0; i < arraySize; ++i )
            rg[i] = (T)i;

        safearray = a.Detach();

        return S_OK;

    } // GetArrayOfTenNumbers


    //------------------------------------------------------------------------------
    // Verifies that the given two arrays of T have the same content
    //------------------------------------------------------------------------------

    static HRESULT VerifyArrayOfTenNumbers( SAFEARRAY* array1, SAFEARRAY* array2 )
    {
        CComVectorData<T> rg1( array1 );
        CComVectorData<T> rg2( array2 );

        if ( rg1.Length() != rg2.Length() )
            HR( E_UNEXPECTED );
        
        for( int i = 0; i < rg1.Length(); ++i )
        {
            if ( rg1[i] != rg2[i] )
                HR( E_UNEXPECTED );
        }

        return S_OK;

    } // VerifyArrayOfTenNumbers


    //------------------------------------------------------------------------------
    // Test a one dimensional array of numbers
    //------------------------------------------------------------------------------

    static HRESULT Test()
    {
        CComVariant         v1 = 0;
        CComVariant         v2 = 0;
        CComPtr<IStream>    pStream;

        USES_CONVERSION;

        // Get an array of 10 numbers, the numbers 0, 1, 2, etc.
        v1.vt = VAR_T | VT_ARRAY;
        HR( GetArrayOfTenNumbers( v1.parray ) );

        // Create a memory stream.
        HR( CreateMemoryStream( &pStream ) );

        // Write out the variant to the stream, rewind the stream and read it back into another variant.
        WriteVariantToStream( &v1, pStream );
        HR( RewindStream( pStream ) );
        ReadVariantFromStream( pStream, v2 );

        // Verify that the new array is the same as the old.
        HR( VerifyArrayOfTenNumbers( v1.parray, v2.parray ) );

        return S_OK;

    } // Test


}; // class COneDimNumbersArrayTest  
