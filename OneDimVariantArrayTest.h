#pragma once

#include "StreamSupport.h"

class COneDimVariantArrayTest
{
public:

    //------------------------------------------------------------------------------
    // Creates an array of 10 numbers, the numbers 0, 1, 2, etc.
    //------------------------------------------------------------------------------

    static HRESULT GetArrayOfTenNumbersAsVariants( SAFEARRAY*& safearray )
    {
        int const arraySize = 10;
        CComVector<VARIANT> a(arraySize);

        CComVectorData<VARIANT> rg(a);
        if ( !rg )
            HR( E_UNEXPECTED );

        for( int i = 0; i < arraySize; ++i )
        {
            rg[i].vt = VT_I4;
            rg[i].lVal = i;
        }

        safearray = a.Detach();

        return S_OK;

    } // GetArrayOfTenNumbersAsVariants


    //------------------------------------------------------------------------------
    // Verifies that the given two arrays of VARIANT have the same content
    //------------------------------------------------------------------------------

    static HRESULT VerifyArrayOfTenNumbersAsVariants( SAFEARRAY* array1, SAFEARRAY* array2 )
    {
        CComVectorData<VARIANT> rg1( array1 );
        CComVectorData<VARIANT> rg2( array2 );

        if ( rg1.Length() != rg2.Length() )
            HR( E_UNEXPECTED );
        
        for( int i = 0; i < rg1.Length(); ++i )
        {
            if ( rg1[i].vt != VT_I4 )
                HR( E_UNEXPECTED );
            if ( rg2[i].vt != VT_I4 )
                HR( E_UNEXPECTED );

            if ( rg1[i].lVal != rg2[i].lVal )
                HR( E_UNEXPECTED );
        }

        return S_OK;

    } // VerifyArrayOfTenNumbersAsVariants


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
        v1.vt = VT_VARIANT | VT_ARRAY;
        HR( GetArrayOfTenNumbersAsVariants( v1.parray ) );

        // Create a memory stream.
        HR( CreateMemoryStream( &pStream ) );

        // Write out the variant to the stream, rewind the stream and read it back into another variant.
        WriteVariantToStream( &v1, pStream );
        HR( RewindStream( pStream ) );
        ReadVariantFromStream( pStream, v2 );

        // Verify that the new array is the same as the old.
        HR( VerifyArrayOfTenNumbersAsVariants( v1.parray, v2.parray ) );

        return S_OK;

    } // Test


}; // class COneDimNumbersArrayTest  
