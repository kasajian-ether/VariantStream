#pragma once

#include "StreamSupport.h"

class COneDimStringArrayTest
{
public:

    //------------------------------------------------------------------------------
    // Creates an array of 10 strings, the numbers "0", "1", "2", etc.
    //------------------------------------------------------------------------------

    static HRESULT GetArrayOfTenStrings( SAFEARRAY*& safearray )
    {
        int const arraySize = 10;
        CComVector<BSTR> a(arraySize);

        CComVectorData<BSTR> rg(a);
        if ( !rg )
            HR( E_UNEXPECTED );

        for( int i = 0; i < arraySize; ++i )
        {
            CComVariant val = (long)i;
            HR( val.ChangeType( VT_BSTR ) );
            rg[i] = val.bstrVal;
            VARIANT dummy;
            val.Detach( &dummy );
        }

        safearray = a.Detach();

        return S_OK;

    } // GetArrayOfTenStrings


    //------------------------------------------------------------------------------
    // Verifies that the given two string arrays have the same content
    //------------------------------------------------------------------------------

    static HRESULT VerifyArrayOfTenStrings( SAFEARRAY* array1, SAFEARRAY* array2 )
    {
        CComVectorData<BSTR> rg1( array1 );
        CComVectorData<BSTR> rg2( array2 );

        if ( rg1.Length() != rg2.Length() )
            HR( E_UNEXPECTED );
        
        for( int i = 0; i < rg1.Length(); ++i )
        {
            if ( lstrcmpW( rg1[i], rg2[i] ) )
                HR( E_UNEXPECTED );
        }

        return S_OK;

    } // VerifyArrayOfTenStrings


    //------------------------------------------------------------------------------
    // Test a one dimensional array of strings
    //------------------------------------------------------------------------------

    static HRESULT Test()
    {
        CComVariant         v1;
        CComVariant         v2;
        CComPtr<IStream>    pStream;

        USES_CONVERSION;

        // Get an array of 10 strings, the numbers "0", "1", "2", etc.
        v1.vt = VT_BSTR | VT_ARRAY;
        HR( GetArrayOfTenStrings( v1.parray ) );

        // Create a memory stream.
        HR( CreateMemoryStream( &pStream ) );

        // Write out the variant to the stream, rewind the stream and read it back into another variant.
        WriteVariantToStream( &v1, pStream );
        HR( RewindStream( pStream ) );
        ReadVariantFromStream( pStream, v2 );

        // Verify that the new array is the same as the old.
        HR( VerifyArrayOfTenStrings( v1.parray, v2.parray ) );

        return S_OK;

    } // Test


}; // class COneDimStringArrayTest  
