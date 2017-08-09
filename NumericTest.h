#pragma once

#include "StreamSupport.h"


template <typename T>
inline T SetData()
{
    return 34;
}

inline CComBSTR SetData()
{
    return L"Test string";
}


template< class T>
class CNumericTest
{
public:


    //------------------------------------------------------------------------------
    // Test a one dimensional array of numbers
    //------------------------------------------------------------------------------

    static HRESULT Test()
    {
        CComVariant         v1;
        CComVariant         v2;
        CComPtr<IStream>    pStream;

        USES_CONVERSION;

        v1 = SetData<T>();

        // Create a memory stream.
        HR( CreateMemoryStream( &pStream ) );

        // Write out the variant to the stream, rewind the stream and read it back into another variant.
        WriteVariantToStream( &v1, pStream );
        HR( RewindStream( pStream ) );
        ReadVariantFromStream( pStream, v2 );

        if ( v1 != v2 )
            HR( E_UNEXPECTED );

        return S_OK;

    } // Test


}; // class COneDimNumbersArrayTest  
