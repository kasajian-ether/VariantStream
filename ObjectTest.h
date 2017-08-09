#pragma once

#include "StreamSupport.h"

long const          number = 34L;
WCHAR const* const  fieldName = L"test";

class CObjectTest
{
public:

    template <class Q>
    static HRESULT TestObject( Q object )
    {
        CComPtr<IStream>        pStream;
        CComVariant             v1;
        CComVariant             v2;
        CComPtr<_Recordset>     recordset;
        CComPtr<IUnknown>       recordsetUnknown;
        CComPtr<Fields>         fields;
        CComPtr<Field>          field;

        // Set the object as IUknown
        v1 = object;

        // Create a memory stream.
        HR( CreateMemoryStream( &pStream ) );

        // Write out the variant to the stream, rewind the stream and read it back into another variant.
        WriteVariantToStream( &v1, pStream );
        HR( RewindStream( pStream ) );
        ReadVariantFromStream( pStream, v2 );

        // Convert the variant into a recordset.
        HR( v2.ChangeType( VT_UNKNOWN ) );
        recordsetUnknown = v2.punkVal;
        v2.punkVal->AddRef();
        recordsetUnknown.QueryInterface( &recordset );

        // Get the recordset count.
        long recordCount;
        HR( recordset->get_RecordCount( &recordCount ) );
        if ( 1 != recordCount )
            HR( E_UNEXPECTED );

        // Get the field object so we can read its data.
        HR( recordset->get_Fields( &fields ) );
        CComVariant variantFieldName( fieldName );
        HR( fields->get_Item( variantFieldName, &field ) );

        CComVariant variantNumberRead;
        HR( field->get_Value( &variantNumberRead ) );
        HR( variantNumberRead.ChangeType( VT_I4 ) );
        long numberRead = variantNumberRead.lVal;
        if ( number != numberRead )
            HR( E_UNEXPECTED );

        return S_OK;
    }

    //------------------------------------------------------------------------------
    // Test object streaming
    //------------------------------------------------------------------------------

    static HRESULT Test()
    {
        CComVariant missing( DISP_E_PARAMNOTFOUND, VT_ERROR );

        CComPtr<_Recordset>     recordset;
        CComPtr<Fields>         fields;
        CComPtr<Field>          field;

        // Create a new record set object.
        // If the recordset cannot be created, we can use ADO to test object streaming.
        // Just tell the user.
        if ( FAILED( recordset.CoCreateInstance( CLSID_Recordset ) ) )
            return S_FALSE;

        // Get the fields
        HR( recordset->get_Fields( &fields ) );
        HR( fields->Append( CComBSTR( fieldName ), adInteger, sizeof( BSTR ), adFldUnspecified ) );
        HR( recordset->Open( missing, missing, adOpenUnspecified, adLockUnspecified, 0 ) );
        HR( recordset->AddNew() );

        CComVariant variantFieldName( fieldName );
        HR( fields->get_Item( variantFieldName, &field ) );

        CComVariant variantNumber( number );
        HR( field->put_Value( variantNumber ) );

        CComPtr<IUnknown>     unknown = recordset;
        TestObject( (IUnknown*)unknown );

        CComPtr<IDispatch>     dispatch = recordset;
        HR( TestObject( (IDispatch*)dispatch ) );

        return S_OK;
    } // Test


}; // class CObjectTest
