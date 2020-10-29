#include <atlbase.h>
#include "VariantStream.h"
#include "comvector.h"
#include "OneDimStringArrayTest.h"
#include "OneDimNumericArrayTest.h"
#include "OneDimVariantArrayTest.h"
#include "NumericTest.h"
#include "NonValuetest.h"

#pragma warning( disable: 4711 )

#import "msado15.tlb"               \
    rename("EOF", "EndOfFile"),     \
    named_guids,                    \
    no_implementation,              \
    raw_interfaces_only,            \
    raw_method_prefix(""),          \
    raw_native_types,               \
    no_namespace

#include "ObjectTest.h"


//------------------------------------------------------------------------------
// TestArrays
//------------------------------------------------------------------------------

static HRESULT TestArrays()
{
    HRESULT hr;

    // Test a one dimensional array of strings
    HR( COneDimStringArrayTest::Test() );

    // Test a one dimensional array of various types.
    hr = COneDimNumericArrayTest<SCODE, VT_ERROR>::Test();
    HR( hr );

    hr = COneDimNumericArrayTest<long, VT_I4>::Test();
    HR( hr );

    hr = COneDimNumericArrayTest<BYTE, VT_UI1>::Test();
    HR( hr );

    hr = COneDimNumericArrayTest<short, VT_I2>::Test();
    HR( hr );

    hr = COneDimNumericArrayTest<double, VT_R8>::Test();
    HR( hr );

    hr = COneDimNumericArrayTest<DATE, VT_DATE>::Test();
    HR( hr );

    hr = COneDimNumericArrayTest<float, VT_R4>::Test();
    HR( hr );

    hr = COneDimVariantArrayTest::Test();
    HR( hr );

    return S_OK;

} // TestArrays


//------------------------------------------------------------------------------
// TestNonArrays
//------------------------------------------------------------------------------

static HRESULT TestNonArrays()
{
    // Test various types.
    HR( CNumericTest<CComBSTR>::Test() );
    HR( CNumericTest<SCODE>::Test() );
    HR( CNumericTest<long>::Test() );
    HR( CNumericTest<BYTE>::Test() );
    HR( CNumericTest<short>::Test() );
    HR( CNumericTest<double>::Test() );
    HR( CNumericTest<DATE>::Test() );
    HR( CNumericTest<float>::Test() );

    return S_OK;

} // TestNonArrays


//------------------------------------------------------------------------------
// TestNonValueTypes
//------------------------------------------------------------------------------

static HRESULT TestNonValueTypes()
{
    // Test various types.
    HR( CNonValueTest::Test( VT_EMPTY ) );
    HR( CNonValueTest::Test( VT_NULL ) );

    return S_OK;

} // TestNonValueTypes


//------------------------------------------------------------------------------
// TestNonArrays
//------------------------------------------------------------------------------

static HRESULT TestObject()
{
    return CObjectTest::Test();

} // TestObject


//------------------------------------------------------------------------------
// Start
//------------------------------------------------------------------------------

HRESULT Start()
{
    // Test non-array access.
    HR( TestNonArrays() );

    HR( TestNonValueTypes() );

    // Test arrays
    HR( TestArrays() );

    // Test object access
    return TestObject();

} // Start


//------------------------------------------------------------------------------
// main
//------------------------------------------------------------------------------

int WINAPI WinMain( HINSTANCE, HINSTANCE, LPSTR, int )
{
    LPCTSTR     msg;
    HRESULT     hr;

    ::CoInitialize( NULL );

    hr = Start();
    if ( SUCCEEDED( hr ) )
    {
        if ( S_FALSE == hr )
            msg = _T( "All tests successful except for object streaming -- Coult not create test ADO Recordset." );
        else
            msg = _T( "All tests successful" );
    }
    else
    {
        msg = _T( "Failed" );
    }

    ::MessageBox( 0, msg, _T( "" ), 0 );

    ::CoUninitialize();

    return 0;

} // main
