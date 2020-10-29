//==============================================================================
// Public domain variant streaming code.
//
// Use global functions WriteVariantToStream and ReadVariantFromStream to
// read and write a variant to a stream.
// Use global functions ReadVariantFromBlob and WriteVariantToBlob to
// read and write a variant to a blob.
//
//==============================================================================

#pragma once

#include <atlbase.h>
#include "stream.h"


//==============================================================================
// namespace VariantStreaming
// Internal namespace used to keep support calls in this header file private.
//==============================================================================

namespace VariantStreaming
{


//==============================================================================
// Prototype
//==============================================================================

inline void  WriteDataToStream( const VARIANT* variant, IStream* pStream );
inline void  ReadDataFromStream( VARTYPE vt, IStream* pStream, VARIANT& variant );



//==============================================================================
// Types
//==============================================================================

const long variantVersion = 1;


//==============================================================================
// CWalkSafeArrayElements
// Walks the elements of a multi-dimensional safe array.  Array must have at
// least one element.
// Example:
//      bool more = true;
//
//      CWalkSafeArrayElements  walk( safeArray );
//
//      // Walk the list.
//      while( more )
//          {
//          ... walk.GetIndex( index ) ...
//
//          more = walk.Next();
//          }
//==============================================================================

class CWalkSafeArrayElements
{
public:
    CWalkSafeArrayElements( SAFEARRAY* safeArray )
        :   m_SafeArray( safeArray ),
            m_index( NULL ),
            m_isInitialized( false )
    {
    }
    
    inline ~CWalkSafeArrayElements()
    {
        ::CoTaskMemFree( m_index );
    }
    
    inline void GetIndex( long*& index )
    {
        if ( !m_isInitialized )
            Initialize();

        index = m_index;
    }
    
    void Next( bool& more )
    {
        if ( !m_isInitialized )
            Initialize();

        more = false;
        
        // Increment the index and determine if we've hit the end.
        for ( int dimension = 0; dimension < m_SafeArray->cDims; dimension++ )
        {
            m_index[m_SafeArray->cDims - dimension - 1]++;
            
            if (    m_index[m_SafeArray->cDims - dimension - 1] >=
                m_SafeArray->rgsabound[dimension].lLbound + (long)m_SafeArray->rgsabound[dimension].cElements )
            {
                m_index[m_SafeArray->cDims - dimension - 1] = m_SafeArray->rgsabound[dimension].lLbound;
            }
            else
            {
                more = true;
                break;
            }
        }
    }

private:
    void Initialize()
    {
        // Allocate an index
        m_index = (long*)::CoTaskMemAlloc( m_SafeArray->cDims * sizeof( long ) );
        VerifyAllocation( m_index );
        
        // Initialize the index.
        for ( int dimension = 0; dimension < m_SafeArray->cDims; dimension++ )
            m_index[m_SafeArray->cDims - dimension - 1] = m_SafeArray->rgsabound[dimension].lLbound;

        m_isInitialized = true;
    }
    
private:
    bool                m_isInitialized;
    SAFEARRAY*          m_SafeArray;
    long*               m_index;
    
}; // class CWalkSafeArrayElements


//------------------------------------------------------------------------------
// SafeArrayGetElementAsVariant
// Like SafeArrayGetElement, but returns the element value as a VARIANT.
//------------------------------------------------------------------------------

inline void SafeArrayGetElementAsVariant(    SAFEARRAY*      safeArray,
                                             long*           index,
                                             VARTYPE         vt,
                                             CComVariant&    variant )
{
    void*           entry;
    void*           element;

    // Allocate enough room to hold a single element of the array.
    element = _alloca( safeArray->cbElements );

    // Get the elemnt from the array.
    CheckResult( SafeArrayGetElement( safeArray, index, element ) );

    // If it's a string, assign the string to the variant.
    if ( ( safeArray->fFeatures & FADF_BSTR ) == FADF_BSTR )
    {
        if ( VT_BSTR != vt )
            ThrowError( DISP_E_TYPEMISMATCH );
    
        if ( sizeof( BSTR ) != safeArray->cbElements )
            ThrowError( E_INVALIDARG );
    
        CComBSTR  bstrEntry;
        entry = *(void**)(void*)element;
        bstrEntry.Attach( (BSTR) (void*)entry );
        variant = bstrEntry;
    }
    // If it's a generic COM object, assign the IUnknown to the variant.
    else if ( ( safeArray->fFeatures & FADF_UNKNOWN ) == FADF_UNKNOWN )
    {
        if ( VT_UNKNOWN != vt )
            ThrowError( DISP_E_TYPEMISMATCH );
    
        if ( sizeof( IUnknown* ) != safeArray->cbElements )
            ThrowError( DISP_E_TYPEMISMATCH );
    
        CComPtr<IUnknown>   unknownEntry;
        entry = *(void**)(void*)element;
        unknownEntry.Attach( (IUnknown*)entry );
        variant = unknownEntry;
    }
    // If it's an Automation object, assign the IDispatch to the variant.
    else if ( ( safeArray->fFeatures & FADF_DISPATCH ) == FADF_DISPATCH )
    {
        if ( VT_DISPATCH != vt )
            ThrowError( DISP_E_TYPEMISMATCH );
    
        if ( sizeof( IDispatch* ) != safeArray->cbElements )
            ThrowError( DISP_E_TYPEMISMATCH );
    
        CComPtr<IDispatch>  dispatchEntry;
        entry = *(void**)(void*)element;
        dispatchEntry.Attach( (IDispatch*)entry );
        variant = dispatchEntry;
    }
    // If the element itself is another variant, assign a dereferenced version
    // to the variant.
    else if ( ( safeArray->fFeatures & FADF_VARIANT ) == FADF_VARIANT )
    {
        if ( VT_VARIANT != vt )
            ThrowError( DISP_E_TYPEMISMATCH );
    
        if ( sizeof( VARIANT ) != safeArray->cbElements )
            ThrowError( DISP_E_TYPEMISMATCH );
    
        CComVariant   variantEntry;
        entry = element;
        variantEntry.Attach( (VARIANT*)entry );
    
        variant = variantEntry;
    }
    // If it's a simple type, assign it based on the type.
    else
    {
        entry = element;
    
        switch( vt )
        {
        case VT_ERROR:
            variant.scode = *(SCODE*) entry;
            break;
        
        case VT_I4:
            variant.lVal = *(long*) entry;
            break;
        
        case VT_UI1:
            variant.bVal = *(BYTE*) entry;
            break;
        
        case VT_I2:
            variant.iVal = *(short*) entry;
            break;
        
        case VT_BOOL:
#pragma warning(disable: 4310) // cast truncates constant value
		    variant.boolVal = *(VARIANT_BOOL*) entry ? VARIANT_TRUE : VARIANT_FALSE;
#pragma warning(default: 4310) // cast truncates constant value
            break;
        
        case VT_R8:
            variant.dblVal = *(double*) entry;
            break;
        
        case VT_DATE:
            variant.date = *(DATE*) entry;
            break;
        
        case VT_R4:
            variant.fltVal = *(float*) entry;
            break;
        
        case VT_CY:
            variant.cyVal = *(CY*)entry;
        
        default:
            ThrowError( DISP_E_TYPEMISMATCH );
        }
    
        // Set the variant's data type.
        variant.vt = vt;
    }

} // SafeArrayGetElementAsVariant


//------------------------------------------------------------------------------
// SafeArrayPutElementFromVariant
// Like SafeArrayPutElement, except that it gets the element's value from the
// variant
//------------------------------------------------------------------------------

inline void SafeArrayPutElementFromVariant(  SAFEARRAY*          safeArray,
                                             long*               index,
                                             const VARIANT&      variant )
{
    const void*     element = NULL;

    // If the array element type itself is a variant, then we just use
    // that variant directly.
    if ( ( safeArray->fFeatures & FADF_VARIANT ) == FADF_VARIANT )
    {
        element = &variant;
    }
    else
    {
        // Based on the data type of the variant, we obtain a direct reference
        // to the variant's data
        switch ( VT_TYPEMASK & variant.vt )
        {
        case VT_BSTR:
        case VT_UNKNOWN:
        case VT_DISPATCH:
            element = variant.byref;
            break;
        
        case VT_VARIANT:
            element = variant.pvarVal;
            break;
        
        case VT_ERROR:
        case VT_I4:
        case VT_UI1:
        case VT_I2:
        case VT_BOOL:
        case VT_R8:
        case VT_DATE:
        case VT_R4:
        case VT_CY:
            element = &variant.byref;
            break;
        
        default:
            ThrowError( DISP_E_TYPEMISMATCH );
        }
    }

    // Put the element into the array.
    CheckResult( SafeArrayPutElement( safeArray, index, const_cast<void*>( element ) ) );

} // SafeArrayPutElementFromVariant


//------------------------------------------------------------------------------
// GetTypeSize
// Given a variant type, determines how many bytes it takes to store data of
// that type.
//------------------------------------------------------------------------------

inline void GetTypeSize( VARTYPE vt, ULONG& size )
{
    switch ( vt ) 
    {
    case VT_BOOL:
        size = sizeof( VARIANT_BOOL );
        break;
    
    case VT_I2:
        size = sizeof( short );
        break;
    
    case VT_ERROR:
        size = sizeof( SCODE );
        break;
    
    case VT_I4:
        size = sizeof( long );
        break;
    
    case VT_BSTR:
        size = sizeof( BSTR );
        break;
    
    case VT_R4:
        size = sizeof( float );
        break;
    
    case VT_DATE:
        size = sizeof( DATE );
        break;
    
    case VT_R8:
        size = sizeof( double );
        break;
    
    case VT_UI1:
        size = sizeof( BYTE );
        break;
    
    case VT_CY:
        size = sizeof( CY );
        break;
    
    case VT_VARIANT:
        size = sizeof( VARIANT );
        break;
    
    case VT_DISPATCH:
        size = sizeof( IDispatch* );
        break;
    
    case VT_UNKNOWN:
        size = sizeof( IUnknown* );
        break;
    
    default:
        ThrowError( DISP_E_TYPEMISMATCH );
    }

} // GetTypeSize


//------------------------------------------------------------------------------
// WriteSafeArrayHeader
// Writes the array's header information, such as the number of dimensions
// and the bounds of each dimension (lower/upper)
//------------------------------------------------------------------------------

inline void WriteSafeArrayHeader( SAFEARRAY* safeArray, IStream* pStream )
{
    unsigned int        dimension;
    CStream             stream( pStream );

    // Write out the dimension count
    stream.Write( safeArray->cDims );

    // Write out the lower bound and the number of elements in each dimension.
    for ( dimension = 0; dimension < safeArray->cDims; dimension++ )
    {
        stream.Write( safeArray->rgsabound[dimension].lLbound );
        stream.Write( safeArray->rgsabound[dimension].cElements );
    }

} // WriteSafeArrayHeader


//------------------------------------------------------------------------------
// ReadSafeArrayHeader
//------------------------------------------------------------------------------

inline void ReadSafeArrayHeader( VARIANT* variant, VARTYPE vt, IStream* pStream )
{
    unsigned short      dimensions;
    unsigned short      dimension;
    CStream             stream( pStream );

    // Read the dimension count
    stream.Read( dimensions );

    CheckResult( SafeArrayAllocDescriptor( dimensions, &variant->parray ) );

    // Read the lower bound and the number of elements in this dimension.
    for ( dimension = 0; dimension < dimensions; dimension++ )
    {
        stream.Read( variant->parray->rgsabound[dimension].lLbound );
        stream.Read( variant->parray->rgsabound[dimension].cElements );
    }

    // Set the element size.
    GetTypeSize( vt, variant->parray->cbElements );

    // Set the features mask.
    switch ( vt )
    {
    case VT_BSTR:
        variant->parray->fFeatures |= FADF_BSTR;
        break;
    
    case VT_UNKNOWN:
        variant->parray->fFeatures |= FADF_UNKNOWN;
        break;
    
    case VT_DISPATCH:
        variant->parray->fFeatures |= FADF_DISPATCH;
        break;
    
    case VT_VARIANT:
        variant->parray->fFeatures |= FADF_VARIANT;
        break;
    }

    CheckResult( SafeArrayAllocData( variant->parray ) );

    // Set variant's type
    variant->vt = (VARTYPE) ( vt | VT_ARRAY );

} // ReadSafeArrayHeader



//------------------------------------------------------------------------------
// WriteSafeArrayElements
// Walks the elements of a multi-dimensional safe array, streaming each
// element out.
//------------------------------------------------------------------------------

inline void WriteSafeArrayElements( VARTYPE vt, SAFEARRAY* safeArray, IStream* pStream )
{
    bool                more = true;
    long*               index = NULL;
    CStream             stream( pStream );
    long                NoOfElements = 0;

    //check whether array has elements:
    for ( int dimension = 0; dimension < safeArray->cDims; dimension++ )
    {
        NoOfElements += safeArray->rgsabound[dimension].cElements;
    }
    if( NoOfElements == 0)
        return;
    CWalkSafeArrayElements  walk( safeArray );
        
    // Walk the list.
    while( more )
    {
        CComVariant         tempVariant;

        // Get the variant for this element and stream it out.
        walk.GetIndex( index );
        SafeArrayGetElementAsVariant( safeArray, index, vt, tempVariant );

        // If the array's type is VT_VARIANT, then write out the element's data
        // type so that it can be known when the variant is streamed back out.
        if ( VT_VARIANT == vt )
            stream.Write( tempVariant.vt );

        // Write the variant to the stream.
        WriteDataToStream( &tempVariant, stream );
    
        // Increment the index and determine if we've hit the end.
        walk.Next( more );
    }

} // WriteSafeArrayElements



//------------------------------------------------------------------------------
// ReadSafeArrayElements
// Read the elements from the stream into the safe array.
//------------------------------------------------------------------------------

inline void ReadSafeArrayElements( VARTYPE vt, SAFEARRAY* safeArray, IStream* pStream )
{
    bool        more = true;
    long*       index = NULL;
    CStream     stream( pStream );
    long        NoOfElements = 0;

    //check whether array has elements:
    for ( int dimension = 0; dimension < safeArray->cDims; dimension++ )
    {
        NoOfElements += safeArray->rgsabound[dimension].cElements;
    }
    if( NoOfElements == 0)
        return;
    CWalkSafeArrayElements  walk( safeArray );

    // Walk the list.
    while( more )
    {
        CComVariant tempVariant;

        VARTYPE elementType = vt;

        // If the array's data type is VT_VARIANT, then read the elements type.
        if ( VT_VARIANT == vt )
            stream.Read( elementType );

        // Get the variant from the stream and put it in the array as an element.
        ReadDataFromStream( elementType, stream, tempVariant );
        walk.GetIndex( index );
        SafeArrayPutElementFromVariant( safeArray, index, tempVariant );
    
        // Increment the index and determine if we've hit the end.
        walk.Next( more );
    }

} // ReadSafeArrayElements


//------------------------------------------------------------------------------
// GetPersistStreamInterface
// Get IPersistStream for the given object.
//------------------------------------------------------------------------------

inline void GetPersistStreamInterface( IUnknown* pUnknown, IPersistStream** ppPersistStream )
{
    CComPtr<IPersistStreamInit>     persistStreamInit;
    CComPtr<IPersistStream>         persistStream;
    HRESULT                         hr;

    // See if the object supports IPersistStreamInit.
    hr = pUnknown->QueryInterface( &persistStreamInit );
    if ( FAILED( hr ) && E_NOINTERFACE != hr )
        ThrowError( hr );

    // If the object doesn't support IPersistStreamInit, see
    // if supports IPersistStream
    if ( !persistStreamInit )
    {
        hr = pUnknown->QueryInterface( &persistStream );
        if ( FAILED( hr ) && E_NOINTERFACE != hr )
            ThrowError( hr );
    }

    // Return either IPersistStreamInit or IPersistStream, whichever
    // one was retrieved.  Since these interfaces are binary
    // compatible, we can use the same pointer to refer to either.
    if ( persistStream )
        *ppPersistStream = persistStream.Detach();
    else
        *ppPersistStream = (IPersistStream*)persistStreamInit.Detach();

} // GetPersistStreamInterface


//------------------------------------------------------------------------------
// SaveObjectToStream
// Wrapper for OleSaveToStream.  Difference is that this function
// takes an IUnknown instead of an IPersistStream and uses the IUnknown
// to get an IPersistStream using GetPersistStreamInterface, which is
// smart enough to QI for both IPersistStreamInit and IPersistStream.
// As with OleSaveToStream, the object pointer may be NULL.
//------------------------------------------------------------------------------

inline void SaveObjectToStream( IUnknown* pUnknown, IStream* pStream )
{
    CComPtr<IPersistStream>       persistStream;

    // Get IPersistStream interface for the object. 
    if ( pUnknown )
    {
        GetPersistStreamInterface( pUnknown, &persistStream );

        // If we didn't get IPersistStream, then fail.  For now,
        // these are the only two interfaces we're supporting.
        // Later, we can add support for others such as IStorage,
        // IPersistMemory, IPersistStorage, IPersistPropertyBag, etc.
        if ( !persistStream )
            ThrowError( E_NOINTERFACE );
    }

    // Save the object to the stream.
    CheckResult( OleSaveToStream( persistStream, pStream ) );

} // SaveObjectToStream


//------------------------------------------------------------------------------
// WriteDataToStream
// Writes the given variant's data to the stream.
// Used by WriteToStream.
// The passed in variant is assumed to be fully dereferenced (i.e. no VT_BYREF)
//------------------------------------------------------------------------------

inline void WriteDataToStream( const VARIANT* variant, IStream* pStream )
{
    SAFEARRAY*          safeArray;
    IDispatch*          pDispatch;
    CComPtr<IUnknown>   unknown;
    CStream             stream( pStream );

    ValidatePointer( variant );

    // If it's an array write out individual value.
    if ( V_ISARRAY( variant ) )
    {
        // Get the array.
        if ( V_ISBYREF( variant ) )
            safeArray = *variant->pparray;
        else
            safeArray = variant->parray;
    
        WriteSafeArrayHeader( safeArray, stream );
        
        //write out the safe array elements:
        WriteSafeArrayElements( (VARTYPE)( VT_TYPEMASK & variant->vt ), safeArray, stream );
    }
    // It's not an array, so write the individual value
    else
    {
        switch ( variant->vt )
        {
        case VT_EMPTY:
        case VT_NULL:
            break;
        
        case VT_BOOL:
            // A VARIANT_BOOL is 16 bits.
            stream.Write( V_BOOL( variant ) );
            break;
        
        case VT_UI1:
            stream.Write( V_UI1( variant ) );
            break;
        
        case VT_I2:
            stream.Write( V_I2( variant ) );
            break;
        
        case VT_I4:
            stream.Write( V_I4( variant ) );
            break;
        
        case VT_CY:
            stream.Write( variant->cyVal.Lo );
            stream.Write( variant->cyVal.Hi );
            break;
        
        case VT_R4:
            stream.Write( V_R4( variant ) );
            break;
        
        case VT_R8:
            stream.Write( V_R8( variant ) );
            break;
        
        case VT_DATE:
            // A Variant DATE is a double.
            stream.Write( V_DATE( variant ) );
            break;
        
        case VT_BSTR:
            stream.Write( V_BSTR( variant ) );
            break;
        
        case VT_ERROR:
            stream.Write( V_ERROR( variant ) );
            break;
        
        case VT_DISPATCH:
            pDispatch = V_DISPATCH( variant );
            if ( pDispatch )
                CheckResult( pDispatch->QueryInterface( &unknown ) );

            SaveObjectToStream( unknown, stream );
            break;

        case VT_UNKNOWN:
            SaveObjectToStream( V_UNKNOWN( variant ), stream );
            break;

        default:
            ThrowError( DISP_E_TYPEMISMATCH );
        }
    }

} // WriteDataToStream


//------------------------------------------------------------------------------
// WriteToStream
// Writes the given variant to the stream.
// First writes out the data type of the variant followed by the 
// variant's data.
//------------------------------------------------------------------------------

inline void WriteToStream( const VARIANT* variantParam, IStream* pStream )
{
    CComVariant     variantCopy;
    const VARIANT*  variant;
    CStream         stream( pStream );

    ValidatePointer( variantParam );

    // Use dereferenced copy if incoming is byref.
    if ( V_ISBYREF( variantParam ) )
    {
        CheckResult( VariantCopyInd( &variantCopy, (VARIANT*) variantParam ) );
        variant = &variantCopy;
    }
    else
    {
        variant = variantParam;
    }

    // Write the VT type.
    stream.Write( variant->vt );

    // Write out the actual data.
    WriteDataToStream( variant, pStream );

} // WriteToStream


//------------------------------------------------------------------------------
// ReadDataFromStream
// Given the variant's data type, reads the variant's data from the stream.
//------------------------------------------------------------------------------

inline void ReadDataFromStream( VARTYPE vt, IStream* pStream, VARIANT& variant )
{
    CComPtr<IUnknown>       unknown;
    CComPtr<IDispatch>      dispatch;
    CStream                 stream( pStream );

    // If it's a blob, then read it in as an array of bytes, otherwise read
    // individual value.
    if ( vt & VT_ARRAY )
    {
        VARTYPE elementVT = (VARTYPE) ( VT_TYPEMASK & vt );

        // Read the variant header.
        ReadSafeArrayHeader( &variant, elementVT, stream );
    
        // Read the elements from the stream.
        ReadSafeArrayElements( elementVT, variant.parray, stream );
    }
    // It's not an array, so read the individual value
    else
    {
        switch ( vt )
        {
        case VT_EMPTY:
        case VT_NULL:
            break;
        
        case VT_BOOL:
            // A VARIANT_BOOL is 16 bits.
            stream.Read( V_BOOL( &variant ) );
            break;
        
        case VT_UI1:
            stream.Read( V_UI1( &variant ) );
            break;
        
        case VT_I2:
            stream.Read( V_I2( &variant ) );
            break;
        
        case VT_I4:
            stream.Read( V_I4( &variant ) );
            break;
        
        case VT_CY:
            stream.Read( variant.cyVal.Lo );
            stream.Read( variant.cyVal.Hi );
            break;
        
        case VT_R4:
            stream.Read( V_R4( &variant ) );
            break;
        
        case VT_R8:
            stream.Read( V_R8( &variant ) );
            break;
        
        case VT_DATE:
            stream.Read( V_DATE( &variant ) );
            break;
        
        case VT_BSTR:
            stream.Read( &variant.bstrVal );
            break;
        
        case VT_ERROR:
            stream.Read( V_ERROR( &variant ) );
            break;
        
        case VT_DISPATCH:
            CheckResult( OleLoadFromStream( pStream, IID_IUnknown, (void**)(IUnknown*)&unknown ) );
            CheckResult( unknown->QueryInterface( &dispatch ) );
            V_DISPATCH( &variant ) = dispatch.Detach();
            break;

        case VT_UNKNOWN:
            CheckResult( OleLoadFromStream( pStream, IID_IUnknown, (void**)(IUnknown*)&unknown ) );
            V_UNKNOWN( &variant ) = unknown.Detach();
            break;
        
        default:
            ThrowError( DISP_E_TYPEMISMATCH );
        }
    
        variant.vt = vt;
    }

} // ReadDataFromStream


//------------------------------------------------------------------------------
// ReadFromStream
// Reads the variant from the stream.
// First reads the variant's data type and then calls ReadDataFromStream
// to read the variant's data.
//------------------------------------------------------------------------------

inline void ReadFromStream( IStream* pStream, VARIANT& variant )
{
    VARTYPE             vt;
    CStream             stream( pStream );

    // Read the VT type.
    stream.Read( vt );

    ReadDataFromStream( vt, pStream, variant );

} // ReadFromStream


} // namespace VariantStreaming


//------------------------------------------------------------------------------
// WriteVariantToStream
// Writes the given variant to the stream.
//------------------------------------------------------------------------------

inline void WriteVariantToStream( const VARIANT* variant, IStream* pStream )
{
    // Write the version number of this class.
    CStream( pStream ).Write( VariantStreaming::variantVersion );

    // Call the main routine to write a variant from the stream.
    VariantStreaming::WriteToStream( variant, pStream );

} // WriteVariantToStream


//------------------------------------------------------------------------------
// ReadVariantFromStream
// The passed in variant should be initialized.
//------------------------------------------------------------------------------

inline void ReadVariantFromStream( IStream* pStream, VARIANT& variant )
{
    long    version;

    // Read the version.  If the version is later needed by the reading
    // code, it can be passed as parameter to ReadFromStream, and
    // internally in ReadSafeArrayElements.
    CStream( pStream ).Read( version );

    // Call the main routine to read a variant from the stream.
    VariantStreaming::ReadFromStream( pStream, variant );

} // ReadVariantFromStream


//------------------------------------------------------------------------------
// WriteVariantToBlob
// Streams out a variant to a BLOB.
// The returned BLOB data structure is owned by the caller and should be
// freed using CoTaskMemFree
//------------------------------------------------------------------------------

inline void WriteVariantToBlob( const VARIANT& v, BLOB& blob )
{
    CComPtr<IStream>    stream;
    
    // Create an IStream object that stores data in memory.
    CheckResult( ::CreateStreamOnHGlobal( NULL, TRUE, &stream ) );

    // Stream out the input variant into the memory stream
    ::WriteVariantToStream( &v, stream );

    // Convert the stream to task memory.
    StreamToTaskMemory( stream, blob );

} // WriteVariantToBlob


//------------------------------------------------------------------------------
// ReadVariantFromBlob
// Given a BLOB, streams out a variant.  The caller owns the data in the
// blob parameter.
//------------------------------------------------------------------------------

inline void ReadVariantFromBlob( const BLOB& blob, VARIANT& v )
{
    CComPtr<IStream>    stream;

    // Converts the given BLOB to a stream.
    BlobToStream( blob, &stream );

    // Convert the stream into the output variant.
    ReadVariantFromStream( stream, v );

} // ReadVariantFromBlob
