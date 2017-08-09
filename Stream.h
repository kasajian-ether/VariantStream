#pragma once

//==============================================================================
// Included code:
//  CStream - A wrapper for IStream.
//  StreamToTaskMemory -- Converts a stream to a blob
//  BlobToStream -- Converts a blob to a stream.
//==============================================================================

#ifndef HR
#define HR(_ex) { HRESULT _hr = _ex; if( FAILED(_hr) ) return _hr; }
#endif
#define ThrowError(_ex) RaiseException( 0xE0000001, 0, 0, 0);
#define CheckResult(_ex) do { HRESULT _hr = _ex; if( FAILED(_hr) ) ThrowError(_hr) } while (0,0)
#define VerifyAllocation(_ex) do { if ( NULL == (_ex) ) ThrowError( E_OUTOFMEMORY ); } while(0, 0)
#define ValidatePointer(_ex) do { if ( NULL == (_ex) ) ThrowError( E_POINTER ); } while( 0, 0 )


//==============================================================================
// Wrapper for IStream.  Supplies template Read and Write method so that 
// data sizes are automatically calculated.  Standard overload assumes
// an intrinsic type and uses the & operator on the parameter to to get to the data
// to be read or written and the sizeof operator on the parameter to determine
// the size of the data to be read or written.
// A BSTR operator is supplied for reading and writing strings.
// Example:
//        IStream*    pStream = ...;
//        ...
//        CStream     stream( pStream );
//    
//        stream.Write( longVal );
//        stream.Write( shortVal );
//        stream.Write( floatVal );
//        stream.Write( bstrVal );
//        stream.Write( (BSTR)CComBSTR( lpszStringVal ) );
//    
// ToDo:
//   Could add support for LPCSTR, LPWSTR, string and wstring.
//==============================================================================

class CStream
{
public:
    //------------------------------------------------------------------------------
    // Construct using an IStream pointer
    //------------------------------------------------------------------------------
    
    inline CStream( IStream* stream )
        :   stream( stream )
    {
    }
    
    //------------------------------------------------------------------------------
    // Alows code to get the IStream pointer.
    //------------------------------------------------------------------------------
    
    inline operator IStream*()
    {
        return stream;
    }
    
    
    //------------------------------------------------------------------------------
    // Template method that will write any data type that supports & to return
    // the address of the data and sizeof that will return the size of the data.
    //------------------------------------------------------------------------------
    
    template <class Q>
    inline void Write( Q value )
    {
        ULONG   written;
        CheckResult( stream->Write( &value, sizeof( value ), &written ) );
        if ( written != sizeof( value ) )
            ThrowError( E_FAIL );
    }
    
    
    //------------------------------------------------------------------------------
    // Special code for writing BSTRs
    //------------------------------------------------------------------------------
    
    inline void Write( BSTR value )
    {
        ULONG   written;
        UINT    bufferSize = ::SysStringLen( value ) * sizeof( WCHAR );
        
        Write( bufferSize );
        CheckResult( stream->Write( value, bufferSize, &written ) );
        if ( written != bufferSize )
            ThrowError( E_FAIL );
    }
    
    
    //------------------------------------------------------------------------------
    // Template method that will read any data type that supports & to return
    // the address of the data and sizeof that will return the size of the data.
    //------------------------------------------------------------------------------
    
    template <class Q>
    inline void Read( Q& value )
    {
        ULONG   read;
        CheckResult( stream->Read( &value, sizeof( value ), &read ) );
        if ( read != sizeof( value ) )
            ThrowError( E_FAIL );
    }
    
    
    //------------------------------------------------------------------------------
    // Special code for reading BSTRs
    //------------------------------------------------------------------------------
    
    inline void Read( BSTR* value )
    {
        ULONG   read;
        UINT    bufferSize;
        
        Read( bufferSize );
        LPWSTR buf = (LPWSTR)_alloca( bufferSize + sizeof( WCHAR ) );
        CheckResult( stream->Read( buf, bufferSize, &read ) );
        if ( read != bufferSize )
            ThrowError( E_FAIL );

        UINT len = bufferSize / sizeof( WCHAR );
        buf[len] = L'\0';
        *value = ::SysAllocStringLen( buf, len );
        if ( !*value )
            ThrowError( E_OUTOFMEMORY );
    }


    //------------------------------------------------------------------------------
    // Special code for reading BSTRs
    //------------------------------------------------------------------------------

    DWORD GetSize()
    {
        ULARGE_INTEGER                  largeSize = { 0 };
        ULARGE_INTEGER                  uLargeCurrentLocation;
        LARGE_INTEGER                   largeZero;
        largeZero.QuadPart = 0i64;

        // Get the Size of the stream.
        CheckResult( stream->Seek( largeZero, STREAM_SEEK_CUR, &uLargeCurrentLocation ) );
        CheckResult( stream->Seek( largeZero, STREAM_SEEK_END, &largeSize ) );

        LARGE_INTEGER largeCurrentLocation;
        largeCurrentLocation.QuadPart = uLargeCurrentLocation.QuadPart;
        CheckResult( stream->Seek( largeCurrentLocation, STREAM_SEEK_SET, NULL ) );

        return largeSize.LowPart;
    }


    //------------------------------------------------------------------------------
    // GetStreamHGlobal
    //------------------------------------------------------------------------------

    HGLOBAL GetStreamHGlobal()
    {
        HGLOBAL             handle;

        // Get the memory handle associated with the stream.
        CheckResult( ::GetHGlobalFromStream( stream, &handle ) );

        return handle;
    }


private:
    // If these get accessed, change your code to explicitly cast the parameter to
    // a BSTR for writes and BSTR* for reads.
    void Write( CComBSTR value );
    void Read( CComBSTR* value );
    
private:
    CComPtr<IStream>    stream;
    
}; // class CStream


//------------------------------------------------------------------------------
// StreamToTaskMemory
// Given a memory stream, converts it to a task memory pointer, which the caller
// is responsible for deleting with CoTaskMemFree.
//------------------------------------------------------------------------------

inline void StreamToTaskMemory( IStream* pStream, BLOB& blob )
{
    CStream         stream( pStream );

    // Get the memory handle associated with the stream.
    HGLOBAL handle = stream.GetStreamHGlobal();

    // Allocate new memory that will be returned to the caller as a BLOB.
    blob.cbSize = stream.GetSize();
    blob.pBlobData = (BYTE*) ::CoTaskMemAlloc( blob.cbSize );

    // Copy the memory from the stream to the new memory.
    ::CopyMemory( blob.pBlobData, ::GlobalLock( handle ), blob.cbSize );
    ::GlobalUnlock( handle );

} // StreamToTaskMemory


//------------------------------------------------------------------------------
// BlobToStream
// Given a BLOB object, converts to a memory stream.
// The BLOB may have null data, in which case an empty memory stream is created.
//------------------------------------------------------------------------------

inline void BlobToStream( const BLOB& blob, IStream** ppStream )
{
    HGLOBAL             handle = NULL;

    // Create a handle from the output BLOB
    if ( blob.cbSize )
    {
        handle = ::GlobalAlloc( GMEM_MOVEABLE, blob.cbSize );
        if ( !handle )
            CheckResult( HRESULT_FROM_WIN32( ::GetLastError() ) );

        // Copy the blob to the new memory.
        if ( blob.pBlobData )
        {
            ::memcpy( ::GlobalLock( handle ), blob.pBlobData, blob.cbSize );
            ::GlobalUnlock( handle );
        }
    }

    // Create an IStream object that stores data in memory.
    CheckResult( ::CreateStreamOnHGlobal( handle, TRUE, ppStream ) );

} // BlobToStream
