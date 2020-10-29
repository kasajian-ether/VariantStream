#pragma once


//------------------------------------------------------------------------------
// Constants
//------------------------------------------------------------------------------

const LARGE_INTEGER   largeZero = { 0 };


//------------------------------------------------------------------------------
// Creates an IStream object that owns the data in memory.
//------------------------------------------------------------------------------

inline HRESULT CreateMemoryStream( IStream** ppStream )
{
    return CreateStreamOnHGlobal( NULL, TRUE, ppStream );

} // CreateMemoryStream


//------------------------------------------------------------------------------
// Rewinds the given IStream
//------------------------------------------------------------------------------

inline HRESULT RewindStream( IStream* pStream )
{
   return pStream->Seek( largeZero, STREAM_SEEK_SET, NULL );

} // RewindStream


