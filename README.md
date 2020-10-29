# Overview
Variant streaming code

The only file you need from here is VariantStream.h.  Everything else is just
used for testing.

See the top of VariantStream.h for usage details.


This code provides two global functions, WriteVariantToStream and ReadVariantFromStream, that enable you to read and write a variant to a stream. In addition, there are two other global functions, ReadVariantFromBlob and WriteVariantToBlob, for reading and writing a variant to a BLOB.

*	Uses any given IStream to stream the Variant into and out of. 
*	Data is streamed in efficient binary form. 
*	Stream is versioned for backwards compatibility. 
*	Supports arbitrary size and arbitrary dimension safe-arrays. 
*	Object streaming is supported if the object in variant supports IPersistStream[Init]. 
*	All code is in one header file (VariantStream.h) and only two routines are exposed: WriteVariantToStream and ReadVariantFromStream. 
*	Comes with supporting test code that tests the header file -- in case code is modified 
*	Does not use C++ exception handling.  Test project has EH flag turned off
*	Doesn't use the CRT. 
*	Does not use any Direct-To-COM (VC++'s comdef.h, such as _variant_t, _bstr_t, _com_ptr, _com_error) 
*	Works in both Unicode and ANSI 

(Originally published http://www.codeguru.com/cpp/com-tech/activex/com/article.php/c2669/Variant-Streaming-Code.htm)
