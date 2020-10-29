The only file you need from here is VariantStream.h.  Everything else is just
used for testing.

See the top of VariantStream.h for usage details.



Variant streaming code
o	Uses any given IStream to stream the Variant into and out of. 
o	Data is streamed in efficient binary form. 
o	Stream is versioned for backwards compatibility. 
o	Supports arbitrary size and arbitrary dimension safe-arrays. 
o	Object streaming is supported if the object in variant supports IPersistStream[Init]. 
o	All code is in one header file (VariantStream.h) and only two routines are exposed: WriteVariantToStream and ReadVariantFromStream. 
o	Comes with supporting test code that tests the header file -- in case code is modified 
o	Does not use C++ exception handling.  Test project has EH flag turned off
o	Doesn't use the CRT. 
o	Does not use any Direct-To-COM (VC++'s comdef.h, such as _variant_t, _bstr_t, _com_ptr, _com_error) 
o	Works in both Unicode and ANSI 
