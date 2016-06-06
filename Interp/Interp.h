/* This file is part of INTERP.

   Program and Documentation are Copyright 2013  by Charles Alan Ford.
   All Rights Reserved.
This work is licensed under a Creative Commons Attribution-NonCommercial 3.0 Unported License
*/

// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the INTERP_EXPORTS
// symbol defined on the command line. This symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// INTERP_API functions as being imported from a DLL, whereas this DLL sees symbols
// defined with this macro as being exported.
#ifdef INTERP_EXPORTS
#define INTERP_API __declspec(dllexport)
#else
#define INTERP_API __declspec(dllimport)
#endif

INTERP_API int xlAutoOpen(void);

INTERP_API LPXLOPER12 // returns a y values interpolated from x and y vectors and a specified x value 
INTERP(LPXLOPER12 x_vector // x vector
	 , LPXLOPER12 y_vector // y vector
	 , double x // specified x value
	  );
