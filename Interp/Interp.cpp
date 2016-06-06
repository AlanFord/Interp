// Interp.cpp : Defines the exported functions for the DLL application.
//
/* This file is part of INTERP.

   Program and Documentation are Copyright 2013  by Charles Alan Ford.
   All Rights Reserved.
   This work is licensed under a Creative Commons Attribution-NonCommercial 3.0 Unported License.
*/

#include "stdafx.h"
#include <xlcall.h>
#include <math.h>
#include <string>
#include "framewrk.h"
#include "interp.h"
// Interp.cpp : Defines the exported functions for the DLL application.
//

/*
** rgFuncs
**
** This is a table of all the functions exported by this module.
** These functions are all registered (in xlAutoOpen) when you
** open the XLL. Before every string, leave a space for the
** byte count. The format of this table is the same as
** the last seven arguments to the REGISTER function.
** rgFuncsRows define the number of rows in the table. The
** dimension [3] represents the number of columns in the table.
*/
#define rgFuncsRows 1
#define rgFuncsCols 12

static LPWSTR rgFuncs[rgFuncsRows][rgFuncsCols] = {
	{L"INTERP",					L"UUUB",  L"INTERP",  L"X Vector, Y Vector, X Value", L"1", L"Dominion", L"", L"", L"Returns a linearly interpolated value from the two specified arrays", L"X array", L"Y array", L"X value"}
};

/*
** xlAutoOpen
**
** xlAutoOpen is how Microsoft Excel loads XLL files.
** When you open an XLL, Microsoft Excel calls the xlAutoOpen
** function, and nothing more.
**
** More specifically, xlAutoOpen is called by Microsoft Excel:
**
**  - when you open this XLL file from the File menu,
**  - when this XLL is in the XLSTART directory, and is
**		automatically opened when Microsoft Excel starts,
**  - when Microsoft Excel opens this XLL for any other reason, or
**  - when a macro calls REGISTER(), with only one argument, which is the
**		name of this XLL.
**
** xlAutoOpen is also called by the Add-in Manager when you add this XLL
** as an add-in. The Add-in Manager first calls xlAutoAdd, then calls
** REGISTER("EXAMPLE.XLL"), which in turn calls xlAutoOpen.
**
** xlAutoOpen should:
**
**  - register all the functions you want to make available while this
**		XLL is open,
**
**  - add any menus or menu items that this XLL supports,
**
**  - perform any other initialization you need, and
**
**  - return 1 if successful, or return 0 if your XLL cannot be opened.
*/
INTERP_API int xlAutoOpen(void)
{

	static XLOPER12 xDLL;	/* name of this DLL */
	int i;					/* Loop index */

	/*
	** In the following block of code the name of the XLL is obtained by
	** calling xlGetName. This name is used as the first argument to the
	** REGISTER function to specify the name of the XLL. Next, the XLL loops
	** through the rgFuncs[] table, registering each function in the table using
	** xlfRegister. Functions must be registered before you can add a menu
	** item.
	*/

	Excel12f(xlGetName, &xDLL, 0);

        for (i=0;i<rgFuncsRows;i++) 
		{
			Excel12f(xlfRegister, 0, 1 + rgFuncsCols,
				(LPXLOPER12)&xDLL,
				(LPXLOPER12)TempStr12(rgFuncs[i][0]),
				(LPXLOPER12)TempStr12(rgFuncs[i][1]),
				(LPXLOPER12)TempStr12(rgFuncs[i][2]),
				(LPXLOPER12)TempStr12(rgFuncs[i][3]),
				(LPXLOPER12)TempStr12(rgFuncs[i][4]),
				(LPXLOPER12)TempStr12(rgFuncs[i][5]),
				(LPXLOPER12)TempStr12(rgFuncs[i][6]),
				(LPXLOPER12)TempStr12(rgFuncs[i][7]),
				(LPXLOPER12)TempStr12(rgFuncs[i][8]),
				(LPXLOPER12)TempStr12(rgFuncs[i][9]),
				(LPXLOPER12)TempStr12(rgFuncs[i][10]),
				(LPXLOPER12)TempStr12(rgFuncs[i][11]));
		}

	/* Free the XLL filename */
	Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDLL);

	return 1;
}



// ClipSize is a utility function that will determine the size of a "multi" array
// structure.  It checks to see if the data is organized in columns or rows (giving
// preference to columns), and ignores empty cells at the end of the array.
// It returns the size of the 1D table of valid data.

WORD ClipSize( XLOPER12 *multi )
{
	WORD		size, i;
	LPXLOPER12	ptr;
	
	// get the number of columns in the data
	size = multi->val.array.columns;
	
	// if there's only one column, then it must be organized in multiple rows.
	if ( size == 1 )
		size = multi->val.array.rows;

	// ignore empty or error values at the end of the array.
	for ( i = size - 1; i >= 0; i-- )
	{
		ptr = multi->val.array.lparray + i;

		if ( ptr->xltype != xltypeNil )
			break;
	}
	
	return i + 1;
}

INTERP_API LPXLOPER12 // returns interpolated value 
INTERP(LPXLOPER12 x_vector // x vector
	 , LPXLOPER12 y_vector // y vector
	 , double currX // specified x value
	  )
{
	short           extrapFlag = 0;					// placeholder for future feature!
    double			xlo, xhi,						// temporary values for interpolation
    				ylo, yhi,
    				h;								// the current x value
	static XLOPER12 xResult;
    static XLOPER12	xMulti,						// x_vector coerced to xltypeMulti
					yMulti,						// y_vector coerced to xltypeMulti
					tempTypeMulti;				// xltypeMulti in an XLOPER12 for passing

	short			hasXMulti = 0,
					hasYMulti = 0,
     				sortFlag = -1;		// 0 for increasing x values, 1 for decreasing

   	WORD			xSize,				// size of the preferred dimension of xArray
    				ySize;				// size of the preferred dimension of yArray

	ULONG			lo, hi, mid,
    				i, j,				// temporary index values
    				xCount = 0;			// the number of x target values to handle

    LPXLOPER12		xPtr, yPtr;					// pointers to arrays of x & y table data

	try {
		// Initialize some variables
		tempTypeMulti.xltype = xltypeInt;
		tempTypeMulti.val.w = xltypeMulti;
		// ======= Get the xArray Data ==============
		// and verify the type

		if (x_vector->xltype != xltypeRef && x_vector->xltype != xltypeSRef && x_vector->xltype != xltypeMulti )
			throw "invalid x vector";

		// Coerce the data into the "Multi" type since that's what we expect.
		// If coerce fails due to an uncalced cell, return immediately and Excel will
		// call us again in a moment after that cell has been calced.
		if (xlretUncalced == Excel12f( xlCoerce, (LPXLOPER12) &xMulti, 2, (LPXLOPER12) x_vector, (LPXLOPER12) &tempTypeMulti ) )
		{
			return 0;
		}    
		hasXMulti = 1;		// indicate that Excel has allocated memory for the xMulti

		// ======= Get the yArray Data ==============

		if (y_vector->xltype != xltypeRef && y_vector->xltype != xltypeSRef && y_vector->xltype != xltypeMulti )
			throw "invalid y vector";

		// Coerce the data into the "Multi" type since that's what we expect.
		// If coerce fails due to an uncalced cell, return immediately and Excel will
		// call us again in a moment after that cell has been calced.
		if ( xlretUncalced == Excel12f( xlCoerce, (LPXLOPER12) &yMulti, 2, (LPXLOPER12) y_vector, (LPXLOPER12) &tempTypeMulti ) )
		{
			// if coerce failed due to an uncalced cell, return immediately.
			// first need to free memory in xMulti
			if ( hasXMulti )
				Excel12f( xlFree, 0, 1, (LPXLOPER12) &xMulti );
			return 0;
		}
		hasYMulti = 1;	// indicate Excel has allocated memory in the yMulti structure

		// determine the size of the x and y tables, ignoring empty cells
		// at the end of the data
		xSize = ClipSize( &xMulti );
		ySize = ClipSize( &yMulti );
    
		// save some temporary pointers to the actual x and y table data
 		xPtr = xMulti.val.array.lparray;
 		yPtr = yMulti.val.array.lparray;
    
		// use the smaller array dimension from x or y
		// from here on out, xSize is the dimension to use
		if ( ySize < xSize )
			xSize = ySize;
	
		// make sure we have at least two values, otherwise there's no table
		// for interpolation
		if ( xSize < 2 )
			throw "invalid x vector size";

		// verify that the entire xArray and yArray are nums
		// also verify that xArray is monotonically increasing or decreasing
		for ( i=0; i<xSize; i++ )
		{
 			if (	xPtr[i].xltype != xltypeNum	|| yPtr[i].xltype != xltypeNum )
				throw "Input data aren't numbers";
 		
 			// make sure that the x table data is monotonically increasing or decreasing
 			// sortflag is set to zero for increasing, and one for decreasing data
 			if ( i > 0 )
 			{
 				// is the current value less than the previous one?
 				if ( xPtr[i].val.num < xPtr[i-1].val.num )
 				{
 					if ( sortFlag == 0 )	// if previous data was increasing, it's an error
						throw "x vector isn't sorted";
					sortFlag = 1;			// indicate decreasing data
 				}
  				// is the current value greater than the previous one?
				else if ( xPtr[i].val.num > xPtr[i-1].val.num )
 				{
 					if ( sortFlag == 1 )	// if previous data was decreasing, it's an error
						throw "x vector isn't sorted";
					sortFlag = 0;			// indicate increasing data
 				}
 			}
		}
		// at this point, we have valid data for xArray and yArray,
		// so begin the interpolation (finally)
			
		// see if x is less than the x table minimum
		if ( ( sortFlag == 0 && currX < xPtr[0].val.num ) ||
			 ( sortFlag == 1 && currX > xPtr[0].val.num ) )
		{
			if ( ! extrapFlag )
			{
				// if extrapolation is not allowed, throw an error
				throw "Extrapolation not permitted";
			}
			else
			{
				// otherwise, use the first two entries to extrapolate
				lo = 0;
				hi = 1;
			}
		}
		// see if x is greater than the x table maximum
		else if ( ( sortFlag == 0 && currX > xPtr[xSize-1].val.num ) ||
			  	  ( sortFlag == 1 && currX < xPtr[xSize-1].val.num ) )
		{
			if ( ! extrapFlag )
			{
				// if extrapolation is not allowed, just return the last y table entry
				throw "Extrapolation not permitted";
			}
			else
			{
				// otherwise, use the last two entries to extrapolate
				lo = xSize - 2;
				hi = xSize - 1;
			}
		}
		else
		{
			// if x is within the bounds of the x table, then do a binary search
			// in the x table to find table entries that bound the x value
			lo = 0;
			hi = xSize - 1;
		    
			// limit to 1000 loops to avoid an infinite loop problem
			for ( j=0; j<1000 && hi > lo + 1; j++ )
			{
				mid = ( hi + lo ) / 2;
				if ( ( sortFlag == 0 && currX > xPtr[mid].val.num ) ||
					( sortFlag == 1 && currX < xPtr[mid].val.num ) )
					lo = mid;
				else
					hi = mid;
			}

			// if we exceeded the max # of loops, just return a #value! error
			// for the current cell
			if ( j >= 1000 )
			{
				throw "not converging";
			}
		}

		// get the bounding x table values and y table values for interpolation
		xlo = xPtr[lo].val.num;
		xhi = xPtr[hi].val.num;
		ylo = yPtr[lo].val.num;
		yhi = yPtr[hi].val.num;
			
		// do the interpolation
    	h = ( currX - xlo ) / ( xhi - xlo ) * ( yhi - ylo ) + ylo;

		// return
		xResult.xltype = xltypeNum;
		xResult.val.num = h;
	}
	catch(...) {
		xResult.xltype = xltypeErr;
		xResult.val.err = xlerrValue;
	}
	// free the memory allocated by Excel on our behalf	
	if ( hasXMulti )
		Excel12f( xlFree, 0, 1, (LPXLOPER12) &xMulti );
	if ( hasYMulti )
		Excel12f( xlFree, 0, 1, (LPXLOPER12) &yMulti );
	return (LPXLOPER12) &xResult;
}


