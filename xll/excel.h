// excel.h - Wrapper for Excel4 or Excel12
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#pragma once
//#include <stdexcept>
#include "xll/loper.h"

#define ExcelX Excel<XLOPERX>

namespace xll {

	static int throw_mask = ~xlretSuccess;

	inline void throw_if(int e)
	{
		if (e&throw_mask) {
			switch (e) {
				case xlretSuccess:        
					throw std::runtime_error("Excel returned success.");
				case xlretAbort:          
					throw std::runtime_error("Excel returned macro halted.");
				case xlretInvXlfn:        
					throw std::runtime_error("Excel returned invalid function number."); 
				case xlretInvCount:       
					throw std::runtime_error("Excel returned invalid number of arguments."); 
				case xlretInvXloper:      
					throw std::runtime_error("Excel returned invalid OPER structure.");  
				case xlretStackOvfl:      
					throw std::runtime_error("Excel returned stack overflow.");  
				case xlretFailed:         
					throw std::runtime_error("Excel returned command failed.");  
				case xlretUncalced:       
					throw std::runtime_error("Excel returned uncalced cell.");
				case xlretNotThreadSafe:  
					throw std::runtime_error("Excel returned not allowed during multi-threaded calc.");
				case xlretInvAsynchronousContext:					  
					throw std::runtime_error("Excel returned invalid asynchronous function handle.");
				case xlretNotClusterSafe: 
					throw std::runtime_error("Excel returned not supported on cluster.");
				default:                  
					throw std::runtime_error("Excel returned an unknown error! This should never happen.");
			}
		}
	}

	template<class X>
	LXOPER<X> Excel(int f)
	{
		X x;

		throw_if (traits<X>::Excel(f, &x, 0));

		return x; // LOPER constructor takes ownership
	}
	template<class X>
	LXOPER<X> Excel(int f, const X& x0)
	{
		X x;

		throw_if (traits<X>::Excel(f, &x, 1, &x0));

		return x;
	}
	template<class X>
	LXOPER<X> Excel(int f, const X& x0, const X& x1)
	{
		X x;

		throw_if (traits<X>::Excel(f, &x, 2, &x0, &x1));

		return x;
	}
	template<class X>
	LXOPER<X> Excel(int f, const X& x0, const X& x1, const X& x2)
	{
		X x;

		throw_if (traits<X>::Excel(f, &x, 3, &x0, &x1, &x2));

		return x;
	}
	template<class X>
	LXOPER<X> Excel(int f, const X& x0, const X& x1, const X& x2, const X& x3)
	{
		X x;

		throw_if (traits<X>::Excel(f, &x, 4, &x0, &x1, &x2, &x3));

		return x;
	}
	template<class X>
	LXOPER<X> Excel(int f, const X& x0, const X& x1, const X& x2, const X& x3,
		const X& x4)
	{
		X x;

		throw_if (traits<X>::Excel(f, &x, 5, &x0, &x1, &x2, &x3,
			&x4));

		return x;
	}
	template<class X>
	LXOPER<X> Excel(int f, const X& x0, const X& x1, const X& x2, const X& x3,
		const X& x4, const X& x5)
	{
		X x;

		throw_if (traits<X>::Excel(f, &x, 6, &x0, &x1, &x2, &x3,
			&x4, &x5));

		return x;
	}
	template<class X>
	LXOPER<X> Excel(int f, const X& x0, const X& x1, const X& x2, const X& x3,
		const X& x4, const X& x5, const X& x6)
	{
		X x;

		throw_if (traits<X>::Excel(f, &x, 7, &x0, &x1, &x2, &x3,
			&x4, &x5, &x6));

		return x;
	}
	template<class X>
	LXOPER<X> Excel(int f, const X& x0, const X& x1, const X& x2, const X& x3,
		const X& x4, const X& x5, const X& x6, const X& x7)
	{
		X x;

		throw_if (traits<X>::Excel(f, &x, 7, &x0, &x1, &x2, &x3,
			&x4, &x5, &x6, &x7));

		return x;
	}
	template<class X>
	LXOPER<X> Excel(int f, const X& x0, const X& x1, const X& x2, const X& x3,
		const X& x4, const X& x5, const X& x6, const X& x7, const X& x8)
	{
		X x;

		throw_if (traits<X>::Excel(f, &x, 7, &x0, &x1, &x2, &x3,
			&x4, &x5, &x6, &x7, &x8));

		return x;
	}
	template<class X>
	LXOPER<X> Excel(int f, const X& x0, const X& x1, const X& x2, const X& x3,
		const X& x4, const X& x5, const X& x6, const X& x7, const X& x8, const X& x9)
	{
		X x;

		throw_if (traits<X>::Excel(f, &x, 7, &x0, &x1, &x2, &x3,
			&x4, &x5, &x6, &x7, &x8, &x9));

		return x;
	}
	template<class X>
	LXOPER<X> Excel(int f, const X& x0, const X& x1, const X& x2, const X& x3,
		const X& x4, const X& x5, const X& x6, const X& x7, const X& x8, const X& x9,
		const X& x10)
	{
		X x;

		throw_if (traits<X>::Excel(f, &x, 7, &x0, &x1, &x2, &x3,
			&x4, &x5, &x6, &x7, &x8, &x9, &x10));

		return x;
	}

}

