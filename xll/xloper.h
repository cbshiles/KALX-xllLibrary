// xloper.h - Helper functions for XLOPER's.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include <iostream>
#include <utility>
#include "xll/excel.h"

namespace xll {

	template<class X>
	inline typename traits<X>::xword rows(const X& x)
	{
		// xltypeSRef, xltypeRef, xltypeBigData?
		return x.xltype == xltypeMulti ? x.val.array.rows : 1;
	}
	template<class X>
	inline typename traits<X>::xword columns(const X& x)
	{
		// xltypeSRef, xltypeRef, xltypeBigData?
		return x.xltype == xltypeMulti ? x.val.array.columns : 1;
	}
	template<class X>
	inline size_t size(const X& x)
	{
		return rows(x) * columns(x);		
	}

	template<class X>
	inline typename X& index(X& x, size_t i)
	{
		ensure (i < size(x));

		if (x.xltype == xltypeMulti)
			return x.val.array.lparray[i];
		else
			return x;
	}
	template<class X>
	inline typename const X& index(const X& x, size_t i)
	{
		ensure (i < size(x));

		if (x.xltype == xltypeMulti)
			return x.val.array.lparray[i];
		else
			return x;
	}

	template<class X>
	inline typename X& index(X& x, size_t i, size_t j)
	{
		ensure (i < rows(x));
		ensure (j < columns(x));

		if (x.xltype == xltypeMulti)
			return x.val.array.lparray[i*x.val.array.columns + j];
		else
			return x;
	}
	template<class X>
	inline typename const X& index(const X& x, size_t i, size_t j)
	{
		ensure (i < rows(x));
		ensure (j < columns(x));

		if (x.xltype == xltypeMulti)
			return x.val.array.lparray[i*x.val.array.columns + j];
		else
			return x;
	}

	template<class X>
	inline typename X* begin(X& x)
	{
		return x.xltype == xltypeMulti ? x.val.array.lparray : &x;
	}
	template<class X>
	inline const typename X* begin(const X& x)
	{
		return x.xltype == xltypeMulti ? x.val.array.lparray : &x;
	}
	template<class X>
	inline typename X* end(X& x)
	{
		return x.xltype == xltypeMulti ? x.val.array.lparray + size(x) : &x + 1;
	}
	template<class X>
	inline const typename X* end(const X& x)
	{
		return x.xltype == xltypeMulti ? x.val.array.lparray + size(x) : &x + 1;
	}

	template<class X>
	inline double to_number(const X& x)
	{
		switch (x.xltype) {
		case xltypeNum:
			return x.val.num;
		case xltypeStr:
			return x.val.str[0]; // so if(x) ... works
		case xltypeBool:
			return x.val.xbool;
		case xltypeRef: xltypeSRef;
			return 1; // what else???
		case xltypeErr:
			return 0; // so ensure(Err()) fails
		case xltypeMulti: // similar to str
			return x.val.array.rows * x.val.array.columns;
		case xltypeMissing: case xltypeNil:
			return 0;
		case xltypeInt:
			return x.val.w;
		}

		return std::numeric_limits<double>::quiet_NaN();
	}
	template<class X>
	inline typename traits<X>::xstring 
	to_string(const X& x, 
		typename traits<X>::xcstr fs = traits<X>::null(), 
		typename traits<X>::xcstr rs = traits<X>::null())
	{
		traits<X>::xstring s;

		if (x.xltype == xltypeStr) {
			s = traits<X>::xstring(x.val.str + 1, x.val.str[0]);
		}
		else if (x.xltype == xltypeMulti) {
			s = to_string(index(x, 0));
			for (size_t i = 1; i < size(x); ++i) {
				if (rs && i % columns(x) == 0)
					s += rs;
				else if (fs)
					s += fs;
				s += to_string(index(x, i));
			}
		}
		else {
			X y = Excel<X>(xlfText, x, traits<X>::format());

			size_t size;
#ifdef EXCEL12
			// workarounds for first char not strlen
			if (x.xltype != xltypeNum) {
				size = traits<X>::strlen(y.val.str+1);
			}
			else {
				X n = Excel<X>(xlfLen, y);
				ensure (n.xltype == xltypeNum);
				size = static_cast<size_t>(n.val.num);
			}
#else
			size = y.val.str[0];
#endif
			s = traits<X>::xstring(y.val.str + 1, size); 
		}

		return s;
	}

	template<class X>
	inline bool operator_equals(const X& x, const X& y)
	{
		if (x.xltype != y.xltype)
			return false;

		switch (x.xltype) {
			case xltypeNum:
				return x.val.num == y.val.num;
			case xltypeStr:
				return x.val.str[0] == y.val.str[0] && 0 == xll::traits<X>::strnicmp(x.val.str + 1, y.val.str + 1, x.val.str[0]);
			case xltypeBool:
				return x.val.xbool == y.val.xbool;
			case xltypeRef:
				return true; //!!!fix this
			case xltypeErr:
				return x.val.err == y.val.err;
			case xltypeMulti:
				if (rows(x) != rows(y))
					return false;
				if (columns(x) != columns(y))
					return false;
				for (traits<X>::xword i = 0; i < size(x); ++i)
					if (!xll::operator_equals(index(x, i), index(y, i)))
						return false;
				return true;
			case xltypeMissing:
				return y.xltype == xltypeMissing;
			case xltypeNil:
				return y.xltype == xltypeNil;
			case xltypeSRef:
				return true; //!!!fix this
			case xltypeInt:
				return x.val.w == y.val.w;
		}

		return false;
	}

	template<class X>
	inline bool operator_less(const X& x, const X& y)
	{
		if (x.xltype < y.xltype)
			return true;

		switch (x.xltype) {
			case xltypeNum:
				return x.val.num < y.val.num;
			case xltypeStr:
				return 0 > xll::traits<X>::strnicmp(x.val.str + 1, y.val.str + 1, __min(x.val.str[0], y.val.str[0]))
					|| ((x.val.str[0] < y.val.str[0]) && 0 == xll::traits<X>::strnicmp(x.val.str + 1, y.val.str + 1, x.val.str[0]));
			case xltypeBool:
				return x.val.xbool < y.val.xbool;
			case xltypeRef:
				return true; //!!!fix this
			case xltypeErr:
				return x.val.err < y.val.err;
			case xltypeMulti:
				if (rows(x) >= rows(y))
					return false;
				if (columns(x) >= columns(y))
					return false;
				for (traits<X>::xword i = 0; i < rows(x); ++i)
					for (traits<X>::xword j = 0; j < columns(x); ++j)
						if (!operator_less(index(x, i, j), index(y, i, j)))
							return false;
				return true;
			case xltypeMissing:
				return false;
			case xltypeNil:
				return false;
			case xltypeSRef:
				return false; //!!!fix this
			case xltypeInt:
				return x.val.w < y.val.w;
		}

		return false;
	}} // namespace xll

using namespace std::rel_ops;

// These go in the global namespace.
inline bool operator==(const XLOPER12& x, const XLOPER12& y)
{
	return xll::operator_equals<XLOPER12>(x, y);
}
inline bool operator<(const XLOPER12& x, const XLOPER12& y)
{
	return xll::operator_less<XLOPER12>(x, y);
}
inline bool operator==(const XLOPER& x, const XLOPER& y)
{
	return xll::operator_equals<XLOPER>(x, y);
}
inline bool operator<(const XLOPER& x, const XLOPER& y)
{
	return xll::operator_less<XLOPER>(x, y);
}

inline
std::ostream& operator<<(std::ostream& os, const XLOPER& x)
{
	os << xll::to_string<XLOPER>(x).c_str();

	return os;
}

inline
std::ostream& operator<<(std::ostream& os, const XLOPER12& x)
{
	os << xll::to_string<XLOPER12>(x).c_str();

	return os;
}
