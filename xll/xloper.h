// xloper.h - Helper functions for XLOPER's.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include <iostream>
#include <utility>
#include "excel.h"

// simple reference
inline bool operator==(const XLREF& x, const XLREF& y)
{
	return x.rwFirst == y.rwFirst && x.rwLast == y.rwLast && x.colFirst == y.colFirst && x.colLast == y.colLast;
}
inline bool operator==(const XLREF12& x, const XLREF12& y)
{
	return x.rwFirst == y.rwFirst && x.rwLast == y.rwLast && x.colFirst == y.colFirst && x.colLast == y.colLast;
}
inline bool operator<(const XLREF& x, const XLREF& y)
{
	return x.rwFirst  != y.rwFirst  ? x.rwFirst < y.rwFirst
		 : x.rwLast   != y.rwLast   ? x.rwLast < y.rwLast
		 : x.colFirst != y.colFirst ? x.colFirst < y.colFirst
		 : x.colLast  != y.colLast  ? x.colLast < y.colLast
		 : false;
}
inline bool operator<(const XLREF12& x, const XLREF12& y)
{
	return x.rwFirst  != y.rwFirst  ? x.rwFirst < y.rwFirst
		 : x.rwLast   != y.rwLast   ? x.rwLast < y.rwLast
		 : x.colFirst != y.colFirst ? x.colFirst < y.colFirst
		 : x.colLast  != y.colLast  ? x.colLast < y.colLast
		 : false;
}

template<class X>
class XMREF : public X {
	typename xll::traits<X>::xlmref ref; // only one reference allowed
public:
	XMREF()
	{
		xltype = xltypeRef;
		ref.count = 1;
		val.mref.lpmref = &ref;
	}
	XMREF(const X& x)
	{
		if (x.xltype == xltypeRef) {
			xltype = x.xltype;
			val.mref.idSheet = x.val.mref.idSheet;
			ref.count = 1;
			val.mref.lpmref = &ref;
			val.mref.lpmref->reftbl[0] = x.val.mref.lpmref->reftbl[0];
		}
		else {
			xltype = x.xltype;
			val = x.val;
		}
	}
	XMREF operator=(const X& x)
	{
		if (this != &x) {
			if (x.xltype == xltypeRef) {
				xltype = x.xltype;
				val.mref.idSheet = x.val.mref.idSheet;
				ref.count = 1;
				val.mref.lpmref = &ref;
				val.mref.lpmref->reftbl[0] = x.val.mref.lpmref->reftbl[0];
			}
			else {
				xltype = x.xltype;
				val = x.val;
			}
		}

		return *this;
	}
	~XMREF()
	{ }
	typename xll::traits<X>::xlref& operator[](typename xll::traits<X>::xword i)
	{
		ensure (i == 0);
		return val.mref.lpmref->reftbl[0];
	}
	const typename xll::traits<X>::xlref& operator[](typename xll::traits<X>::xword i) const
	{
		ensure (i == 0);
		return val.mref.lpmref->reftbl[0];
	}
};
typedef XMREF<XLOPER>   MREF;
typedef XMREF<XLOPER12> MREF12;
typedef XMREF<XLOPERX>  MREFX;

namespace xll {

	template<class X>
	inline typename traits<X>::xword rows(const X& x)
	{
		// xltypeRef, xltypeBigData?
		return x.xltype == xltypeMulti ? x.val.array.rows 
			: x.xltype == xltypeSRef ? x.val.sref.ref.rwLast - x.val.sref.ref.rwFirst + 1
			: x.xltype == xltypeRef ? x.val.mref.lpmref->reftbl[0].rwLast - x.val.mref.lpmref->reftbl[0].rwFirst + 1
			: x.xltype == xltypeMissing || x.xltype == xltypeNil ? 0
			: 1;
	}
	template<class X>
	inline typename traits<X>::xword columns(const X& x)
	{
		// xltypeRef, xltypeBigData?
		return (x.xltype == xltypeMulti) ? x.val.array.columns 
			: x.xltype == xltypeSRef ? x.val.sref.ref.colLast - x.val.sref.ref.colFirst + 1
			: x.xltype == xltypeRef ? x.val.mref.lpmref->reftbl[0].colLast - x.val.mref.lpmref->reftbl[0].colFirst + 1
			: x.xltype == xltypeMissing || x.xltype == xltypeNil ? 0
			: 1;
	}
	template<class X>
	inline typename traits<X>::xword size(const X& x)
	{
		return rows(x) * columns(x);		
	}

	template<class X>
	inline typename X& index(X& x, typename traits<X>::xword i)
	{
		ensure (i < size(x));

		if (x.xltype == xltypeMulti)
			return x.val.array.lparray[i];
		else
			return x;
	}
	template<class X>
	inline typename const X& index(const X& x, typename traits<X>::xword i)
	{
		ensure (i < size(x));

		if (x.xltype == xltypeMulti)
			return x.val.array.lparray[i];
		else
			return x;
	}

	template<class X>
	inline typename X& index(X& x, typename traits<X>::xword i, typename traits<X>::xword j)
	{
		ensure (i < rows(x));
		ensure (j < columns(x));

		if (x.xltype == xltypeMulti)
			return x.val.array.lparray[i*x.val.array.columns + j];
		else
			return x;
	}
	template<class X>
	inline typename const X& index(const X& x, typename traits<X>::xword i, typename traits<X>::xword j)
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
		return (x.xltype == xltypeMulti) ? x.val.array.lparray : &x;
	}
	template<class X>
	inline const typename X* begin(const X& x)
	{
		return (x.xltype == xltypeMulti) ? x.val.array.lparray : &x;
	}
	template<class X>
	inline typename X* end(X& x)
	{
		return (x.xltype == xltypeMulti) ? x.val.array.lparray + size(x) : &x + 1;
	}
	template<class X>
	inline const typename X* end(const X& x)
	{
		return (x.xltype == xltypeMulti) ? x.val.array.lparray + size(x) : &x + 1;
	}

	template<class X>
	inline double to_number(const X& x)
	{
		double num;

		switch (x.xltype) {
		case xltypeNum:
			num = x.val.num;

			break;
		case xltypeStr:
			num = x.val.str[0]; // so if(x) ... works right

			break;
		case xltypeBool:
			num = x.val.xbool;

			break;
		case xltypeErr:
			num = 0; // so ensure(Err()) fails

			break;
		case xltypeMulti: {
			if (size(x) == 1)
				return to_number(x.val.array.lparray[0]);

			num = 1; // 'if (x) {...}' fails if any 0 or FALSE entries
			for (traits<X>::xword i = 0; i < size(x); ++i) {
				num = (0 != to_number(x.val.array.lparray[i]));
				if (!num)
					break;
			}


			break;
		}
		case xltypeMissing: case xltypeNil:
			num = 0;

			break;
		case xltypeInt:
			num = x.val.w;

			break;
		default:
			num = std::numeric_limits<double>::quiet_NaN();
		}

		return num;
	}
	template<class X>
	inline typename traits<X>::xstring 
	to_string(const X& x, 
		typename traits<X>::xcstr fs = 0, 
		typename traits<X>::xcstr rs = 0)
	{
		traits<X>::xstring s;

		if (x.xltype == xltypeStr) {
			s = traits<X>::xstring(x.val.str + 1, x.val.str[0]);
		}
		else if (x.xltype == xltypeMulti) {
			s = to_string(index(x, 0));
			for (traits<X>::xword i = 1; i < size(x); ++i) {
				if (i % columns(x) == 0) {
					if (rs && *rs) 
						s += rs;
				}
				else if (fs && *fs) {
					s += fs;
				}
				s += to_string(index(x, i));
			}
		}
		else if (x.xltype == xltypeErr) {
			s = traits<X>::Err(x.val.err);
		}
		else {
			LXOPER<X> y = Excel<X>(xlfText, x, traits<X>::format());
			if (y.xltype != xltypeStr)
				s = traits<X>::Err(xlerrValue);
			else
				s = traits<X>::xstring(y.val.str + 1, y.val.str[0]); 
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
				return x.val.sref.ref == y.val.sref.ref;
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
				return x.val.sref.ref == y.val.sref.ref;
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
		if (x.xltype > y.xltype)
			return false;

		switch (x.xltype) {
			case xltypeNum:
				return x.val.num < y.val.num;
			case xltypeStr:
				return ((x.val.str[0] < y.val.str[0]) && 0 == xll::traits<X>::strnicmp(x.val.str + 1, y.val.str + 1, x.val.str[0]))
					|| 0 > xll::traits<X>::strnicmp(x.val.str + 1, y.val.str + 1, __min(x.val.str[0], y.val.str[0]));
			case xltypeBool:
				return x.val.xbool < y.val.xbool;
			case xltypeRef:
				return x.val.sref.ref < y.val.sref.ref;
			case xltypeErr:
				return x.val.err < y.val.err;
			case xltypeMulti:
				if (rows(x) > rows(y))
					return false;
				if (columns(x) > columns(y))
					return false;
				for (traits<X>::xword i = 0; i < rows(x); ++i)
					for (traits<X>::xword j = 0; j < columns(x); ++j) {
						if (operator_less(index(x, i, j), index(y, i, j)))
							return true;
						if (!operator_equals(index(x, i, j), index(y, i, j)))
							return false;
					}
				return false;
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
	}
} // namespace xll

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
