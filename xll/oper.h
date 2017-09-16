// oper.h - Well behaved data type with the same footprint as the Microsoft SDK XLOPER.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include <cstdlib>
#include <limits>
#include "xloper.h"

// In the global namespace, just like XLOPER.
template<class X>
class XOPER : public X {
public:

	typedef XOPER value_type;
	typedef value_type* pointer;		
	typedef const value_type* const_pointer;		
	typedef value_type& reference;		
	typedef const value_type& const_reference;		
	typedef typename xll::traits<X>::xword size_type;
	typedef std::ptrdiff_t difference_type;

	typedef typename xll::traits<X>::xchar xchar;
	typedef typename xll::traits<X>::xcstr xcstr;
	typedef typename xll::traits<X>::xword xword;
	typedef typename xll::traits<X>::xrw xrw;
	typedef typename xll::traits<X>::xcol xcol;

	XOPER()
	{ 
		xltype = xltypeNil;
	}
	XOPER(const XOPER& x)
	{
		construct(x);
	}
	XOPER(const X& x)
	{
		construct(x);
	}
	XOPER& operator=(const X& x)
	{
		return assign(x);
	}
	XOPER& operator=(const XOPER& x)
	{
		return assign(x);
	}
	~XOPER()
	{
		free();
	}

	bool operator==(const XOPER& x)
	{
		return xll::operator_equals<X>(*this, x);
	}
	bool operator< (const XOPER& x)
	{
		return xll::operator_less<X>(*this, x);
	}
	bool operator==(const X& x)
	{
		return xll::operator_equals<X>(*this, x);
	}
	bool operator<(const X& x)
	{
		return xll::operator_less<X>(*this, x);
	}

	// xltypeNum
	explicit XOPER(double num)
	{
		xltype = xltypeNum;
		val.num = num;
	}
	// Excel promotes all of these types to double.
#define XOPER_NUM(T) explicit XOPER(T num) { xltype = xltypeNum; val.num = num; }
	XOPER_NUM(short)
	XOPER_NUM(unsigned short)
	XOPER_NUM(int)
	XOPER_NUM(unsigned int)
	XOPER_NUM(long)
	XOPER_NUM(unsigned long)
//	XOPER_NUM(size_t)
#undef  XOPER_NUM

	// This may or may not be the right thing to do.
	operator double() const
	{
		return xll::to_number<X>(*this);
	}
	XOPER& operator=(double num)
	{
		free();
		xltype = xltypeNum;
		val.num = num;

		return *this;
	}
	// Excel promotes all of these types to double.
#define XOPER_EQ(T) XOPER& operator=(T num) { free(); xltype = xltypeNum; val.num = num; return *this; }
	XOPER_EQ(short)
	XOPER_EQ(unsigned short)
	XOPER_EQ(int)
	XOPER_EQ(size_t) // unsigned int causes warnings
	XOPER_EQ(long)
	XOPER_EQ(unsigned long)
#undef  XOPER_EQ

#define XOPER_EQU(T) bool operator==(T num) {        \
	return ((xltype == xltypeNum) && val.num == num) \
	|| ((xltype == xltypeInt) && val.w == num)       \
	|| ((xltype == xltypeMulti) && size() == 1 && val.array.lparray[0].xltype == xltypeNum && val.array.lparray[0].val.num == num); }
	//XOPER_EQU(short)
	XOPER_EQU(unsigned short)
	XOPER_EQU(double)
	XOPER_EQU(int)
	XOPER_EQU(size_t)
/*	bool operator==(xword w)
	{
		return operator==(static_cast<double>(w));
	}
*/	//XOPER_EQU(xword)
	//XOPER_EQU(unsigned long)
#undef  XOPER_EQU

	// xltypeStr
	XOPER(xcstr str)
	{
		xltype = xltypeStr;
		val.str = 0;

		realloc(str, xll::traits<X>::strlen(str));
	}
	// Use as XOPER(s + 1, s[0]) for counted strings.
	XOPER(xcstr str, xchar len)
	{
		xltype = xltypeStr;
		val.str = 0;

		realloc(str, len);
	}
	XOPER& operator=(xcstr str)
	{
		free();
		xltype = xltypeStr;
		val.str = 0;

		realloc(str, xll::traits<X>::strlen(str));

		return *this;
	}
	bool operator==(xcstr str)
	{
		return xltype == xltypeStr
			&& val.str[0] == xll::traits<X>::strlen(str)
			&& 0 == xll::traits<X>::strnicmp(str, val.str + 1, val.str[0]);
	}
	XOPER& append(xcstr str, xchar len = 0)
	{
		if (len == 0)
			len = xll::traits<X>::strlen(str);

		if (len > 0) {
			ensure (xltype == xltypeStr);
			val.str = static_cast<xll::traits<X>::xchar*>(::realloc(val.str, (val.str[0] + len + 1)*sizeof(xchar)));
			xll::traits<X>::strncpy(val.str + val.str[0] + 1 , str, len);
			val.str[0] = val.str[0] + len;
		}

		return *this;
	}

	// xltypeBool
	explicit XOPER(bool xbool)
	{
		xltype = xltypeBool;
		val.xbool = xbool;
	}
	XOPER& operator=(bool xbool)
	{
		free();
		xltype = xltypeBool;
		val.xbool = xbool;

		return *this;
	}
	bool operator==(bool xbool)
	{
		return xltype == xltypeBool && val.xbool == static_cast<xll::traits<X>::xbool>(xbool);
	}

	// xltypeRef - omitted

	// xltypeErr - use XErr<X>

	// xltypeMulti
	XOPER(xword r, xword c)
	{
		xltype = xltypeMulti;

		val.array.rows = 0;
		val.array.columns = 0;
		val.array.lparray = 0;

		realloc(r, c);
	}
	xword rows(void) const
	{
		return xll::rows<X>(*this);
	}
	xword columns(void) const
	{
		return xll::columns<X>(*this);
	}
	xword size(void) const
	{
		return xll::size<X>(*this);
	}

	XOPER& resize(xword r, xword c)
	{
		if (xltype != xltypeMulti) {
			XOPER tmp = *this;
			free();

			xltype = xltypeMulti;

			val.array.rows = 0;
			val.array.columns = 0;
			val.array.lparray = 0;

			realloc(r, c, &tmp, 1);
		}
		else {
			realloc(r, c);
		}

		return *this;
	}
	XOPER& operator[](xword i)
	{
		return static_cast<XOPER&>(xll::index<X>(*this, i));
	}
	const XOPER& operator[](xword i) const
	{
		return static_cast<const XOPER&>(xll::index<X>(*this, i));
	}
	XOPER& operator()(xword i, xword j)
	{
		return static_cast<XOPER&>(xll::index<X>(*this, i, j));
	}
	const XOPER& operator()(xword i, xword j) const
	{
		return static_cast<const XOPER&>(xll::index<X>(*this, i, j));
	}

	// STL friendly
	XOPER* begin(void)
	{
		return static_cast<XOPER*>(xll::begin<X>(*this));
	}
	const XOPER* begin(void) const
	{
		return static_cast<const XOPER*>(xll::begin<X>(*this));
	}
	XOPER* end(void)
	{
		return static_cast<XOPER*>(xll::end<X>(*this));
	}
	const XOPER* end(void) const
	{
		return static_cast<const XOPER*>(xll::end<X>(*this));
	}
	void push_back(const X& x)
	{
		if (!size()) {
			operator=(x);
		}
		else {
			xword n = size();

			if (columns() == xll::columns(x)) {
				resize(rows() + xll::rows(x), columns());
			}
			else {
				ensure (rows() == 1);
				resize(1, n + xll::size(x));
			}
			for (xword i = 0; i < xll::size(x); ++i)
				operator[](n + i) = xll::index(x, i);
		}
	}
	void pop_back(xword n = 1)
	{
		if (rows() == 1) {
			ensure (n <= columns());
			resize(1, columns() - n);
		}
		else {
			ensure (n <= rows());
			resize(rows() - n, columns());
		}
	}

	// xltypeMissing - use XMissing<X>
	// xltypeNil - use XNil<X>

	// xltypeSRef
	XOPER(xrw row, xcol column, xrw height, xcol width)
	{
		xltype = xltypeSRef;
		val.sref.count = 1;
		val.sref.ref.rwFirst = row;
		val.sref.ref.rwLast = row + height - 1;
		val.sref.ref.colFirst = column;
		val.sref.ref.colLast = column + width - 1;
	}
	
	// xltypeInt - use XInt<X> if needed

	// xltypeBigData
	XOPER(long cbData, void* lpbData)
	{
		xltype = xltypeBigData;
		val.bigdata.h.lpbData = 0;

		realloc(cbData, lpbData);
	}

private:
	void construct(const X& x)
	{
		// ensure (x.xltype & (xltypeNum|...));
		if (x.xltype == xltypeStr) {
			xltype = xltypeStr;
			val.str = 0;

			realloc(x.val.str + 1, x.val.str[0]);
		}
		else if (x.xltype == xltypeMulti) {
			xltype = xltypeMulti;
			val.array.rows = 0;
			val.array.columns = 0;
			val.array.lparray = 0;

			realloc(x.val.array.rows, x.val.array.columns, x.val.array.lparray);
		}
		else if (x.xltype == xltypeBigData) {
			xltype = xltypeBigData;
			val.bigdata.h.lpbData = 0;

			realloc(x.val.bigdata.cbData, x.val.bigdata.h.lpbData);
		}
		else {
			xltype = x.xltype;
			val = x.val;
		}
	}

	XOPER& assign(const X& x)
	{
		if (this != &x) {
			free();
			construct(x);
		}

		return *this;
	}

	void realloc(xcstr str, xchar count)
	{
		ensure (xltype == xltypeStr);

		ensure (0 != (val.str = static_cast<xll::traits<X>::xchar*>(::realloc(val.str, (count + 1)*sizeof(xchar)))));
		val.str[0] = count;
		if (count)
			xll::traits<X>::strncpy(val.str + 1, str, count);
	}

	// reallocate and initialize if pa is not null
	void realloc(xword r, xword c, const X* pa = 0, xword na = 0)
	{
		if (r*c == 0) {
			free();

			return;
		}
		ensure (r <= xll::limits<X>::maxrows);
		ensure (c <= xll::limits<X>::maxcols);
		ensure (xltype == xltypeMulti);

		xword n = size();
		// free existing OPERs
		for (xword i = r*c; i < n; ++i)
			operator[](i).free();

		val.array.rows = r;
		val.array.columns = c;

		if (n != size())
			ensure (0 != (val.array.lparray = static_cast<X*>(::realloc(val.array.lparray, r*c*sizeof(X)))));

		// intitialize new OPERs
		for (xword i = n; i < size(); ++i)
			val.array.lparray[i].xltype = xltypeNil;

		if (pa) {
			ensure (na <= size());

			if (na == 0)
				na = size();

			for (xword i = 0; i < na; ++i)
				operator[](i) = pa[i];
		}
	}
	void realloc(long cbData, void* lpbData)
	{
		ensure (xltype == xltypeBigData);

		val.bigdata.cbData = cbData;
		if (cbData) {
			ensure (0 != (val.bigdata.h.lpbData = static_cast<BYTE*>(::realloc(val.bigdata.h.lpbData, cbData))));
			memcpy(val.bigdata.h.lpbData, lpbData, cbData);
		}
		else { // Excel handle
			val.bigdata.h.hdata = lpbData;
		}
	}

	void free()
	{
		switch (xltype) {
		case xltypeStr:
			::free(val.str);
			val.str = 0;

			break;
		case xltypeMulti:
			for (xword i = 0; i < size(); ++i)
				operator[](i).free(); // need to call explicitly
			::free(val.array.lparray);
			val.array.rows = 0;
			val.array.columns = 0;
			val.array.lparray = 0;

			break;
		case xltypeBigData:
			if (val.bigdata.cbData)
				::free(val.bigdata.h.lpbData);
			val.bigdata.cbData = 0;
			val.bigdata.h.lpbData = 0;

			break;
		}

		xltype = xltypeMissing; // just like xlFree
	}
};

typedef XOPER<XLOPER>   OPER,   *LPOPER;
typedef XOPER<XLOPER12> OPER12, *LPOPER12;
typedef X_(OPER)        OPERX,  *LPOPERX;

// Disambiguation classes.
template<class X>
struct XNum : public OPERX {
	explicit XNum(double num = 0)
		: OPERX(num)
	{ }
	XNum& operator=(double num)
	{
		OPERX::operator=(num);

		return *this;
	}
};
typedef XNum<XLOPER>   Num;
typedef XNum<XLOPER12> Num12;
typedef X_(Num)        NumX;

template<class X>
struct XStr : public OPERX {
	explicit XStr(xcstr str = xll::traits<X>::null())
		: OPERX(str)
	{ }
	XStr(xcstr str, xchar len)
		: OPERX(str, len)
	{ }
	XStr& operator=(xcstr str)
	{
		XOPER::operator=(str);

		return *this;
	}
};
typedef XStr<XLOPER12> Str12;
typedef XStr<XLOPER>   Str;
typedef X_(Str)        StrX;

template<class X>
struct XBool : public OPERX {
	XBool(bool xbool = false)
		: OPERX(xbool)
	{ }
	XBool& operator=(bool xbool)
	{
		OPERX::operator=(xbool);

		return *this;
	}
};
typedef XBool<XLOPER12> Bool12;
typedef XBool<XLOPER>   Bool;
typedef X_(Bool)        BoolX;

template<class X>
struct XInt : public OPERX {
	typedef typename xll::traits<X>::xint xint;
	explicit XInt(xint w = 0)
	{
		xltype = xltypeInt;
		val.w = w;
	}
	XInt& operator=(xint w)
	{
		OPERX::operator=(w); // xltypeNum, just like Excel

		return *this;
	}
};
typedef XInt<XLOPER>   Int;
typedef XInt<XLOPER12> Int12;
typedef X_(Int)        IntX;

// The following types inherit from XLOPERX, not OPERX

template<class X>
struct XErr : public X {
	XErr(WORD err)
	{
		xltype = xltypeErr;
		val.err = err;
	}
};
typedef XErr<XLOPER12> Err12;
typedef XErr<XLOPER>   Err;
typedef X_(Err)        ErrX;

template<class X>
struct XMissing : public X {
	XMissing()
	{
		xltype = xltypeMissing;
	}
	operator X&()
	{
		return *this;
	}
	operator const X&() const
	{
		return *this;
	}
};
typedef XMissing<XLOPER12> Missing12;
typedef XMissing<XLOPER>   Missing;
typedef X_(Missing)        MissingX;

template<class X>
struct XNil : public X {
	XNil()
	{
		xltype = xltypeNil;
	}
};
typedef XNil<XLOPER12> Nil12;
typedef XNil<XLOPER>   Nil;
typedef X_(Nil)        NilX;

// no XMulti class - use XOPER

inline bool operator==(const OPER12& x, const OPER12& y)
{
	return xll::operator_equals<XLOPER12>(x, y);
}
inline bool operator<(const OPER12& x, const OPER12& y)
{
	return xll::operator_less<XLOPER12>(x, y);
}
inline bool operator==(const OPER& x, const OPER& y)
{
	return xll::operator_equals<XLOPER>(x, y);
}
inline bool operator<(const OPER& x, const OPER& y)
{
	return xll::operator_less<XLOPER>(x, y);
}

inline bool operator==(const XLOPER12& x, const OPER12& y)
{
	return xll::operator_equals<XLOPER12>(x, y);
}
inline bool operator<(const XLOPER12& x, const OPER12& y)
{
	return xll::operator_less<XLOPER12>(x, y);
}
inline bool operator==(const XLOPER& x, const OPER& y)
{
	return xll::operator_equals<XLOPER>(x, y);
}
inline bool operator<(const XLOPER& x, const OPER& y)
{
	return xll::operator_less<XLOPER>(x, y);
}

inline bool operator==(const OPER12& x, const XLOPER12& y)
{
	return xll::operator_equals<XLOPER12>(x, y);
}
inline bool operator<(const OPER12& x, const XLOPER12& y)
{
	return xll::operator_less<XLOPER12>(x, y);
}
inline bool operator==(const OPER& x, const XLOPER& y)
{
	return xll::operator_equals<XLOPER>(x, y);
}
inline bool operator<(const OPER& x, const XLOPER& y)
{
	return xll::operator_less<XLOPER>(x, y);
}

