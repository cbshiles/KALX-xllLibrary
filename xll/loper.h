// loper.h - XLOPER with memory handled by Excel
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "xll/fp.h"

// LOPER o = Excel(xlfSomething, NumX(x), StrX(s), ...);
template <class X>
struct LXOPER : public X {
	typedef typename xll::traits<X>::xword xword;

	LXOPER()
	{
		xltype = xltypeNil;
		own(true); // so LXOPER x; Excel4(fun, &x, ...); works right
	}
	LXOPER(const LXOPER& x_)
	{
		own(x_.owner);
		assign(rfree(x_));
	}
	LXOPER(const X& x)
	{
		own(false);
		assign(x);
	}
	LXOPER& operator=(const LXOPER& x_)
	{
		if (this != &x_) {
			lfree();
			own(x_.owner);
			assign(rfree(x_));
		}

		return *this;
	}
	LXOPER& operator=(const X& x)
	{
		lfree();
		own(false);
		assign(x);

		return *this;
	}
	~LXOPER()
	{
		lfree();
	}
	operator double() const
	{
		return to_number(*this);
	}
	const X& operator[](xword i) const
	{
		return xll::Index<X>(*this, i);
	}
	const X& operator()(xword i, xword j) const
	{
		return xll::Index<X>(*this, i, j);
	}
	X* operator&()
	{
		return (X*)this;
	}
	const X* operator&() const
	{
		return (const X*)this;
	}
	X* XLFree(void)
	{
		if (owner)
			xltype |= xlbitXLFree;

		return this;
	}
	// total number of LOPERs that need to be freed
	static int count(void)
	{
		return owned();
	}
private:
	static int owned(int incr = 0)
	{
		static int owned_(0);

		owned_ += incr;

		return owned_;
	}
	void own(bool b) const
	{
		owner = b;
		owned(b);
	}
	void assign(const X& x)
	{
		xltype = x.xltype;
		val = x.val;
	}
	// l-value really gets freed
	void lfree(void)
	{
		if (owner) {
			owned(-1);
			xll::traits<XLOPERX>::Excel(xlFree, 0, 1, this);
		}
	}
	// r-value gets count decremented
	const LXOPER& rfree(const LXOPER& x_) const
	{
		if (x_.owner) {
			owned(-1);
			x_.owner = false;
		}

		return x_;
	}
	mutable bool owner;
};

typedef LXOPER<XLOPER>   LOPER, *LPLOPER;
typedef LXOPER<XLOPER12> LOPER12, *LPLOPER12;
typedef X_(LOPER) LOPERX, *LPLOPERX;
