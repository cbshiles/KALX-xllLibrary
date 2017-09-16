// handles.cpp - test handle class
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "test.h"

using namespace xll;

class base {
	int data_;
public:
	base(int data) 
		: data_(data) 
	{ }
	virtual ~base() // for RTTI
	{ }
	int value(void) const
	{
		return data_;
	}
};

class derived : public base {
	int data_;
public:
	derived(int bdata, int ddata)
		: base(bdata), data_(ddata)
	{ }
	~derived()
	{ }
	int value2(void) const
	{
		return data_;
	}
};

// construct C++ object in Excel
static AddInX xai_base(
	_T("?xll_base"), XLL_HANDLEX XLL_SHORTX XLL_UNCALCEDX,
	_T("BASE"), _T("Int")
);
HANDLEX WINAPI
xll_base(short b)
{
#pragma XLLEXPORT
	handle<base> h;
	
	try {
		h = new base(b);
	}
	catch (const std::exception& ex) {
		ex.what();

		return 0;
	}

	return h.get();
}
// member function
static AddInX xai_base_value(
	_T("?xll_base_value"), XLL_LONGX XLL_HANDLEX,
	_T("BASE.VALUE"), _T("Handle")
);
LONG WINAPI
xll_base_value(HANDLEX h)
{
#pragma XLLEXPORT
	handle<base> b_;
	
	try {
		b_ = h;
		ensure (b_);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return b_->value();
}

// BASE.VALUE(DERIVED(i1,i2)) returns i2.
static AddInX xai_derived(
	_T("?xll_derived"), XLL_HANDLEX XLL_SHORTX XLL_SHORTX XLL_UNCALCEDX,
	_T("DERIVED"), _T("Int, Int2")
);
HANDLEX WINAPI
xll_derived(short b, short d)
{
#pragma XLLEXPORT
	// put pointer in the base bucket
	handle<base> h;
	
	try {
		h = new derived(b, d);
		ensure (h);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return h.get();
}
// member function in derived class
static AddInX xai_derived_value(
	_T("?xll_derived_value"), XLL_LONGX XLL_HANDLEX,
	_T("DERIVED.VALUE2"), _T("Handle")
);
LONG WINAPI
xll_derived_value(HANDLEX h)
{
#pragma XLLEXPORT
	// use RTTI
	derived *pd(0);
	
	try {
		pd = dynamic_cast<derived*>(handle<base>(h).ptr());
		ensure (pd);
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return pd->value2();
}

#ifdef _DEBUG

void
test_base_derived(void)
{
	base b(1);
	ensure (1 == b.value());

	derived d(2, 3);
	ensure (2 == d.value());
	ensure (3 == d.value2());

	handle<base> hb(new base(1));
	HANDLEX xhb = hb.get();
	handle<base> bb(xhb);
	ensure (1 == bb->value());

	handle<derived> hd(new derived(2, 3));
	HANDLEX xhd = hd.get();
	handle<derived> dd(xhd);
	ensure (2 == dd->value());
	ensure (3 == dd->value2());

	handle<base> hb2;
	hb2 = &b;
	HANDLEX xhb2 = hb2.get();
	handle<base> hbb2(xhb2);
	ensure (1 == hbb2->value());

}

int
test_handles(void)
{
	try {
		test_base_derived();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<Open> xao_handles(test_handles);

#endif // _DEBUG