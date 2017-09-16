// handles.cpp - test handle class
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "test.h"
//#include "../xll/nhandle.h"

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
	_T("BASE_"), _T("Int")
);
HANDLEX WINAPI
xll_base(short b)
{
#pragma XLLEXPORT
	handlex h; // lower case defaults to #NUM!
	
	try {
		handle<base> h_(new base(b));
		ensure (h_);

		h = h_.get();
	}
	catch (const std::exception& ex) {
		ex.what();

		return 0;
	}

	return h;
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
	int value;

	try {
		handle<base> b_(h);
		ensure (b_);

		value = b_->value();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return value;
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
	handlex h;

	try {
		// put pointer in the base bucket
		handle<base> h_(new derived(b, d));
		ensure (h_);

		h = h_.get();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return h;
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
	int value2;
	
	try {
		// use RTTI
		derived *pd = dynamic_cast<derived*>(handle<base>(h).ptr());
		ensure (pd);
		value2 = pd->value2();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return value2;
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
	ensure (bb);
	ensure (1 == bb->value());

	handle<derived> hd(new derived(2, 3));
	HANDLEX xhd = hd.get();
	handle<derived> dd(xhd);
	ensure (2 == dd->value());
	ensure (3 == dd->value2());
/*
	handle<base> hb2;
	hb2 = &b;
	HANDLEX xhb2 = hb2.get();
	handle<base> hbb2(xhb2);
	ensure (1 == hbb2->value());
*/
}

struct foo {
	char* s_;
};
foo* foo_new(const char* s)
{
	foo* pf = new foo;
	pf->s_ = new char[strlen(s) + 1];
	strcpy(pf->s_, s);

	return pf;
}
void foo_delete(foo* pf)
{
	delete[] pf->s_;
	delete pf;
}
void
test_delete(void)
{
	handle<foo> h1(new foo());
//	handle<foo> h2(foo_new("abc")); // leaks
	handle<foo,void (*)(foo*)> h3(foo_new("def"), foo_delete);
}
static AddIn xai_test_foo(
	"?xll_test_foo", XLL_HANDLE XLL_CSTRING XLL_UNCALCED,
	"TEST.FOO", "Str"
);
HANDLEX WINAPI
xll_test_foo(const char* s)
{
#pragma XLLEXPORT
	handle<foo, void(*)(foo*)> h(foo_new(s), foo_delete);

	return h.get();
}
static AddIn xai_test_foo_get(
	"?xll_test_foo_get", XLL_CSTRING XLL_HANDLE,
	"TEST.FOO.GET", "Str"
);
const char* WINAPI
xll_test_foo_get(HANDLEX h)
{
#pragma XLLEXPORT
	// handle<foo> h_(h); is different bucket
	handle<foo,void(*)(foo*)> h_(h);

	return h_ ? h_->s_ : 0;
}

int
test_handles(void)
{
	try {
//		_crtBreakAlloc = 1289;
		test_base_derived();
		test_delete();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<OpenX> xao_handles(test_handles);

#endif // _DEBUG
/*
// construct C++ object in Excel
static AddInX xai_nbase(
	_T("?xll_nbase"), XLL_HANDLEX XLL_SHORTX XLL_UNCALCEDX,
	_T("NBASE"), _T("Int")
);
HANDLEX WINAPI
xll_nbase(short b)
{
#pragma XLLEXPORT
	HANDLEX h;
	
	try {
		nxll::nhandle<base> h_(new base(b));
		ensure (h_);

		h = h_.get();

//		nxll::nhandle<base>::gc();
	}
	catch (const std::exception& ex) {
		ex.what();

		return 0;
	}

	return h;
}

// member function
static AddInX xai_nbase_value(
	_T("?xll_nbase_value"), XLL_LONGX XLL_HANDLEX XLL_UNCALCEDX,
	_T("NBASE.VALUE"), _T("Handle")
);
LONG WINAPI
xll_nbase_value(HANDLEX h)
{
#pragma XLLEXPORT
	int value;

	try {
		nxll::nhandle<base> b_(h);
		ensure (b_);

		value = b_->value();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return value;
}
*/

// return #NUM!
static AddInX xai_double_error(
	_T("?xll_double_error"), XLL_DOUBLEX XLL_DOUBLEX,
	_T("DOUBLE.ERROR"), _T("")
);
double WINAPI
xll_double_error(double x)
{
#pragma XLLEXPORT
	x = x;

	return std::numeric_limits<double>::quiet_NaN();
}