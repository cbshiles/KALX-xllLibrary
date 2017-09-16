// nhandle.h - new handle class
// Copyright (c) 2012 KALX, LLC. All rights reserved. No warranty is made.
/*
handle
 set.name _ptr cell address (_ptr follows cell moves)
 get cell formula
 get cell address
*/
#pragma once
#include "oper.h"

//#pragma comment(linker, "/section:.handleseg,RWS")

typedef XLOPERX HANDLEX;

#define XLL_HANDLE XLL_DOUBLE
#define XLL_HANDLE12 XLL_DOUBLE12
#ifdef EXCEL12
#define XLL_HANDLEX XLL_HANDLE12
#else
#define XLL_HANDLEX XLL_HANDLE
#endif

#define XLL_HANDLE_PREFIX OPERX(_T("_."))

namespace nxll {

	// handle to pointer
	template<class T>
	inline T* h2p(HANDLEX x)
	{
		union { T* p_; HANDLEX x_; long l_; } px;
		
		px.l_ = static_cast<long>(x);

		return px.p_;
	}
	// pointer to handle
	template<class T>
	inline HANDLEX p2h(T* p)
	{
		union { T* p_; HANDLEX x_; long l_; } px;
		
		px.p_ = p;

		return static_cast<HANDLEX>(px.l_);
	}

	template<class T>
	class nhandle {
		static OPERX name(T* p)
		{
			OPERX o;
			
			o = ExcelX(xlfConcatenate, XLL_HANDLE_PREFIX, OPERX(nxll::p2h<T>(p)));

			return o;
		}
		void insert(T* p)
		{
			ExcelX(xlfSetName, name(p), OPERX(p2h<T>(p)));
		}

		static void erase(const OPERX& name)
		{
			OPERX h = ExcelX(xlfGetName, name);

			if (h.xltype == xltypeNum) {
				T* p = h2p<T>(h.val.num);
				if (p)
					delete p;
			}
		}
		static T* check(T* p)
		{
			OPERX x;

			x = ExcelX(xlfGetName, name(p));
			if (xll::type(x) == xltypeErr)
				return 0;

			return p;
		}
	public:
		static void gc(void)
		{
			OPERX names = ExcelX(xlfNames, MissingX(), OPERX(3), ExcelX(xlfConcatenate, XLL_HANDLE_PREFIX, OPERX(_T("*"))));

			for (USHORT i = 0; i < names.size(); ++i) {
				erase(names[i]);
			}
		}

		T* p_;
	public:
		// Called as handle<T> h(new T(...));
		// in constructor. Must be uncalced.
		nhandle(T* p = 0)
			: p_(p)
		{
			if (p) {
				insert(p);
			}
		}
		// Called as handle<T> h(handle) in "member" function.
		// Need not be uncalced.
		nhandle(HANDLEX p, bool checked = true)
		{
			if (!p) {
				p_ = 0;
			}
			else {
				p_ = h2p<T>(p);
				if (checked)
					p_ = check(p_);
			}
		}
		~nhandle()
		{ }

		bool operator!() const
		{
			return !p_;
		}
		HANDLEX get() const
		{
			return p2h<T>(p_);
		}
		T* ptr() const
		{
			p_;
		}
		operator T*()
		{
			return p_;
		}
		const T* operator->() const
		{
			return p_;
		}
		T* ptr()
		{
			return p_;
		}
		T* operator->()
		{
			return p_;
		}
		T& operator*()
		{
			return *p_;
		}
		const T& operator*() const
		{
			return *p_;
		}
	};

} // namespace xll
