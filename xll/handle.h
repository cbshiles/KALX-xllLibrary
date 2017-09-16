// handle.h - simple handle class
// Copyright (c) 2006 KALX, LLC. All rights reserved. No warranty is made.
// Loading xll will cause handles to recalculate. (???)

#pragma once
#include <list>
#include <memory>
#include "xll/oper.h"

// use char* instead???
typedef double HANDLEX;
//#define HANDLEX double
/*
template<class T>
inline T* handlex(double h)
{
	union { double d; T* h; } dh;

	dh.d = h;

	return dh.h;
}
*/
// handle must be same size as pointer
//#define compile_time_assert(cond) extern char assertion[(cond) ? 1 : 0]
//compile_time_assert(sizeof(long) == sizeof(void*));

#define XLL_HANDLE XLL_DOUBLE
#define XLL_HANDLE12 XLL_DOUBLE12
#ifdef EXCEL12
#define XLL_HANDLEX XLL_HANDLE12
#else
#define XLL_HANDLEX XLL_HANDLE
#endif

namespace xll {


	// complete list of all handles
	extern std::list< std::tr1::shared_ptr<void> > handle_list;

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
	class handle {
		T* p_;

		//!!!could be more efficient!!!
		static std::list< std::tr1::shared_ptr<void> >::iterator find(T* p)
		{
			std::list< std::tr1::shared_ptr<void> >::iterator i;

			for (i = handle_list.begin(); i != handle_list.end(); ++i) {
				if (i->get() == p) {
					break;
				}
			}

			return i;
		}
		// get previous cell contents
		T* old_handle()
		{
			T* q(0);

			LOPERX x = Excel<XLOPERX>(xlCoerce, Excel<XLOPERX>(xlfCaller));
			if (x.xltype == xltypeNum) {
				// otherwise it is junk
				q = h2p<T>(static_cast<HANDLEX>(x.val.num));
			}

			return q;
		}
	public:
		// Called as handle<T> h(new T(...));
		// in constructor. Must be uncalced.
		handle(T* p = 0)
			: p_(p)
		{
			if (p) {
				T* q = old_handle();
				if (q) {
					// if we've seen it before, erase it
					std::list< std::tr1::shared_ptr<void> >::iterator i(find(q));
					if (i != handle_list.end())
						handle_list.erase(i);
				}
				// do after checking old handle
				// might have same handle in old spreadsheet
				handle_list.push_back(std::tr1::shared_ptr<T>(p));
			}
		}
		// Called as handle<T> h(handle) in "member" function.
		// Need not be uncalced.
		handle(HANDLEX p, bool checked = true)
			: p_(0)
		{
			p_ = h2p<T>(p);

			if (checked) {
				std::list< std::tr1::shared_ptr<void> >::iterator i(find(p_));
				if (i == handle_list.end())
					p_ = 0;
			}
		}
		handle(const handle& p)
			: p_(p.p_)
		{ }
		handle& operator=(const handle& p)
		{
			if (this != &p)
				p_ = p.p_;

			return *this;
		}
		~handle()
		{ }

		// ignore p
		void free()
		{
			delete p_;
		}

		HANDLEX handlex() const
		{
			return p2h<T>(p_);
		}
		operator T*()
		{
			return p_;
		}
		const T* operator->() const
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