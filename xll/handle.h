// handle.h - simple handle class
// Copyright (c) 2006 KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include <set>
#include <memory>
#include "oper.h"

typedef double HANDLEX;

#define XLL_HANDLE XLL_DOUBLE
#define XLL_HANDLE12 XLL_DOUBLE12
#ifdef EXCEL12
#define XLL_HANDLEX XLL_HANDLE12
#else
#define XLL_HANDLEX XLL_HANDLE
#endif

//static_assert (sizeof(HANDLEX) == sizeof(long long), "handles and long longs must be the same size");

namespace xll {

	// handle to pointer
	template<class T>
	inline T* h2p(HANDLEX x)
	{
		union { T* p_; HANDLEX x_; long l_; } px;
		
		px.l_ = static_cast<long>(x);
//		px.x_ = x;

		return px.p_;
	}
	// pointer to handle
	template<class T>
	inline HANDLEX p2h(T* p)
	{
		union { T* p_; HANDLEX x_; long l_; } px;
		
		px.p_ = p;

		return static_cast<HANDLEX>(px.l_);
//		return static_cast<HANDLEX>(px.x_);
	}

	template<class T>
	class handle {
		typedef typename std::set<std::shared_ptr<T>> handle_set;
		static handle_set& handles(void)
		{
			static handle_set h_;

			return h_;
		}
		static void erase(T* p)
		{
			handle_set& h(handles());
			handle_set::iterator i = find(p);

			if (i != h.end()) {
				h.erase(i);
			}
		}
		// return 0 if p not in handle set
		static T* check(T* p)
		{
			handle_set& h(handles());
			handle_set::const_iterator i = find(p);
		
			return i == h.end() ? 0 : p;
		}
		static void insert(T* p)
		{
			std::pair<handle_set::iterator,bool> ib = handles().insert(std::shared_ptr<T>(p));

			// better not already exist
			ensure (ib.second == true);
		}

		//!!!could be more efficient!!!
		static typename handle_set::iterator find(T* p)
		{
			// calls delete on p ???WTF!!!
			//return handles().find(std::shared_ptr<T>(p));

			handle_set::const_iterator i;

			for (i = handles().begin(); i != handles().end(); ++i) {
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

			OPERX x = Excel<XLOPERX>(xlCoerce, Excel<XLOPERX>(xlfCaller));
			if (x.xltype == xltypeNum) {
				// otherwise it is junk
				q = h2p<T>(static_cast<HANDLEX>(x.val.num));
			}

			return q;
		}

		T* p_;

	public:
		// Called as handle<T> h(new T(...));
		// in constructor. Must be uncalced.
		handle(T* p = 0)
			: p_(p)
		{
			if (p) {
				T* q = old_handle();
				if (q) {
					erase(q); // 
				}
				// do after checking old handle
				// might have same handle in old spreadsheet
				insert(p);
				
			}
		}
		// Called as handle<T> h(handle) in "member" function.
		// Need not be uncalced.
		handle(HANDLEX p, bool checked = true)
			: p_(0)
		{
			p_ = h2p<T>(p);

			if (checked) {
				p_ = check(p_);
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

		// ignore p???
		void free()
		{
			delete p_;
		}

		HANDLEX get() const
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