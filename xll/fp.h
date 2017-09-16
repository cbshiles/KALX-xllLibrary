// fp.h - Two dimensional array of doubles.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include <algorithm>
#include <cstdlib>
#include <new>
#include "xll/error.h"

namespace xll {

	inline bool 
	is_empty(const _FP& fp)
	{
		return fp.rows == 1
			&& fp.columns == 1
			&& _isnan(fp.array[0]);
	}
	inline bool 
	is_empty(const _FP12& fp)
	{
		return fp.rows == 1
			&& fp.columns == 1
			&& _isnan(fp.array[0]);
	}

	inline traits<XLOPER>::xword
	size(const _FP& fp)
	{
		return is_empty(fp) ? 0 : fp.rows * fp.columns;
	}
	inline traits<XLOPER12>::xword
	size(const _FP12& fp)
	{
		return is_empty(fp) ? 0 : fp.rows * fp.columns;
	}

	inline double
	index(const _FP& fp, traits<XLOPER>::xword i, traits<XLOPER>::xword j)
	{
		return fp.array[i*fp.columns + j];
	}
	inline double
	index(const _FP12& fp, traits<XLOPER12>::xword i, traits<XLOPER12>::xword j)
	{
		return fp.array[i*fp.columns + j];
	}

	inline double&
	index(_FP& fp, traits<XLOPER>::xword i, traits<XLOPER>::xword j)
	{
		return fp.array[i*fp.columns + j];
	}
	inline double&
	index(_FP12& fp, traits<XLOPER12>::xword i, traits<XLOPER12>::xword j)
	{
		return fp.array[i*fp.columns + j];
	}

	inline double*
	begin(_FP& fp)
	{
		return &fp.array[0];
	}
	inline double*
	begin(_FP12& fp)
	{
		return &fp.array[0];
	}
	inline double*
	end(_FP& fp)
	{
		return &fp.array[0] + size(fp);
	}
	inline double*
	end(_FP12& fp)
	{
		return &fp.array[0] + size(fp);
	}

	inline const double*
	begin(const _FP& fp)
	{
		return &fp.array[0];
	}
	inline const double*
	begin(const _FP12& fp)
	{
		return &fp.array[0];
	}
	inline const double*
	end(const _FP& fp)
	{
		return &fp.array[0] + size(fp);
	}
	inline const double*
	end(const _FP12& fp)
	{
		return &fp.array[0] + size(fp);
	}

	template<class X>
	class XFP {
		typedef typename traits<X>::xword xword;
		typedef typename traits<X>::xfp xfp;
	public:
		XFP()
			: buf(0)
		{
			realloc(0,0);
		}
		XFP(xword r, xword c) 
			: buf(0) 
		{
			realloc(r, c);
		}
		XFP(const XFP& x) 
			: buf(0)
		{
			realloc(x.rows(), x.columns());
			copy(x.pf);
		}
		XFP(const xfp& x)
			: buf(0)
		{
			realloc(x.rows, x.columns);
			copy(&x);
		}
		XFP& operator=(const XFP& x)
		{
			if (x.this_() != this) {
				realloc(x.rows(), x.columns());
				copy(x.pf);
			}

			return *this;
		}
		~XFP()
		{
			free(buf);
		}

		XFP& operator=(double x)
		{
			std::fill(pf->array, pf->array + size(), x);

			return *this;
		}
		// so FPX x; ... return &x returns the right thing
		xfp* operator&()
		{
			return pf;
		}
		const xfp* operator&() const
		{
			return pf;
		}
		const XFP* this_() const
		{
			return this;
		}
		double* array()
		{
			return pf->array;
		}
		xword rows() const
		{
			return pf->rows;
		}
		xword columns() const
		{
			return pf->columns;
		}
		xword size() const
		{
			return buf ? (is_empty() ? 0 : pf->rows * pf->columns) : 0;
		}

		void reshape(int r, int c)
		{
			realloc(r, c);
		}
		double operator[](xword i) const
		{
			return pf->array[i];
		}
		double& operator[](xword i)
		{
			return pf->array[i];
		}
		double operator()(xword i, xword j) const
		{
			return index(*this, i, j);
//			return operator[](i*columns() + j);
		}
		double& operator()(xword i, xword j)
		{
			return operator[](i*columns() + j);
		}

		double* begin()
		{
			return pf->array;
		}
		const double* begin() const
		{
			return pf->array;
		}
		double* end()
		{
			return pf->array + size();
		}
		const double* end() const
		{
			return pf->array + size();
		}

		bool is_empty(void) const
		{
			return xll::is_empty(*pf);
		}
	private:
		void copy(const xfp* p)
		{
			ensure (pf->rows == p->rows);
			ensure (pf->columns == p->columns);

			pf->array[0] = p->array[0]; // could be empty array
			for (xword i = 1; i < xll::size(*p); ++i)
				pf->array[i] = p->array[i];
		}
		void realloc(int r, int c)
		{
			if (!buf || size() != static_cast<xword>(r)*static_cast<xword>(c)) {
				buf = (char*)::realloc(buf, sizeof(xfp) + r*c*sizeof(double));
				pf = reinterpret_cast<xfp*>(new (buf) xfp);
			}
			//!!! check size
			pf->rows = static_cast<xword>(r);
			pf->columns = static_cast<xword>(c);

			// empty array
			if (r*c == 0) {
				pf->rows = 1;
				pf->columns = 1;
				pf->array[0] = std::numeric_limits<double>::quiet_NaN();
			}
			else if (is_empty()) {
				pf->array[0] = 0; // so resize(1,1) on empty array is not empty
			}
		}
		char* buf;
		xfp* pf;
	};

	typedef XFP<XLOPER>   FP;
	typedef XFP<XLOPER12> FP12;
	typedef XFP<XLOPERX>  FPX;

} // namespace xll

