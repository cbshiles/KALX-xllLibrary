// hash.h - simple hash function
#pragma once

namespace utility {

	template<class T>
	inline unsigned int hash(const T* str, T count = 0, bool positive = false)
	{			
		unsigned int ui = 5381;

		T n = count;
		while (count ? n-- : *str)
			ui = ui * 33 + *str++;

		return positive ? ui >> 1 : ui;
	}	

} // namespace xll