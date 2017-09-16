// strings.h - coversion routines for strings
#include <cstdlib>
#include <string>

namespace strings {

inline std::string
wcstombs(const std::basic_string<wchar_t>& xs)
{
	std::string s;

	s.resize(xs.size() + 1);
	wcstombs_s(0, &s[0], s.size(), &xs[0], xs.size());

	return s;
}
inline std::basic_string<wchar_t>
mbstowcs(const std::string& s)
{
	std::basic_string<wchar_t> xs;

	xs.resize(s.size() + 1);
	mbstowcs_s(0, &xs[0], xs.size(), &s[0], s.size()); 

	return xs;
}

} // strings