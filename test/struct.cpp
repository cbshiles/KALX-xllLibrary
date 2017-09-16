// struct.cpp - accessors for structs
/*
Constructors return handle to struct.
STRUCT.SET(h, {"k",...}, {v,...}) -> h
STRUCT.GET(h, {"k",...}) -> {v,...}

	OPERX v;
	set(h, "k", v);
	v = get(h, "k"); 

	Struct s = map[typeid(h).hash_code()];
	Member m = s["k"];
	v = m.get();
	m.set(v);

CamelCaseName -> CCN must be unique
allow "ccn" exactly in lower case
xll::struct_ aka_curve(AKACurve, "AKA.CURVE")
	.member(double, "d", "Double", "is a double.", 3.14)
	.member(int, "i", "Int", "is an int.", 123)

	map<hash,map<string,member>>
OPERX struct_get(const struct& s, const string& k)
{
}
);

inline int struct_get(

*/
#include "xll/xll.h"

typedef xll::traits<XLOPERX>::xword xword;

namespace xll {

}

using namespace xll;


template<class T>
struct member {
	typedef T type;
	size_t offset;
};

template<class T>
inline OPERX get(void* p, const member<T>& m)
{
	T t = *reinterpret_cast<const T*>(static_cast<const char*>(p) + m.offset);

	return OPERX(t);
}

template<typename T>
struct type_traits {
	static const int type;
};
template<>
struct type_traits<double> {
	static const int type = xltypeNum;
};
template<>
struct type_traits<char*> {
	static const int type = xltypeStr;
};
template<>
struct type_traits<int> {
	static const int type = xltypeInt;
};

template<class T>
inline T get(const void* p, const member<T>& m)
{
	ensure (type_traits<T>::type == m.type);

	return *reinterpret_cast<const T*>(static_cast<const char*>(p) + m.offset);
}

//std::map<std::basic_string<TCHAR>, Member>;

#ifdef _DEBUG

void test_ccs(void)
{
	ensure (canonical_string(_T("CamelCaseString")) == _T("ccs"));
	ensure (canonical_string(_T("Camel Case String")) == _T("ccs"));
	ensure (canonical_string(_T("ccs")) == _T("ccs"));
	ensure (canonical_string(_T("ccss")) != _T("ccs"));
}
#include <tuple>
int xll_test_struct(void)
{
	try {
		test_ccs();

		auto m = std::make_tuple(1, 3);

		


		struct S {
			int i;
			double d;
			char* s;
		};
//		S s = {1, 2.3, "foo"};
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<OpenAfterX> xao_test_struct(xll_test_struct);

#endif // _DEBUG