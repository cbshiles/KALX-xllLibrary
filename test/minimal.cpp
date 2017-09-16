// minimal.cpp
//#define EXCEL12
#include "xll/xll.h"

using namespace xll;
/*
static AddIn xai_minimal(
	"?xll_minimal", XLL_DOUBLE XLL_DOUBLE, 
	"XLL.MINIMAL", "Category"
);
double WINAPI 
xll_minimal(double x) 
{
#pragma XLLEXPORT
	return x; 
}
*/
static AddIn xai_minimal_macro("?xll_minimal_macro", "XLL.MACRO");
int WINAPI
xll_minimal_macro(void)
{
#pragma XLLEXPORT

	return 1;
}
