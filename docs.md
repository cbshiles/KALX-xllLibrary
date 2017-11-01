Creating Excel add-ins using the xll library

Creating Excel add-ins from scratch is not trivial, as anyone who has attempted to do so using the Microsoft Excel SDK soon realizes. A good exercise would be to download that and attempt to create your own add-in.

Back already? I can't blame you.

An xll add-in is just a dll with well-known entry points. When the add-in is opened, Excel looks for a function called xlAutoOpen and calls it. The job of xlAutoOpen is to register functions that extend the functionality of Excel.

The xll library provides this and other well-known functions so you don't have to write them. All you need to do is create AddIn objects with the information Excel needs to register add-in functions (or macros). To create the add-in function XLL.EXP that calls the C standard library double exp(double) function all you need to write is:

#include <cmath>
#include "xll/xll.h"

using namespace xll;

static AddIn xai_exp(
	"?xll_exp", XLL_DOUBLE XLL_DOUBLE,
	"XLL.EXP", "Number"
);
double WINAPI xll_exp(double number)
{
#pragma XLLEXPORT

	return exp(number);
}
Click on each class to learn more.

class	description
AddIn	The class for creating Excel add-in functions and macros.
OPER	The variant datatype corresponding to a cell or range of cells.
Excel	Calls Excel internal functions and macros.
Register	Call xlfRegister with the appropriate arguments.
Args	Arguments passed to Register.
FP	A two dimensional array of floating point numbers.
Handles	Embed C++ objects in Excel.
Enum	Enumerated values.
Auto	Run macros from xlAutoXXX functions.
Last edited Apr 28, 2015 at 8:18 AM by keithalewis, version 18