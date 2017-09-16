#define EXCEL12 
#include "xll/xll.h"
#include "xll/utility/cred.h"

using namespace xll;

int
xll_authenticate(void)
{
#pragma XLLEXPORT
	// Call authentication on stored credentials.
	Cred cred(_T("XLL Application"), [](PCTSTR u, PCTSTR p) { return u && *u == _T('u') && p && *p == _T('p'); });

	if (!cred) {
		// Prompt loop for user name and password.
		cred.Prompt(0, _T("Caption"), _T("User: u Password: p"));
	}

	if (!cred) {
		ExcelX(xlcAlert, OPERX(_T("Login failed.")));
	}

	return cred;
}

//static Auto<OpenX> xao_authenticate(xll_authenticate);