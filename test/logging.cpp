// logging.cpp - test event logging.
#include "xll/utility/log.h"
#include "xll/xll.h"

using namespace xll;

static Log::EventSource e(_T("TestLog"));

static int
xll_test_log(void)
{
	e.ReportError(_T("Error"), 1);
	e.ReportWarning(_T("Warning"), 2);
	e.ReportInformation(_T("Information"), 3);

	return 1;
}
static Auto<Open> xao_test_log(xll_test_log);
