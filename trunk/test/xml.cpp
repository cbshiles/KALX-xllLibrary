// xml.cpp - test xml.h classes
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#include "xll/maml/xml.h"
//#define EXCEL12
#include "xll/xll.h"

using namespace xll;
using namespace nxml;

typedef traits<XLOPERX>::xchar xchar;
typedef traits<XLOPERX>::xcstr xcstr;
typedef traits<XLOPERX>::xstring xstring;

inline xstring
q(xcstr t)
{
	xstring s(t);

	s.append(_T("\""));
	s.insert(0, _T("\""));

	return s;
}
void
xll_test_string(void)
{
	xcstr s;
	string<xchar> t1;
	ensure (t1.size() == 0);
	ensure (t1 == _T(""));
	s = t1;
	ensure (*s == 0);

	string<xchar> t2(_T("foo"));
	ensure (t2.size() == 3);
	ensure (t2 == _T("foo"));
	s = t2;
	ensure (*s == _T('f'));

	xstring s1(_T("bar"));
	string<xchar> t3(s1);
	ensure (t3.size() == 3);
	ensure (t3 == _T("bar"));
	s = t3;
	ensure (*s == _T('b'));

	t2.append(_T(", ")).append(t3);
	ensure (t2.size() == 8);
	ensure (t2 == _T("foo, bar"));

	xstring s2 = q((string<xchar>(),_T("text")));
	ensure (s2.size() == 6);
	ensure (s2 == _T("\"text\""));

}

void
xll_test_attribute(void)
{
	attribute<xchar> a(_T("k1"), _T("v1"));
	ensure (a == _T(" k1=\"v1\""));
}

void
xll_test_tag(void)
{
	tag<xchar> s(_T("tag"));
	ensure (s == _T("<tag"));

	s[attribute<xchar>(_T("k1"), _T("v1"))];
	ensure (s == _T("<tag k1=\"v1\""));

	s[attribute<xchar>(_T("k2"), _T("v2"))];
	ensure (s == _T("<tag k1=\"v1\" k2=\"v2\""));

	etag<xchar> e(_T("tag"));
	ensure (e == _T("</tag>"));
}

void
xll_test_content(void)
{
	content<xchar> c;
	ensure (c.size() == 0);
}

void
xll_test_element(void)
{
	element<xchar> e(
		tag<xchar>(_T("tag"))
			[attribute<xchar>(_T("k1"), _T("v1"))]
			[attribute<xchar>(_T("k2"), _T("v2"))]
		,
		content<xchar>(_T("content"))
	);
	ensure (e == _T("<tag k1=\"v1\" k2=\"v2\">content</tag>"));

	element<xchar> e2(
		tag<xchar>(_T("html"))
		,
		content<xchar>()
			.element(
				tag<xchar>(_T("head"))
				,
				content<xchar>(
					element<xchar>(tag<xchar>(_T("title")), content<xchar>(_T("This is the title.")))
				)
			)
			.element(
				tag<xchar>(_T("body"))
				,
				content<xchar>(_T("This is the body."))
			)
	);
	ensure (e2 == _T("<html><head><title>This is the title.</title></head><body>This is the body.</body></html>"));

}

void
xll_test_temp(xcstr s, xcstr t)
{
	ensure (0 == _tcscmp(s, t));
}

void
xll_test_document(void)
{
	document<xchar> doc(true,
		elements<xchar>(
			tag<xchar>(_T("html"))
			,
			content<xchar>(_T("This is the content."))
		)
		.element(
			tag<xchar>(_T("html2"))
			,
			content<xchar>(_T("This is more content."))
		)
			
	);
	ensure (doc == _T("<?xml version=\"1.1\"?><html>This is the content.</html><html2>This is more content.</html2>"));

	xll_test_temp(document<xchar>(true,
			elements<xchar>(
				tag<xchar>(_T("html"))
				,
				content<xchar>(_T("This is the content."))
			)
			.element(
				tag<xchar>(_T("html2"))
				,
				content<xchar>(_T("This is more content."))
			)
		)
		,
		_T("<?xml version=\"1.1\"?><html>This is the content.</html><html2>This is more content.</html2>")
	);

}

void
xll_test_macros(void)
{
	content<xchar> c(XML_CONTENT(xchar, "content"));
	ensure (c = _T("content"));
	
	tag<xchar> t(_T("name"));
	t XML_ATTRIBUTE(xchar, "k1", "v1");
	ensure (t == _T("<name k1=\"v1\""));
	t XML_ATTRIBUTE(xchar, "k2", "v2");
	ensure (t == _T("<name k1=\"v1\" k2=\"v2\""));

	element<xchar> e(XML_ELEMENT(xchar, "name", XML_ATTRIBUTE(xchar, "k", "v"), XML_CONTENT(xchar, "content")));
	ensure (e == _T("<name k=\"v\">content</name>"));

	document<xchar> doc(
		XML_DOCUMENT(xchar, false,
			XML_ELEMENT(xchar, "name", 
				XML_ATTRIBUTE(xchar, "k1", "v1") /* no comma */
				XML_ATTRIBUTE(xchar, "k2", "v2"), 
				XML_CONTENT(xchar, "content")
			)
		)
	);
	ensure (doc == _T("<name k1=\"v1\" k2=\"v2\">content</name>"));

}

int
xll_test_xml()
{
	try {
		xll_test_string();
		xll_test_attribute();
		xll_test_tag();
		xll_test_content();
		xll_test_element();
		xll_test_document();
		xll_test_macros();
	}
	catch (const std::exception& ex) {
		XLL_ERROR(ex.what());

		return 0;
	}

	return 1;
}
static Auto<Open> xao_test_xml(xll_test_xml);