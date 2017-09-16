// menu.h - menus and toolbars
// Copyright (c) KALX, LLC. All rights reserved. No warranty is made.
/*
static Menu xtb_menu(
	Name("MyName")
	.Item(...)
	.Separator()
	.Subitem(...)
);
*/

#pragma once

namespace xll {

	enum BarID {
		Worksheet4 = 1,   // Worksheet and macro sheet (Excel 4.0)
		Ignore1 = 2,      // Chart (Excel 4.0)
		Null  = 3,        // Null (the menu displayed when no workbooks are open)
		Info = 4,         // Info
		Worksheet3 = 5,   // Worksheet and macro sheet (short menus, Excel 3.0 and earlier)
		Ignore2  = 6,     // Chart (short menus, Excel 3.0 and earlier)
		Cell = 7,         // Cell, toolbar, and workbook (shortcut menus)
		Object  = 8,      // Object (shortcut menus)
		Ignore3 = 9,      // Chart (Excel 4.0 shortcut menus)
		Worksheet  = 10,  // Worksheet and Macro sheet 
		Chart  = 11,      // Chart
		VisualBasic = 12, // Visual Basic
	};

	// built-in shortcut menus
	enum CellMenuId {
		Toolbars = 1,
		Toolbar  = 2,
		Workbook4 = 3,
		SheetCells  = 4,
		Column  = 5,
		Row  = 6,
		WorkbookTite  = 7,
		MacroCells  = 8,
		WorkbookTabs  = 9,
		Desktop  = 10,
		Module = 11,
		Watch  = 12,
		Immediate  = 13,
		Debug  = 14,
	};

	enum ObjectMenuId {
		Drawn  = 1,
		Buttons  = 2,
		Text  = 3,
		DialogSheet  = 4,
	};

	enum ChartMenuId {
		Series = 1,
		Axis = 2,
		Area = 3,
		Entire  = 4,
		Axes = 5,
		Gridlines = 6,
		Floor = 7,
		Legend = 8,
	};

	template<class X>
	struct XAddMenu {
		typedef typename traits<X>::xcstr xcstr;
		typedef typename traits<X>::xchar xchar;
		typedef typename traits<X>::xword xword;
		typedef typename traits<X>::xstring xstring;
		XOPER<X> ref_;
		int bar_id;
		int menu_id;
		XOPER<X> name;
	public:
		XAddMenu(int bid, int mid, xcstr _name)
			: bar_id(bid), menu_id(mid), name(_name)
		{
		}
		XAddMenu& Item(xcstr name, xcstr macro)
		{
			xword r = ref_.rows();

			ref_.resize(r + 1, 2);
			ref_(r, 0) = name;
			if (macro && *macro)
				ref_(r, 1) = macro;

			return *this;
		}
		XAddMenu& Separator()
		{
			xchar hypen('-');

			return Item(XOPER<X>(&hypen, 1), 0);
		}
	};

	template<class X>
	class XMenu {
		const XAddMenu<X> add_;
		XMenu(const XMenu&);
		XMenu& operator=(const XMenu&);
	public:
		XMenu(const XAddMenu<X>& add)
			: add_(add)
		{
			Auto<OpenAfter>([this](void) -> int {
				Excel<X>(xlfAddMenu, XOPER<X>(add_.bar_id), XOPER<X>(add_.menu_id), add_.name);

				return 1;
			});
		}
		~XMenu()
		{
			Excel<X>(xlfDeleteMenu, XOPER<X>(add_.bar_id), XOPER<X>(add_.menu_id));
		}
	};

} // namespace xll