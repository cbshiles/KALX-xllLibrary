// args.h - Excel add-in arguments.
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include <string>
#include <vector>
#include "xll/hash.h"
#include "xll/oper.h"

namespace xll {

	// ???Use 'X' for handle and change when registering???
	// Friendly name from argument type.
	template<class X>
	inline typename xll::traits<X>::xcstr ArgType(const XOPER<X>& type)
	{
		ensure (type.xltype == xltypeStr);
		ensure (type.val.str[0] > 0);

		switch (type.val.str[1]) {
			case _T('A') : return _T("Bool");
			case _T('B') : return _T("Num");
			case _T('C') : return _T("Str");
			case _T('D') : return _T("Str");
			case _T('E') : return _T("Num");
			case _T('F') : return _T("Str");
			case _T('G') : return _T("Str");
			case _T('H') : return _T("Int");
			case _T('I') : return _T("Short");
			case _T('J') : return _T("Long");
			case _T('K') : return _T("Array");
			case _T('L') : return _T("Bool");
			case _T('M') : return _T("Short");
			case _T('N') : return _T("Long");
			case _T('P') : return _T("Range");
			case _T('R') : return _T("Reference");
			// not used!!!
			case _T('!') : return _T("Volatile");
			case _T('#') : return _T("Uncalced");
			case _T('$') : return _T("ThreadSafe");
			case _T('&') : return _T("ClusterSafe");
		}

		return _T("???Unknown???");
	}

	template<class X>
	inline typename xll::traits<X>::xcstr ArgTypeJava(const XOPER<X>& type)
	{
		ensure (type.xltype == xltypeStr);
		ensure (type.val.str[0] > 0);

		switch (type.val.str[1]) {
			case _T('A') : return _T("boolean");
			case _T('B') : return _T("double");
			case _T('C') : return _T("String");
			case _T('D') : return _T("String");
			case _T('E') : return _T("double");
			case _T('F') : return _T("String");
			case _T('G') : return _T("String");
			case _T('H') : return _T("int"); //!!!unsigned short, but Java does not support
			case _T('I') : return _T("short");
			case _T('J') : return _T("long");
			case _T('K') : return _T("double[]");
			case _T('L') : return _T("boolean");
			case _T('M') : return _T("short");
			case _T('N') : return _T("long");
			case _T('P') : return _T("Range");
			case _T('R') : return _T("Reference");
			// not used!!!
			case _T('!') : return _T("Volatile");
			case _T('#') : return _T("Uncalced");
			case _T('$') : return _T("ThreadSafe");
			case _T('&') : return _T("ClusterSafe");
		}

		return _T("???Unknown???");
	}

	// Single argument to Excel add-in.
	template<class X>
	class XArg {
		typedef typename xll::traits<X>::xcstr xcstr;
		typedef typename xll::traits<X>::xchar xchar;

		XOPER<X> type_, name_, help_;
	public:
		XArg(xcstr type, xcstr name, xcstr help)
			: type_(type), name_(name), help_(help)
		{ }
		const XOPER<X>& Type(void) const
		{
			return type_;
		}
		const XOPER<X>& Name(void) const
		{
			return name_;
		}
		const XOPER<X>& Help(void) const
		{
			return help_;
		}
	};

	typedef XArg<XLOPER>   Arg;
	typedef XArg<XLOPER12> Arg12;
	typedef X_(Arg)        ArgX;


	// Build up arguments for call to xlfRegister.
	template<class X>
	class XArgs {
		typedef typename xll::traits<X>::xcstr xcstr;
		typedef typename xll::traits<X>::xchar xchar;

		std::vector<XArg<X> > args_;
		mutable XOPER<X> arg_; // to be passed to xlfRegister
		std::basic_string<xchar> doc_;
		unsigned int tid_; // help file topic id
	public:
		XArgs()
			: tid_(0)
		{
		}
		// documentation
		XArgs(xcstr category)
			: arg_(1, 7), tid_(utility::hash(category, 0, true))
		{
			arg_[5] = -1;
			arg_[6] = category;
		}
		// C function and Excel name
		XArgs(xcstr procedure, xcstr function_text)
			: arg_(1, 6), tid_(0)
		{
			arg_[1] = procedure;
			arg_[3] = function_text;
			arg_[5] = 2; // macro
		}
		// C function, C signature, Excel name and argument prompt
		XArgs(xcstr procedure, xcstr type_text, xcstr function_text, xcstr argument_text,
			xcstr category = traits<X>::null(), xcstr function_help = traits<X>::null())
				: arg_(1, 10), tid_(0)
		{
			arg_[1] = procedure;
			arg_[2] = type_text;
			arg_[3] = function_text;
			arg_[4] = argument_text;
			arg_[5] = 1; // function
			if (category)
				arg_[6] = category;
			if (function_help)
				arg_[9] = function_help;
		}
		// return type, C name, Excel name
		XArgs(xcstr type_text, xcstr procedure, xcstr function_text)
			: arg_(1, 10), tid_(0)
		{
			arg_[1] = procedure;
			arg_[2] = type_text;
			arg_[3] = function_text;
			arg_[4] = traits<X>::null();
			arg_[5] = 1; // function
		}
		XArgs(const XArgs& arg)
			: args_(arg.args_), arg_(arg.arg_), doc_(arg.doc_), tid_(arg.tid_)
		{
		}
		XArgs& operator=(const XArgs& arg)
		{
			if (this != &arg) {
				args_ = arg.args_;
				arg_ = arg.arg_;
				doc_ = arg.doc_;
				tid_ = arg.tid_;
			}
		
			return *this;
		}
		~XArgs()
		{
		}

		bool isFunction(void) const
		{
			return arg_[5] == 1;
		}
		bool isMacro(void) const
		{
			return arg_[5] == 2;
		}
		bool isDocument(void) const
		{
			return arg_[5] == -1;
		}
		unsigned int TopicId(void) const
		{
			return tid_;
		}

		// Tack on an argument
		XArgs& Arg(xcstr type, xcstr name, xcstr help)
		{
			arg_[2].append(type);

			if (ArgumentText())
				arg_[4].append(_T(", "));
			arg_[4].append(name);
			
			ArgumentHelp(help);

			args_.push_back(XArg<X>(type, name, help));

			return *this;
		}
		size_t ArgSize(void) const
		{
			return args_.size();
		}
		// 1-based individual argument
		const XArg<X>& Arg(size_t i) const
		{
			ensure (i != 0);

			return args_[i - 1];
		}
		//
		// Accessors
		//

		// C function name
		const XOPER<X>& Procedure() const
		{
			return arg_[1];
		}
		XArgs& Procedure(xcstr procedure)
		{
			arg_[1] = procedure;

			return *this;
		}
		XArgs& Procedure(const X& procedure)
		{
			arg_[1] = procedure;

			return *this;
		}

		// C signature
		const OPERX& TypeText() const
		{
			return arg_[2];
		}
		XArgs& TypeText(xcstr type_text)
		{
			arg_[2] = type_text;

			return *this;
		
		}
		XArgs& TypeText(const X& type_text)
		{
			arg_[2] = type_text;

			return *this;
		
		}

		// Excel function name
		const OPERX& FunctionText() const
		{
			return arg_[3];
		}
		XArgs& FunctionText(xcstr function_text)
		{
			arg_[3] = function_text;

			return *this;
		}
		XArgs& FunctionText(const X& function_text)
		{
			arg_[3] = function_text;

			return *this;
		}

		// Excel Control-Shift-A argument prompt
		const OPERX& ArgumentText() const
		{
			return arg_[4];
		}
		XArgs& ArgumentText(xcstr argument_text)
		{
			arg_[4] = argument_text;

			return *this;
		}
		XArgs& ArgumentText(const X& argument_text)
		{
			arg_[4] = argument_text;

			return *this;
		}

		// 0 - hidden function, 1 - function, 2 - macro
		const OPERX& MacroType() const
		{
			return arg_[5];
		}
		XArgs& MacroType(int macro_type)
		{
			arg_[5] = macro_type;

			return *this;
		}
		XArgs& MacroType(xcstr macro_type)
		{
			arg_[5] = macro_type;

			return *this;
		}
		XArgs& MacroType(const X& macro_type)
		{
			arg_[5] = macro_type;

			return *this;
		}

		// Function Wizard category
		const OPERX& Category() const
		{
			return arg_[6];
		}
		XArgs& Category(xcstr category)
		{
			arg_[6] = category;

			return *this;
		}

		// CONTROL+SHIFT+shorcut_text single character macro shortcut
		const OPERX& ShortcutText() const
		{
			return arg_[7];
		}
		XArgs& ShortcutText(xcstr shortcut_text)
		{
			arg_[7] = shortcut_text;

			return *this;
		}
		XArgs& ShortcutText(const X& shortcut_text)
		{
			arg_[7] = shortcut_text;

			return *this;
		}

		// path\file.chm!help_context_id
		const OPERX& HelpTopic() const
		{
			return arg_[8];
		}
		XArgs& HelpTopic(xcstr help_topic)
		{
			arg_[8] = help_topic;

			return *this;
		}
		XArgs& HelpTopic(const X& help_topic)
		{
			arg_[8] = help_topic;

			return *this;
		}

		// Function description in Function Wizard
		const OPERX& FunctionHelp() const
		{
			return arg_[9];
		}
		XArgs& FunctionHelp(xcstr function_help)
		{
			arg_[9] = function_help;

			return *this;
		}
		XArgs& FunctionHelp(const X& function_help)
		{
			arg_[9] = function_help;

			return *this;
		}

		size_t ArgumentCount(void) const
		{
			return arg_.size() - 9;
		}
		// Individual help in the Function Wizard.

		const OPERX& ArgumentHelp(size_t i) const
		{
			ensure (i > 0);
			ensure (9 + i < arg_.size());

			return arg_[9 + i];
		}
		XArgs& ArgumentHelp(xcstr argument_help, size_t i = 0)
		{
			if (i) {
				if (10 + i >= arg_.size()) {
					arg_.resize(1, 11 + i);
				}
				arg_[10 + i] = argument_help;
			}
			else {
				arg_.push_back(XStr<X>(argument_help));
			}

			return *this;
		}
		XArgs& ArgumentHelp(const X& argument_help, size_t i = 0)
		{
			if (i) {
				if (10 + i >= arg_.size()) {
					arg_.resize(1, 11 + i);
				}
				arg_[10 + i] = argument_help;
			}
			else {
				arg_.push_back(argument_help);
			}

			return *this;
		}

		// Additional documentation for help file.
		xcstr Documentation(void) const
		{
			return doc_.c_str();
		}
		// Only generate documentation if this is called.
		XArgs& Documentation(xcstr doc = traits<X>::null())
		{
			doc_ = doc;

			if (!tid_) {
				LPCTSTR ft = FunctionText().val.str;
				tid_ = utility::hash(ft + 1, ft[0], true);
			}
			
			return *this;
		}

		// arguments xlfRegister needs
		int size() const
		{
			return static_cast<int>(arg_.size());
		}
		typename const traits<X>::pointer_type* pointers() const
		{
			ensure (size() < limits<X>::maxargs);
			static traits<X>::pointer_type p[limits<X>::maxargs]; // It's not as bad as it looks.
			
			p[0] = &GetName(); // have to wait for xlAutoOpen to be called
			if (tid_) {
				XOPER<X> topic(GetName());
				// replace xll by chm
				traits<X>::strncpy(topic.val.str + topic.val.str[0] - 2, traits<X>::chm(), 3);
				topic = Excel<X>(xlfConcatenate, topic, XOPER<X>(_T("!")), XOPER<X>(tid_));
				arg_[8] = topic;
			}
			for (int i = 1; i < size(); ++i)
				p[i] = arg_.val.array.lparray + i;

			return p;
		}
		static X& GetName(void) {
			static LXOPER<X> o_(Excel<X>(xlGetName));

			return o_;
		}
	private:
		bool CheckArgs()
		{
			int arity = arg_size() - 10;
			// parse type_text!!!
		}
	};

	typedef XArgs<XLOPER>   Args;
	typedef XArgs<XLOPER12> Args12;
	typedef X_(Args)        ArgsX;

	typedef XArgs<XLOPER>   Function;
	typedef XArgs<XLOPER12> Function12;
	typedef X_(Args)        FunctionX;

} // namespace xll