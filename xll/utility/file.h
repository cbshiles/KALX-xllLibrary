// file.h -  Windows file manipulation
// Copyright (c) 2011 KALX, LLC. All rights reserved. No warranty is made.
#pragma once
#include <cassert>
#include <string>
#define WINDOWS_LEAN_AND_MEAN
#include <windows.h>
#include "xll/utility/ensure.h"

namespace File {

	class Create {
		HANDLE h_;
	public:
		Create()
			: h_(0)
		{ }
		Create(LPCTSTR lpFileName, 
			DWORD dwDesiredAccess = (GENERIC_READ | GENERIC_WRITE),
			DWORD dwCreationDisposition = CREATE_NEW,
			DWORD dwFlagsAndAttributes = FILE_ATTRIBUTE_NORMAL,
			BOOL bInheritHandle = FALSE)
		{
			// Set up the security attributes struct.
			SECURITY_ATTRIBUTES sa;
			sa.nLength= sizeof(SECURITY_ATTRIBUTES);
			sa.lpSecurityDescriptor = NULL;
			sa.bInheritHandle = bInheritHandle;

			h_ = CreateFile(lpFileName, dwDesiredAccess, 0, &sa, dwCreationDisposition, dwFlagsAndAttributes, 0);

			ensure (h_ != INVALID_HANDLE_VALUE);
		}
		Create(const Create& h)
		{
			CloseHandle(h_);
			ensure (DuplicateHandle(GetCurrentProcess(), h.h_, GetCurrentProcess(), &h_, 0, FALSE, DUPLICATE_SAME_ACCESS));
		}
		Create& operator=(const Create& h)
		{
			if (this != &h) {
				CloseHandle(h_);
				ensure (DuplicateHandle(GetCurrentProcess(), h.h_, GetCurrentProcess(), &h_, 0, FALSE, DUPLICATE_SAME_ACCESS));
			}

			return *this;
		}
		~Create()
		{
			CloseHandle(h_);
		}
		operator HANDLE()
		{
			return h_;
		}
		BOOL Write(LPCTSTR lpBuffer, DWORD dwLen = 0)
		{
			DWORD dwOut;

			if (!dwLen)
				dwLen = static_cast<DWORD>(_tcslen(lpBuffer));
			
			return WriteFile(h_, lpBuffer, dwLen, &dwOut, 0);
		}
	
		Create& operator<<(LPCTSTR lpBuffer)
		{
			ensure (Write(lpBuffer, static_cast<DWORD>(_tcslen(lpBuffer))));
			
			return *this;
		}
	
		BOOL Read(std::basic_string<TCHAR>& str)
		{
			BOOL result(FALSE);
			str.erase();
			TCHAR buf[2048];

			DWORD dwRead, dwSize(2048*sizeof(TCHAR));
			do {
				result = ReadFile(h_, buf, dwSize, &dwRead, 0);
				if (!result || dwRead == 0)
					break;
				str.append(buf, dwRead/sizeof(TCHAR));
			} while (dwRead == dwSize);

			return result;
		}
	};

	class Find {
		HANDLE h_;
		WIN32_FIND_DATA fd_;
		Find(const Find&);
		Find& operator=(const Find&);
	public:
		Find(LPCTSTR path)
		{
			ZeroMemory(&fd_, sizeof(fd_));
			h_ = FindFirstFile(path, &fd_);
		}
		~Find()
		{
			FindClose(h_);
		}
		operator const HANDLE&(void) const
		{
			return h_;
		}
		BOOL Next(void)
		{
			return FindNextFile(h_, &fd_);
		}
		const WIN32_FIND_DATA& Data(void) const
		{
			return fd_;
		}
		// full file path
		std::basic_string<TCHAR> FilePath(LPCTSTR path)
		{
			return std::basic_string<TCHAR>(path).append(_T("\\")).append(fd_.cFileName);
		}
	};

} // namespace File

namespace Directory {

	inline BOOL isCurrent(LPCTSTR path)
	{
		return path && path[0] == _T('.') && path[1] == 0;
	}
	inline Bool isParent(LPCTSTR path)
	{
		return path && path[0] == _T('.') && path[1] == _T('.') && path[2] == 0;
	}

	// recusively remove all files and subdirectories
	inline BOOL Remove(LPCTSTR path)
	{
		File::Find ff(std::basic_string<TCHAR>(path).append(_T("\\*")).c_str());

		if (INVALID_HANDLE_VALUE == ff)
			return FALSE;

		do {
			if (ff.Data().dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
				if (!(isCurrent(ff.Data().cFileName) || isParent(ff.Data().cFileName)))
					if (!Remove(ff.FilePath(path).c_str()))
						return FALSE;
			}
			else {
				if (!DeleteFile(ff.FilePath(path).c_str()))
					return FALSE;
			}
		} while (ff.Next());

		return RemoveDirectory(path);
	}

} // namespace Directory
