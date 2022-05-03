#include "stdafx.h"
#include "WinMergeScript.h "
#include "Config.h"
#include "resource.h"
#include <commctrl.h>
#include <algorithm>

/**
 * @brief Get the name of the current dll
 */
LPTSTR GetDllFilename(LPTSTR name, int len)
{
	// careful for the last char, the doc does not give this detail
	name[len] = 0;
	GetModuleFileName(_Module.GetModuleInstance(), name, len - 1);
	// find last backslash
	TCHAR* lastslash = _tcsrchr(name, '/');
	if (lastslash == 0)
		lastslash = name;
	else
		lastslash++;
	TCHAR* lastslash2 = _tcsrchr(lastslash, '\\');
	if (lastslash2 == 0)
		lastslash2 = name;
	else
		lastslash2++;
	if (lastslash2 != name)
		lstrcpy(name, lastslash2);
	return name;
}

CString KeyName()
{
	CString baseKey(L"Software\\WinMergePlugins\\");
	WCHAR name[MAX_PATH]{};
	GetModuleFileName(_Module.GetModuleInstance(), name, ARRAYSIZE(name));
	baseKey.Append(PathFindFileName(name));
	return baseKey;
}

CString RegReadString(const CString& key, const CString& valuename, const CString& defaultValue)
{
	CRegKey reg;
	if (reg.Open(HKEY_CURRENT_USER, key, KEY_READ) == ERROR_SUCCESS)
	{
		TCHAR value[512] = { 0 };
		DWORD dwSize = sizeof(value) / sizeof(value[0]);
		reg.QueryStringValue(valuename, value, &dwSize);
		return value;
	}
	return defaultValue;
}

std::unique_ptr<Config> Config::Load()
{
	auto p = std::make_unique<Config>();
	Config& config = *p;
	DWORD type{ REG_DWORD };
	DWORD cbData{ sizeof(DWORD) };
	SHGetValue(HKEY_CURRENT_USER, KeyName(), key_Scale, &type, &config.m_scale, &cbData);
	config.m_scale = std::clamp(config.m_scale, 1UL, 100UL);
	SHGetValue(HKEY_CURRENT_USER, KeyName(), key_ForceD2D1, &type, &config.m_forceD2D1, &cbData);
	SHGetValue(HKEY_CURRENT_USER, KeyName(), key_FileNameOrdering, &type, &config.m_fileNameOrdering, &cbData);

	return p;
}

void SHSetDword(LPCWSTR key, DWORD value)
{
	SHSetValueW(HKEY_CURRENT_USER, KeyName(), key, REG_DWORD, &value, sizeof(value));
}

INT_PTR CALLBACK SettingDlgProc(HWND hWnd, UINT uiMsg, WPARAM wParam, LPARAM lParam) noexcept
{
	switch (uiMsg) {
	case WM_INITDIALOG:
	{
		auto config = Config::Load();
		SendDlgItemMessage(hWnd, IDC_SLIDER1, TBM_SETTICFREQ, 10, 0);
		SendDlgItemMessage(hWnd, IDC_SLIDER1, TBM_SETRANGE, FALSE, MAKELPARAM(1, 100));
		SendDlgItemMessage(hWnd, IDC_SLIDER1, TBM_SETPOS, TRUE, config->scaleInt());
		SetDlgItemInt(hWnd, IDC_EDIT1, config->scaleInt(), FALSE);
		return TRUE;
	}
	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK)
		{
			SHSetDword(Config::key_Scale, SendDlgItemMessage(hWnd, IDC_SLIDER1, TBM_GETPOS, 0, 0));
			EndDialog(hWnd, IDOK);
		}
		else if (LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hWnd, IDCANCEL);
		}
		else if (LOWORD(wParam) == IDC_BUTTON1)
		{
			SHDeleteKey(HKEY_CURRENT_USER, KeyName());
		}
		return TRUE;
	case WM_HSCROLL:
		if (GetDlgItem(hWnd, IDC_SLIDER1) == (HWND)lParam)
		{
			auto pos = SendDlgItemMessage(hWnd, IDC_SLIDER1, TBM_GETPOS, 0, 0);
			SetDlgItemInt(hWnd, IDC_EDIT1, pos, FALSE);
			return TRUE;
		}
		break;
	default:
		break;
	}
	return FALSE;
}