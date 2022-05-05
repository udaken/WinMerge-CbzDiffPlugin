#include "stdafx.h"
#include "WinMergeScript.h "
#include "Config.h"
#include "resource.h"
#include <commctrl.h>
#include <algorithm>

CString KeyName()
{
	CString baseKey(LR"(Software\WinMergePlugins\)");
	WCHAR name[MAX_PATH]{};
	GetModuleFileName(_Module.GetModuleInstance(), name, ARRAYSIZE(name));
	baseKey.Append(PathFindFileName(name));
	return baseKey;
}

std::unique_ptr<Config> Config::Load()
{
	auto p = std::make_unique<Config>();
	Config& config = *p;
	DWORD type{ REG_DWORD };
	DWORD cbData{ sizeof(DWORD) };
	auto keyName = KeyName();
	SHGetValue(HKEY_CURRENT_USER, keyName, key_Scale, &type, &config.m_scale, &cbData);
	config.m_scale = std::clamp(config.m_scale, 1UL, 100UL);
	SHGetValue(HKEY_CURRENT_USER, keyName, key_FileMode, &type, &config.m_Mode, &cbData);
	SHGetValue(HKEY_CURRENT_USER, keyName, key_ForceD2D1, &type, &config.m_forceD2D1, &cbData);
	SHGetValue(HKEY_CURRENT_USER, keyName, key_FileNameOrdering, &type, &config.m_fileNameOrdering, &cbData);
	SHGetValue(HKEY_CURRENT_USER, keyName, key_BackgroundColor, &type, &config.m_backgroundColor, &cbData);

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
