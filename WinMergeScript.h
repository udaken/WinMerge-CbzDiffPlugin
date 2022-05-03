// WinMergeScript.h : Declaration of the CWinMergeScript

#ifndef __WINMERGESCRIPT_H_
#define __WINMERGESCRIPT_H_

#include "stdafx.h"
#include "resource.h"       // main symbols

// change 1 : add this include
#include "typeinfoex.h"
#include "CbzDiffPlugin.h"

/////////////////////////////////////////////////////////////////////////////
// CWinMergeScript

// change 2 : add this
typedef CComTypeInfoHolderModule<1>  CComTypeInfoHolderFileOnly;

class ATL_NO_VTABLE CWinMergeScript :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CWinMergeScript, &CLSID_WinMergeScript>,
	// change 3 : insert the text ", 1, 0, CComTypeInfoHolderFileOnly" 
	// public IDispatchImpl<IWinMergeScript, &IID_IWinMergeScript, &LIBID_CbzDiffPlugin>
	public IDispatchImpl<IWinMergeScript, &IID_IWinMergeScript, &LIBID_CbzDiffPlugin, 1, 0, CComTypeInfoHolderFileOnly>
{
public:
	CWinMergeScript()
	{
	}

	DECLARE_REGISTRY_RESOURCEID(IDR_WINMERGESCRIPT)

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	BEGIN_COM_MAP(CWinMergeScript)
		COM_INTERFACE_ENTRY(IWinMergeScript)
		COM_INTERFACE_ENTRY(IDispatch)
	END_COM_MAP()

	// IWinMergeScript
public:
	STDMETHOD(ShowSettingsDialog)(/*[out, retval]*/ VARIANT_BOOL* pbHandled) override;

	STDMETHOD(get_PluginIsAutomatic)(/*[out, retval]*/ VARIANT_BOOL* pVal) override
	{
		*pVal = VARIANT_TRUE;
		return S_OK;
	}
	STDMETHOD(get_PluginFileFilters)(/*[out, retval]*/ BSTR* pVal) override
	{
		*pVal = SysAllocString(LR"(\.cbz$;\.cbr$;\.cb7$)");
		return S_OK;
	}
	STDMETHOD(get_PluginDescription)(/*[out, retval]*/ BSTR* pVal) override
	{
		*pVal = SysAllocString(L"Compare .CBZ/.CBR as Image");
		return S_OK;
	}

	STDMETHOD(get_PluginEvent)(/*[out, retval]*/ BSTR* pVal) override
	{
		*pVal = SysAllocString(L"FILE_PACK_UNPACK");
		return S_OK;
	}

	STDMETHOD(get_PluginUnpackedFileExtension)(/* [retval][out] */ BSTR* pVal) override
	{
		*pVal = SysAllocString(L".png");
		return S_OK;
	}

	STDMETHOD(get_PluginExtendedProperties)(/* [retval][out] */ BSTR* pVal) override
	{
		*pVal = SysAllocString(L"ProcessType=Content Extraction;FileType=CBZ/CBR/CB7;MenuCaption=CBZ/CBR/CB7");
		return S_OK;
	}
	HRESULT STDMETHODCALLTYPE UnpackFile(
		/* [in] */ BSTR fileSrc,
		/* [in] */ BSTR fileDst,
		VARIANT_BOOL* pbChanged,
		INT* pSubcode,
		/* [retval][out] */ VARIANT_BOOL* pbSuccess) override;

	HRESULT STDMETHODCALLTYPE PackFile(
		/* [in] */ BSTR fileSrc,
		/* [in] */ BSTR fileDst,
		VARIANT_BOOL* pbChanged,
		INT pSubcode,
		/* [retval][out] */ VARIANT_BOOL* pbSuccess) override
	{
		*pbSuccess = VARIANT_FALSE;
		return S_OK;
	}
};

extern INT_PTR CALLBACK SettingDlgProc(HWND hWnd, UINT uiMsg, WPARAM wParam, LPARAM lParam) noexcept;

#endif //__WINMERGESCRIPT_H_