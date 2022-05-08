// WinMergeScript.h : Declaration of the CWinMergeScript

#ifndef __WINMERGESCRIPT_H_
#define __WINMERGESCRIPT_H_

#include "resource.h" // main symbols
#include "stdafx.h"

// change 1 : add this include
#include "CbzDiffPlugin.h"
#include "typeinfoex.h"

/////////////////////////////////////////////////////////////////////////////
// CWinMergeScript

// change 2 : add this
typedef CComTypeInfoHolderModule<1> CComTypeInfoHolderFileOnly;

class ATL_NO_VTABLE CWinMergeScript
    : public CComObjectRootEx<CComSingleThreadModel>,
      public CComCoClass<CWinMergeScript, &CLSID_WinMergeScript>,
      // change 3 : insert the text ", 1, 0, CComTypeInfoHolderFileOnly"
      // public IDispatchImpl<IWinMergeScript, &IID_IWinMergeScript, &LIBID_CbzDiffPlugin>
      public IDispatchImpl<IWinMergeScript, &IID_IWinMergeScript, &LIBID_CbzDiffPlugin, 1, 0,
                           CComTypeInfoHolderFileOnly>
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
    IFACEMETHODIMP ShowSettingsDialog(/*[out, retval]*/ VARIANT_BOOL *pbHandled) override;

    IFACEMETHODIMP get_PluginIsAutomatic(/*[out, retval]*/ VARIANT_BOOL *pVal) override
    {
        *pVal = VARIANT_TRUE;
        return S_OK;
    }
    IFACEMETHODIMP get_PluginFileFilters(/*[out, retval]*/ BSTR *pVal) override
    {
        *pVal = SysAllocString(LR"(\.pdf$;\.cbz$;\.cbr$;\.cb7$)");
        return S_OK;
    }
    IFACEMETHODIMP get_PluginDescription(/*[out, retval]*/ BSTR *pVal) override
    {
        *pVal = SysAllocString(L"Compare .PDF/.CBZ/.CBR/.CB7 as Image");
        return S_OK;
    }

    IFACEMETHODIMP get_PluginEvent(/*[out, retval]*/ BSTR *pVal) override;

    IFACEMETHODIMP get_PluginUnpackedFileExtension(/* [retval][out] */ BSTR *pVal) override;

    IFACEMETHODIMP get_PluginExtendedProperties(/* [retval][out] */ BSTR *pVal) override
    {
        *pVal = SysAllocString(L"ProcessType=Content Extraction;FileType=PDF/CBZ/CBR/CB7;MenuCaption=PDF/CBZ/CBR/CB7");
        return S_OK;
    }

    IFACEMETHODIMP UnpackBufferA(
        /* [in] */ SAFEARRAY **pBuffer,
        /* [in] */ INT *pSize,
        /* [in] */ VARIANT_BOOL *pbChanged,
        /* [in] */ INT *pSubcode,
        /* [retval][out] */ VARIANT_BOOL *pbSuccess) override;

    IFACEMETHODIMP PackBufferA(
        /* [in] */ SAFEARRAY **pBuffer,
        /* [in] */ INT *pSize,
        /* [in] */ VARIANT_BOOL *pbChanged,
        /* [in] */ INT subcode,
        /* [retval][out] */ VARIANT_BOOL *pbSuccess) override
    {
        *pbSuccess = VARIANT_FALSE;
        return S_OK;
    }

    IFACEMETHODIMP UnpackFile(
        /* [in] */ BSTR fileSrc,
        /* [in] */ BSTR fileDst, VARIANT_BOOL *pbChanged, INT *pSubcode,
        /* [retval][out] */ VARIANT_BOOL *pbSuccess) override;

    IFACEMETHODIMP PackFile(
        /* [in] */ BSTR fileSrc,
        /* [in] */ BSTR fileDst, VARIANT_BOOL *pbChanged, INT pSubcode,
        /* [retval][out] */ VARIANT_BOOL *pbSuccess) override
    {
        *pbSuccess = VARIANT_FALSE;
        return S_OK;
    }

    IFACEMETHODIMP IsFolder(
        /* [in] */ BSTR file,
        /* [retval][out] */ VARIANT_BOOL *pbSuccess) override
    {
        *pbSuccess = VARIANT_TRUE; // always true
        return S_OK;
    }

    IFACEMETHODIMP UnpackFolder(
        /* [in] */ BSTR fileSrc,
        /* [in] */ BSTR folderDst, VARIANT_BOOL *pbChanged, INT *pSubcode,
        /* [retval][out] */ VARIANT_BOOL *pbSuccess) override;

    IFACEMETHODIMP PackFolder(
        /* [in] */ BSTR fileSrc,
        /* [in] */ BSTR fileDst, VARIANT_BOOL *pbChanged, INT pSubcode,
        /* [retval][out] */ VARIANT_BOOL *pbSuccess) override
    {
        *pbSuccess = VARIANT_FALSE;
        return S_OK;
    }
};

extern INT_PTR CALLBACK SettingDlgProc(HWND hWnd, UINT uiMsg, WPARAM wParam, LPARAM lParam) noexcept;

#endif //__WINMERGESCRIPT_H_
