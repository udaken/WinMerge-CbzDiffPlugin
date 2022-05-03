// ConsoleApplication1.cpp : このファイルには 'main' 関数が含まれています。プログラム実行の開始と終了がそこで行われます。
//

#include <iostream>
#include <assert.h>
#include <windows.h>
#include <comip.h>
#include <comdef.h>
#include <comdefsp.h>
#include "../CbzDiffPlugin.h"

#define CHECK(_a) { HRESULT hr = _a; assert(SUCCEEDED(hr) );  }

_COM_SMARTPTR_TYPEDEF(IWinMergeScript, __uuidof(IWinMergeScript));

	// 2980B00C-CF8C-405D-8FC9-338FCBCF0B48
static const GUID CLSID_WinMergeScript =
{ 0x2980B00C, 0xCF8C, 0x405D, { 0x8F, 0xC9, 0x33, 0x8F, 0xCB, 0xCF, 0x0B, 0x48 } };

int main()
{
	HeapSetInformation(NULL, HeapEnableTerminationOnCorruption, NULL, 0);

	// std::wcoutのロケールを設定
	std::wcout.imbue(std::locale("", std::locale::ctype));

	// COM初期化
	CHECK(::CoInitializeEx(nullptr, 0));
	auto hMod = LoadLibrary(L"CbzDiffPlugin.dll");
	auto fpDllGetClassObject = (decltype(&DllGetClassObject))GetProcAddress(hMod, "DllGetClassObject");

	IClassFactoryPtr factory;
	CHECK(fpDllGetClassObject(CLSID_WinMergeScript, IID_PPV_ARGS(&factory)));
	IDispatchPtr pdisp;
	CHECK(factory->CreateInstance(nullptr, IID_PPV_ARGS(&pdisp)));

	IWinMergeScriptPtr plugin = pdisp;
	VARIANT_BOOL bChanged;
	INT subCode{};
	VARIANT_BOOL bSucceed{};

	CHECK(plugin->ShowSettingsDialog(&bSucceed));
	CHECK(plugin->UnpackFile(SysAllocString(L"test.cbz"), SysAllocString(L"result.png"), &bChanged, &subCode, &bSucceed));


}
