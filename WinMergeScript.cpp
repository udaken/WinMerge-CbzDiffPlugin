// WinMergeScript.cpp : Implementation of CWinMergeScript
#include "stdafx.h"
#pragma warning(disable: 28251)
#include "CbzDiffPlugin.h"
#include "WinMergeScript.h"
#include "InFileStream.h"
#include "resource.h"
#include "Config.h"

#include <vector>
#include <algorithm>
#include <string>
#include <string_view>
#include <assert.h>

#pragma comment(lib, "d3d11.lib")
#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "dxguid.lib")

#ifdef NDEBUG
#define CHECK_return(hr) do{ HRESULT _hr = hr; if(FAILED(_hr )) { return _hr;}}while(0)
#else
#define CHECK_return(hr) do{ HRESULT _hr = hr; assert(SUCCEEDED(_hr)); if(FAILED(_hr )) { return _hr;} }while(0)
#endif

using namespace std::string_literals;
using namespace std::string_view_literals;

EXTERN_C const GUID FAR CLSID_CFormat7z;
EXTERN_C const GUID FAR CLSID_CFormatZip;
EXTERN_C const GUID FAR CLSID_CFormatRar;

const REFWICPixelFormatGUID InternalWICFormat = GUID_WICPixelFormat32bppPBGRA;

class PropVarinat final : public PROPVARIANT
{
public:
	PropVarinat(const PropVarinat&) = delete;
	PropVarinat& operator=(const PropVarinat&) = delete;

	PropVarinat() : PROPVARIANT{}
	{

	}
	~PropVarinat()
	{
		PropVariantClear(this);
	}
};

class StreamWrapper final
	: public ISequentialOutStream
{
	CComPtr< IStream> const baseStream;
	long m_Ref{ 1 };

	explicit StreamWrapper(IStream* s) : baseStream(s)
	{
	}
public:
	StreamWrapper(const StreamWrapper&) = delete;
	StreamWrapper& operator=(const StreamWrapper&) = delete;

	static StreamWrapper* Create(IStream* s)
	{
		return new(std::nothrow) StreamWrapper(s);
	}

	STDMETHODIMP QueryInterface(REFIID riid, void** ppvObject) override
	{
		static const QITAB qit[] =
		{
			{ &IID_ISequentialOutStream, OFFSETOFCLASS(ISequentialOutStream, StreamWrapper)},
			{ 0 },
		};
		return QISearch(this, qit, riid, ppvObject);
	}
	STDMETHODIMP_(ULONG) AddRef(void) override
	{
		return InterlockedIncrement(&m_Ref);
	}
	STDMETHODIMP_(ULONG) Release(void) override
	{
		ULONG count = InterlockedDecrement(&m_Ref);
		if (count == 0)
		{
			delete this;
			return 0;
		}
		return count;
	}
	STDMETHOD(Write)(const void* data, UInt32 size, UInt32* processedSize) override
	{
		return baseStream->Write(data, static_cast<ULONG>(size), reinterpret_cast<ULONG*>(processedSize));
	}

};

inline HRESULT WriteToStream(IWICImagingFactory* factory, IWICBitmapSource* srcImage, IStream* piStream)
{
	HRESULT hr;

	for (auto& container : { GUID_ContainerFormatPng, GUID_ContainerFormatJpeg })
	{
		CComPtr<IWICBitmapEncoder> pngEncoder;
		hr = factory->CreateEncoder(container, nullptr, &pngEncoder);
		if (FAILED(hr)) continue;
		hr = pngEncoder->Initialize(piStream, WICBitmapEncoderNoCache);
		if (FAILED(hr)) continue;
		CComPtr< IWICBitmapFrameEncode> piBitmapFrame;
		CComPtr<IPropertyBag2> pPropertybag;
		hr = pngEncoder->CreateNewFrame(&piBitmapFrame, &pPropertybag);
		if (FAILED(hr)) continue;
		hr = piBitmapFrame->Initialize(pPropertybag);
		GUID format = GUID_WICPixelFormat32bppPBGRA;
		hr = piBitmapFrame->SetPixelFormat(&format);
		if (FAILED(hr)) continue;
		hr = piBitmapFrame->WriteSource(srcImage, nullptr);
		if (FAILED(hr)) continue;
		hr = piBitmapFrame->Commit();
		if (FAILED(hr)) continue;
		hr = pngEncoder->Commit();
		if (FAILED(hr)) continue;
		hr = piStream->Commit(STGC_DEFAULT);
		if (FAILED(hr)) continue;

		if (SUCCEEDED(hr))
			return hr;
	}

	return hr;
}

inline LPCWSTR GetPixelFormatName(WICPixelFormatGUID guid)
{
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormatDontCare))       return L"GUID_WICPixelFormatDontCare";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat1bppIndexed))    return L"GUID_WICPixelFormat1bppIndexed";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat2bppIndexed))    return L"GUID_WICPixelFormat2bppIndexed";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat4bppIndexed))    return L"GUID_WICPixelFormat4bppIndexed";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat8bppIndexed))    return L"GUID_WICPixelFormat8bppIndexed";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormatBlackWhite))     return L"GUID_WICPixelFormatBlackWhite";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat2bppGray))       return L"GUID_WICPixelFormat2bppGray";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat4bppGray))       return L"GUID_WICPixelFormat4bppGray";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat8bppGray))       return L"GUID_WICPixelFormat8bppGray";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat8bppAlpha))      return L"GUID_WICPixelFormat8bppAlpha";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat16bppBGR555))    return L"GUID_WICPixelFormat16bppBGR555";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat16bppBGR565))    return L"GUID_WICPixelFormat16bppBGR565";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat16bppBGRA5551))  return L"GUID_WICPixelFormat16bppBGRA5551";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat16bppGray))      return L"GUID_WICPixelFormat16bppGray";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat24bppBGR))       return L"GUID_WICPixelFormat24bppBGR";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat24bppRGB))       return L"GUID_WICPixelFormat24bppRGB";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppBGR))       return L"GUID_WICPixelFormat32bppBGR";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppBGRA))      return L"GUID_WICPixelFormat32bppBGRA";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppPBGRA))     return L"GUID_WICPixelFormat32bppPBGRA";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppGrayFloat)) return L"GUID_WICPixelFormat32bppGrayFloat";

	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppRGBA))     return L"GUID_WICPixelFormat32bppRGBA";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppPRGBA))    return L"GUID_WICPixelFormat32bppPRGBA";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat48bppRGB))      return L"GUID_WICPixelFormat48bppRGB";
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat48bppBGR))      return L"GUID_WICPixelFormat48bppBGR";

	return L"";
}


inline D2D1_PIXEL_FORMAT WicPixelFormatToD2D1(WICPixelFormatGUID guid)
{
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppPRGBA)) return D2D1::PixelFormat(DXGI_FORMAT_R8G8B8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED);
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppBGR)) return D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_IGNORE);
	if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppPBGRA)) return D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED);
	return D2D1::PixelFormat();
}

inline HRESULT ScaledSize(IWICBitmapSource* srcImage, float scale, UINT& width, UINT& height)
{
	HRESULT hr;
	hr = srcImage->GetSize(&width, &height);
	CHECK_return(hr);

	width = static_cast<UINT>(width * scale);
	height = static_cast<UINT>(height * scale);

	return hr;
}

inline HRESULT ScaleImage(IWICImagingFactory* factory, IWICBitmapSource* srcImage, float scale, IWICBitmapSource** result)
{
	CComPtr<IWICBitmapScaler> scaler;
	HRESULT hr;
	hr = factory->CreateBitmapScaler(&scaler);
	CHECK_return(hr);

	UINT width{}, height{};
	hr = ScaledSize(srcImage, scale, width, height);
	CHECK_return(hr);
	hr = scaler->Initialize(srcImage, width, height, WICBitmapInterpolationModeHighQualityCubic);
	CHECK_return(hr);

	CComPtr<IWICFormatConverter> converter;
	hr = factory->CreateFormatConverter(&converter);
	CHECK_return(hr);
	hr = converter->Initialize(scaler, InternalWICFormat, WICBitmapDitherTypeNone, nullptr, 0.f, WICBitmapPaletteTypeMedianCut);
	CHECK_return(hr);

	hr = converter.QueryInterface(result);
	CHECK_return(hr);
	return hr;
}

inline HRESULT GetSupportedExtentions(IWICImagingFactory* factory, std::wstring& result)
{
	CComPtr<IEnumUnknown> pEnum;
	HRESULT hr;
	hr = factory->CreateComponentEnumerator(WICDecoder, WICComponentEnumerateDefault, &pEnum);
	CHECK_return(hr);

	IUnknown* pComponent{};
	ULONG num{};
	while (S_OK == pEnum->Next(1, &pComponent, &num) && num == 1)
	{
		CComPtr<IWICBitmapDecoderInfo> pDecoderInfo;
		hr = pComponent->QueryInterface(&pDecoderInfo);
		CHECK_return(hr);
		pComponent->Release();
		pComponent = nullptr;

		UINT length{};
		hr = pDecoderInfo->GetFileExtensions(0, nullptr, &length);
		CHECK_return(hr);

		std::wstring ext(static_cast<size_t>(length - 1), L'\0');

		hr = pDecoderInfo->GetFileExtensions(length, ext.data(), &length);
		CHECK_return(hr);
		if (not result.empty())
			result.append(L",");

		result.append(ext);
	}
	std::replace(result.begin(), result.end(), L',', L';');

	constexpr auto target{ L"."sv };
	constexpr auto replacement{ L"*."sv };

	std::string::size_type pos = 0;
	while ((pos = result.find(target, pos)) != std::string::npos) {
		result.replace(pos, target.length(), replacement);
		pos += replacement.length();
	}

	return hr;
}

inline HRESULT Load7zDll(HMODULE& h7z)
{
	WCHAR moduleFileName[MAX_PATH]{};
	if (GetModuleFileName(nullptr, moduleFileName, ARRAYSIZE(moduleFileName)))
	{
		// relative path from WinMergeU.exe
		PathAppend(moduleFileName, LR"(\..\Merge7z\7z.dll)");
		h7z = LoadLibrary(moduleFileName);
		if (h7z)
			return S_OK;
	}

	if (not GetModuleFileName(_Module.GetModuleInstance(), moduleFileName, ARRAYSIZE(moduleFileName)))
	{
		return HRESULT_FROM_WIN32(GetLastError());
	}

	PathAppend(moduleFileName, LR"(\..\7z.dll)");
	h7z = LoadLibrary(moduleFileName);
	if (h7z == nullptr)
		return HRESULT_FROM_WIN32(GetLastError());
	return S_OK;
}

class CbzToImage final : IArchiveExtractCallback
{
	typedef UINT32(WINAPI* CreateObjectFunc)(REFGUID clsID, REFIID interfaceID, void** outObject);

	struct ItemInfo
	{
		UInt32 index;
		CComBSTR name;
	};

	CComPtr<IInArchive> m_pInArchive;
	CComPtr< IWICImagingFactory> m_pWicFactory;
	CComPtr<IWICBitmap> m_canvas;

	HMODULE h7z{};
	CreateObjectFunc CreateObject{};
	std::vector< ItemInfo> m_items;
	std::wstring m_supportedExt;

	Config& m_config;

	ULONG64 m_totalWidth{};
	UINT m_maxWidth{};
	ULONG64 m_totalHeight{};
	UINT m_maxHeight{};

public:
	CbzToImage(const CbzToImage&) = delete;
	CbzToImage& operator=(const CbzToImage&) = delete;

	explicit CbzToImage(Config& config)
		: m_config(config)
	{
	}
	HRESULT Init()
	{
		HRESULT hr;
		hr = Load7zDll(h7z);
		CHECK_return(hr);

		CreateObject = (CreateObjectFunc)GetProcAddress(h7z, "CreateObject");
		if (CreateObject == nullptr)
			return HRESULT_FROM_WIN32(GetLastError());

		hr = m_pWicFactory.CoCreateInstance(CLSID_WICImagingFactory);
		CHECK_return(hr);

		hr = GetSupportedExtentions(m_pWicFactory, m_supportedExt);
		CHECK_return(hr);

		hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_MULTI_THREADED, &m_pD2D1Factory);
		CHECK_return(hr);

		return hr;
	}

	static REFGUID GetFormatGUID(LPCWSTR path)
	{
		if (PathMatchSpecEx(path, L"*.cbz;*.zip", PMSF_MULTIPLE) == S_OK) return CLSID_CFormatZip;
		if (PathMatchSpecEx(path, L"*.cbr;*.rar", PMSF_MULTIPLE) == S_OK) return CLSID_CFormatRar;
		if (PathMatchSpecEx(path, L"*.cb7;*.7z", PMSF_MULTIPLE) == S_OK) return CLSID_CFormat7z;
		return CLSID_NULL;
	}
	static REFGUID GetFormatGUIDFromSig(const void* sig, size_t size)
	{
		if (memcmp(sig, "PK\03\04", sizeof(4)) == 0) return CLSID_CFormatZip;
		if (memcmp(sig, "Rar!", sizeof(4)) == 0) return CLSID_CFormatRar;
		constexpr BYTE magic7z[] = { '7','z', 0xBC, 0xAF, 0x27, 0x1C };
		if (memcmp(sig, magic7z, sizeof(magic7z)) == 0)  return CLSID_CFormat7z;
		return CLSID_NULL;
	}

	HRESULT ProcessFile(LPCWSTR path, LPCWSTR dest)
	{
		HRESULT hr;

		CComPtr<InFileStream> inputStream;
		hr = InFileStream::OpenPath(path, &inputStream);
		CHECK_return(hr);

		CComPtr<IWICStream> piStream;
		hr = m_pWicFactory->CreateStream(&piStream);
		CHECK_return(hr);
		hr = piStream->InitializeFromFilename(dest, GENERIC_WRITE);
		CHECK_return(hr);

		hr = Process(inputStream, GetFormatGUID(path), piStream);

		return hr;
	}

	HRESULT ProcessBuffer(
		/* [in] */ LPSAFEARRAY& pBuffer,
		/* [in] */ INT& pSize)
	{
		HRESULT hr;

		void* sig;
		hr = SafeArrayAccessData(pBuffer, &sig);
		CHECK_return(hr);
		auto guid = GetFormatGUIDFromSig(sig, 4);
		hr = SafeArrayUnaccessData(pBuffer);

		CComPtr<InSafeArrayStream> inputStream;
		hr = InSafeArrayStream::From(std::move(pBuffer), &inputStream);
		CHECK_return(hr);


		CComPtr<ISafeArrayStream> piStream;
		hr = SafeArrayStream::From(nullptr, &piStream);
		CHECK_return(hr);

		hr = Process(inputStream, guid, piStream);
		CHECK_return(hr);

		ULONG size;
		piStream->GetSize(&size);
		pSize = size;

		hr = piStream->Detach(&pBuffer);

		return hr;
	}

	~CbzToImage()
	{
		if (h7z) {
			FreeLibrary(h7z);
			h7z = nullptr;
		}
	}

private:

	HRESULT Process(IInStream* inputStream, REFGUID guid, IStream* outputStream)
	{
		if (guid == IID_NULL)
			return ERROR_INVALID_DATA;

		CComPtr<IInArchive> pInArchive;
		HRESULT hr;
		hr = CreateObject(guid, IID_IInArchive, IID_PPV_ARGS_Helper(&pInArchive));
		CHECK_return(hr);

		hr = pInArchive->Open(inputStream, nullptr, nullptr);
		CHECK_return(hr);

		PropVarinat vSolid{};
		hr = pInArchive->GetArchiveProperty(kpidSolid, &vSolid);
		if (vSolid.boolVal)
			return ERROR_INVALID_DATA;

		UInt32 itemsCount{};
		hr = pInArchive->GetNumberOfItems(&itemsCount);
		CHECK_return(hr);
		m_items.reserve(itemsCount);

		for (UInt32 index = 0; index < itemsCount; ++index)
		{
			PropVarinat vPath{};
			hr = pInArchive->GetProperty(index, kpidPath, &vPath);
			CHECK_return(hr);
			CComBSTR name;
			hr = PropVariantToBSTR(vPath, &name);
			CHECK_return(hr);

			if (S_OK == PathMatchSpecExW(name, m_supportedExt.c_str(), PMSF_MULTIPLE))
			{
				m_items.emplace_back(index, name);
			}
		}

		SortItems();

		std::vector<UInt32> indics;
		indics.reserve(static_cast<size_t>(itemsCount));
		for (auto& i : m_items)
		{
			indics.push_back(i.index);
		}

		hr = pInArchive->Extract(indics.data(), static_cast<UInt32>(indics.size()), FALSE, this);
		CHECK_return(hr);

		UINT width, height;
		m_canvas->GetSize(&width, &height);
		hr = WriteToStream(m_pWicFactory, m_canvas, outputStream);
		CHECK_return(hr);
		return hr;
	}

	void SortItems()
	{
		auto naturalCompareFunc = [](const ItemInfo& a, const ItemInfo& b) -> bool {
			return StrCmpLogicalW(a.name, b.name) < 0;
		};
		auto ordinalCompareFunc = [](const ItemInfo& a, const ItemInfo& b) -> bool {
			return StrCmpIW(a.name, b.name) < 0;
		};
		auto cultureCompareFunc = [](const ItemInfo& a, const ItemInfo& b) -> bool {
			return StrCmpICW(a.name, b.name) < 0;
		};

		// sort
		std::sort(m_items.begin(), m_items.end(),
			(m_config.fileNameOrdering() == FileNameOrdering::Natural) ? naturalCompareFunc :
			(m_config.fileNameOrdering() == FileNameOrdering::Culture) ? cultureCompareFunc :
			ordinalCompareFunc);
	}

	STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override
	{
		static const QITAB qit[] =
		{
			{ &IID_IArchiveExtractCallback, OFFSETOFCLASS(IArchiveExtractCallback, CbzToImage)},
			{ 0 },
		};
		return QISearch(this, qit, riid, ppv);
	}

	STDMETHODIMP_(ULONG) AddRef(void) override
	{
		// On stack only
		return 0;
	}

	STDMETHODIMP_(ULONG) Release(void) override
	{
		// On stack only
		return 0;
	}

	//IArchiveExtractCallback
	STDMETHODIMP SetTotal(UInt64 totalBytes) override
	{
		return S_OK;
	}

	// IArchiveExtractCallback
	STDMETHODIMP SetCompleted(const UInt64* completeValue) override
	{
		if (completeValue == nullptr)
			return E_INVALIDARG;

		if (*completeValue == 0)
		{
			m_lastExtractImageSream = nullptr;
			return S_OK;
		}

		if (m_lastExtractImageSream)
		{
			HRESULT hr = m_lastExtractImageSream->Seek({}, STREAM_SEEK_SET, nullptr);
			CHECK_return(hr);

			CComPtr<IWICBitmapDecoder> pDecoder;
			hr = m_pWicFactory->CreateDecoderFromStream(m_lastExtractImageSream,
				nullptr, WICDecodeMetadataCacheOnDemand, &pDecoder);
			CHECK_return(hr);

			CComPtr<IWICBitmapFrameDecode> pFrame;
			hr = pDecoder->GetFrame(0, &pFrame);
			CHECK_return(hr);

			UINT width{}, height{};
			hr = pFrame->GetSize(&width, &height);
			CHECK_return(hr);

			m_totalHeight += height;
			m_totalWidth += width;
			m_maxHeight = max(m_maxHeight, height);
			m_maxWidth = max(m_maxWidth, width);

			hr = ConcatImage(pFrame);

			m_lastExtractImageSream = nullptr;
			return hr;
		}
		else
			return S_FALSE;
	}

	CComPtr<IStream> m_lastExtractImageSream;

	// IArchiveExtractCallback
	STDMETHODIMP GetStream(UInt32 index, ISequentialOutStream** outStream, Int32 askExtractMode) override
	{
		auto askMode{ static_cast<decltype(NArchive::NExtract::NAskMode::kExtract)>(askExtractMode) };
		if (askMode != NArchive::NExtract::NAskMode::kExtract)
		{
			return S_OK;
		}
		CComPtr<IStream> stream;
		HRESULT hr = CreateStreamOnHGlobal(nullptr, TRUE, &stream);
		*outStream = StreamWrapper::Create(stream);
		if (*outStream == nullptr)
			hr = E_OUTOFMEMORY;

		m_lastExtractImageSream = stream;
		return hr;
	}
	// IArchiveExtractCallback
	STDMETHODIMP PrepareOperation(Int32 askExtractMode) override
	{
		auto askMode{ static_cast<decltype(NArchive::NExtract::NAskMode::kExtract)>(askExtractMode) };
		return S_OK;
	}
	// IArchiveExtractCallback
	STDMETHODIMP SetOperationResult(Int32 opRes) override
	{
		auto operationResult{ static_cast<decltype(NArchive::NExtract::NOperationResult::kOK)>(opRes) };
		return S_OK;
	}


	CComPtr< ID2D1Factory> m_pD2D1Factory;
	HRESULT ConcatImage(IWICBitmapSource* srcImage2)
	{
		HRESULT hr;
		if (not m_canvas)
		{
			CComPtr<IWICBitmapSource> pScaledImage;
			hr = ScaleImage(m_pWicFactory, srcImage2, m_config.scale(), &pScaledImage);
			CHECK_return(hr);
			hr = m_pWicFactory->CreateBitmapFromSource(pScaledImage, WICBitmapCacheOnLoad, &m_canvas);
			return hr;
		}
		UINT width{}, height{};
		hr = m_canvas->GetSize(&width, &height);
		CHECK_return(hr);

		WICPixelFormatGUID format{};
		hr = m_canvas->GetPixelFormat(&format);
		auto formatName = GetPixelFormatName(format);

		UINT width2{}, height2{};
		hr = ScaledSize(srcImage2, m_config.scale(), width2, height2);
		CHECK_return(hr);

		UINT newheight = height + height2;
		if (newheight > 32767)
		{
			return S_FALSE;
		}

		CComPtr<IWICBitmap> target;
		hr = m_pWicFactory->CreateBitmap(
			max(width, width2),
			newheight,
			InternalWICFormat, WICBitmapCacheOnDemand, &target);
		CHECK_return(hr);

		CComPtr<ID2D1RenderTarget> pRT;
		auto rtp = D2D1::RenderTargetProperties(
			D2D1_RENDER_TARGET_TYPE_SOFTWARE
		);
		hr = m_pD2D1Factory->CreateWicBitmapRenderTarget(target, rtp, &pRT);
		CHECK_return(hr);

		CComPtr<ID2D1Bitmap> pD2dBitmap1;
		hr = pRT->CreateBitmapFromWicBitmap(m_canvas, &pD2dBitmap1);
		CHECK_return(hr);

		CComPtr< ID2D1DeviceContext> pDC;
		hr = pRT.QueryInterface(&pDC);
		if (SUCCEEDED(hr) && not m_config.forceD2D1())

		{   // Direct2D1.1
			CComPtr<IWICFormatConverter> converter;
			hr = m_pWicFactory->CreateFormatConverter(&converter);
			CHECK_return(hr);
			hr = converter->Initialize(srcImage2, InternalWICFormat, WICBitmapDitherTypeNone, nullptr, 0.f, WICBitmapPaletteTypeMedianCut);

			CComPtr<ID2D1Bitmap> pD2dBitmap2;
			hr = pDC->CreateBitmapFromWicBitmap(converter, &pD2dBitmap2);
			CHECK_return(hr);

			CComPtr<ID2D1Effect> pD2dScale;
			hr = pDC->CreateEffect(CLSID_D2D1Scale, &pD2dScale);
			CHECK_return(hr);
			pD2dScale->SetInput(0, pD2dBitmap2);
			CHECK_return(hr);
			hr = pD2dScale->SetValue(D2D1_SCALE_PROP_SCALE, D2D1::Vector2F(m_config.scale(), m_config.scale()));
			CHECK_return(hr);
			hr = pD2dScale->SetValue(D2D1_SCALE_PROP_INTERPOLATION_MODE, D2D1_SCALE_INTERPOLATION_MODE_HIGH_QUALITY_CUBIC);
			CHECK_return(hr);

			pDC->BeginDraw();
			pDC->Clear();
			pDC->DrawImage(pD2dBitmap1);

			pDC->DrawImage(pD2dScale,
				D2D1::Point2F(0, height));
			hr = pDC->EndDraw();
		}
		else
		{   // Direct2D1
			CComPtr<IWICBitmapSource> pScaledImage;
			hr = ScaleImage(m_pWicFactory, srcImage2, m_config.scale(), &pScaledImage);
			CHECK_return(hr);

			CComPtr<ID2D1Bitmap> pD2dBitmap2;
			hr = pRT->CreateBitmapFromWicBitmap(pScaledImage, &pD2dBitmap2);
			CHECK_return(hr);


			pRT->BeginDraw();
			pRT->Clear();
			pRT->DrawBitmap(pD2dBitmap1, nullptr, 1.0F, D2D1_BITMAP_INTERPOLATION_MODE_LINEAR);
			pRT->DrawBitmap(pD2dBitmap2, D2D1::RectF(0, height, width2, height + height2),
				1.0F, D2D1_BITMAP_INTERPOLATION_MODE_LINEAR);
			hr = pRT->EndDraw();
		}
		m_canvas = target;

		return hr;
	}
};


STDMETHODIMP CWinMergeScript::get_PluginEvent(/*[out, retval]*/ BSTR* pVal)
{
	auto config = Config::Load();
	*pVal = SysAllocString(config->fileMode() ? L"FILE_PACK_UNPACK" : L"BUFFER_PACK_UNPACK");
	return S_OK;
}

STDMETHODIMP CWinMergeScript::UnpackFile(
	/* [in] */ BSTR fileSrc,
	/* [in] */ BSTR fileDst,
	VARIANT_BOOL* pbChanged,
	INT* pSubcode,
	/* [retval][out] */ VARIANT_BOOL* pbSuccess)
{
	if (fileSrc == nullptr || fileDst == nullptr)
		return E_INVALIDARG;

	*pbSuccess = VARIANT_FALSE;
	auto config = Config::Load();
	CbzToImage sevenzip{ *config };

	HRESULT hr = sevenzip.Init();
	if (FAILED(hr))
		return hr;

	hr = sevenzip.ProcessFile(fileSrc, fileDst);
	if (FAILED(hr))
		return hr;
	*pbChanged = VARIANT_TRUE;
	*pSubcode = 0;
	*pbSuccess = VARIANT_TRUE;
	return S_OK;
}

STDMETHODIMP CWinMergeScript::UnpackBufferA(
	/* [in] */ SAFEARRAY** pBuffer,
	/* [in] */ INT* pSize,
	/* [in] */ VARIANT_BOOL* pbChanged,
	/* [in] */ INT* pSubcode,
	/* [retval][out] */ VARIANT_BOOL* pbSuccess)
{
	if (pBuffer == nullptr || pSize == nullptr)
		return E_INVALIDARG;

	*pbSuccess = VARIANT_FALSE;
	auto config = Config::Load();
	CbzToImage sevenzip{ *config };

	HRESULT hr = sevenzip.Init();
	if (FAILED(hr))
		return hr;

	hr = sevenzip.ProcessBuffer(*pBuffer, *pSize);
	if (FAILED(hr))
		return hr;
	*pbChanged = VARIANT_TRUE;
	*pSubcode = 0;
	*pbSuccess = VARIANT_TRUE;
	return S_OK;
}

IFACEMETHODIMP CWinMergeScript::ShowSettingsDialog(VARIANT_BOOL* pbHandled)
{
	auto result = DialogBoxParam(_Module.GetModuleInstance(),
		MAKEINTRESOURCE(IDD_DIALOG1), nullptr, &SettingDlgProc, 0);

	if (result == -1)
		return HRESULT_FROM_WIN32(GetLastError());

	*pbHandled = (result == IDOK) ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}
