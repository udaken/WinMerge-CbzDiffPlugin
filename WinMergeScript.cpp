// WinMergeScript.cpp : Implementation of CWinMergeScript
#include "stdafx.h"
#pragma warning(disable : 28251)
#include "CbzDiffPlugin.h"
#include "Config.h"
#include "Streams.h"
#include "WinMergeScript.h"
#include "resource.h"

#include <algorithm>
#include <assert.h>
#include <oleauto.h>
#include <string>
#include <string_view>
#include <vector>

#include "7zip/CPP/7zip/Archive/IArchive.h"
#include "7zip/CPP/7zip/ICoder.h"
#include "7zip/CPP/7zip/IPassword.h"

#pragma comment(lib, "d3d11.lib")
#pragma comment(lib, "d2d1.lib")
#pragma comment(lib, "dxguid.lib")

#ifdef NDEBUG
#define CHECK_return(hr)                                                                                               \
    do                                                                                                                 \
    {                                                                                                                  \
        HRESULT _hr = hr;                                                                                              \
        if (FAILED(_hr))                                                                                               \
        {                                                                                                              \
            return _hr;                                                                                                \
        }                                                                                                              \
    } while (0)
#else
#define CHECK_return(hr)                                                                                               \
    do                                                                                                                 \
    {                                                                                                                  \
        HRESULT _hr = hr;                                                                                              \
        assert(SUCCEEDED(_hr));                                                                                        \
        if (FAILED(_hr))                                                                                               \
        {                                                                                                              \
            return _hr;                                                                                                \
        }                                                                                                              \
    } while (0)
#endif

using namespace std::string_literals;
using namespace std::string_view_literals;

EXTERN_C const GUID FAR CLSID_CFormat7z;
EXTERN_C const GUID FAR CLSID_CFormatZip;
EXTERN_C const GUID FAR CLSID_CFormatRar;

const REFWICPixelFormatGUID InternalWICFormat = GUID_WICPixelFormat32bppPBGRA;

static std::tuple<REFGUID, std::vector<PROPBAG2>, std::vector<VARIANT>> containerGuidList[] = {
    {GUID_ContainerFormatTiff,
     {{.pstrName = const_cast<LPOLESTR>(L"TiffCompressionMethod")}},
     {
         {.vt = VT_UI1, .bVal = WICTiffCompressionLZW},
     }},
    {GUID_ContainerFormatPng, {}, {}},
    {GUID_ContainerFormatJpeg, {}, {}},
};

class PropVarinat final : public PROPVARIANT
{
  public:
    PropVarinat(const PropVarinat &) = delete;
    PropVarinat &operator=(const PropVarinat &) = delete;

    PropVarinat() : PROPVARIANT{}
    {
    }
    ~PropVarinat()
    {
        PropVariantClear(this);
    }
};

class StreamWrapper final : public ISequentialOutStream
{
    CComPtr<IStream> const baseStream;
    long m_Ref{1};

    explicit StreamWrapper(IStream *s) : baseStream(s)
    {
    }

  public:
    StreamWrapper(const StreamWrapper &) = delete;
    StreamWrapper &operator=(const StreamWrapper &) = delete;

    static StreamWrapper *Create(IStream *s)
    {
        return new (std::nothrow) StreamWrapper(s);
    }

    IFACEMETHODIMP QueryInterface(REFIID riid, void **ppvObject) override
    {
        static const QITAB qit[] = {
            {&IID_ISequentialOutStream, OFFSETOFCLASS(ISequentialOutStream, StreamWrapper)},
            {0},
        };
        return QISearch(this, qit, riid, ppvObject);
    }
    IFACEMETHODIMP_(ULONG) AddRef(void) override
    {
        return InterlockedIncrement(&m_Ref);
    }
    IFACEMETHODIMP_(ULONG) Release(void) override
    {
        ULONG count = InterlockedDecrement(&m_Ref);
        if (count == 0)
        {
            delete this;
            return 0;
        }
        return count;
    }
    STDMETHOD(Write)(const void *data, UInt32 size, UInt32 *processedSize) override
    {
        return baseStream->Write(data, static_cast<ULONG>(size), reinterpret_cast<ULONG *>(processedSize));
    }
};

inline HRESULT GetExteinsionFromEncoder(IWICBitmapEncoder *pngEncoder, std::wstring &ext)
{
    HRESULT hr = E_FAIL;
    CComPtr<IWICBitmapEncoderInfo> pEncodedrInfo;
    hr = pngEncoder->GetEncoderInfo(&pEncodedrInfo);
    UINT length{};
    hr = pEncodedrInfo->GetFileExtensions(0, nullptr, &length);

    ext.resize(static_cast<size_t>(length - 1), L'\0');

    hr = pEncodedrInfo->GetFileExtensions(length, ext.data(), &length);

    std::string::size_type pos = 0;
    if ((pos = ext.find(',', pos)) != std::string::npos)
    {
        ext.resize(pos);
    }

    return hr;
}

HRESULT WriteToStream(IWICImagingFactory *factory, IWICBitmapSource *srcImage, IStream *piStream,
                      std::wstring &extentionName)
{
    HRESULT hr = E_FAIL;

    for (auto &[container, prop, variant] : containerGuidList)
    {
        CComPtr<IWICBitmapEncoder> pngEncoder;
        hr = factory->CreateEncoder(container, nullptr, &pngEncoder);
        if (FAILED(hr))
            continue;
        hr = pngEncoder->Initialize(piStream, WICBitmapEncoderNoCache);
        if (FAILED(hr))
            continue;
        CComPtr<IWICBitmapFrameEncode> piBitmapFrame;
        CComPtr<IPropertyBag2> pPropertybag;
        hr = pngEncoder->CreateNewFrame(&piBitmapFrame, &pPropertybag);
        if (FAILED(hr))
            continue;
        if (prop.size())
            pPropertybag->Write(prop.size(), prop.data(), variant.data());
        hr = piBitmapFrame->Initialize(pPropertybag);
        GUID format = GUID_WICPixelFormat32bppPBGRA;
        hr = piBitmapFrame->SetPixelFormat(&format);
        if (FAILED(hr))
            continue;
        hr = piBitmapFrame->WriteSource(srcImage, nullptr);
        if (FAILED(hr))
            continue;
        hr = piBitmapFrame->Commit();
        if (FAILED(hr))
            continue;
        hr = pngEncoder->Commit();
        if (FAILED(hr))
            continue;
        hr = piStream->Commit(STGC_DEFAULT);
        if (FAILED(hr))
            continue;

        if (SUCCEEDED(hr))
        {
            hr = GetExteinsionFromEncoder(pngEncoder, extentionName);
            return hr;
        }
    }

    return hr;
}

inline LPCWSTR GetPixelFormatName(WICPixelFormatGUID guid)
{
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormatDontCare))
        return L"GUID_WICPixelFormatDontCare";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat1bppIndexed))
        return L"GUID_WICPixelFormat1bppIndexed";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat2bppIndexed))
        return L"GUID_WICPixelFormat2bppIndexed";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat4bppIndexed))
        return L"GUID_WICPixelFormat4bppIndexed";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat8bppIndexed))
        return L"GUID_WICPixelFormat8bppIndexed";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormatBlackWhite))
        return L"GUID_WICPixelFormatBlackWhite";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat2bppGray))
        return L"GUID_WICPixelFormat2bppGray";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat4bppGray))
        return L"GUID_WICPixelFormat4bppGray";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat8bppGray))
        return L"GUID_WICPixelFormat8bppGray";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat8bppAlpha))
        return L"GUID_WICPixelFormat8bppAlpha";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat16bppBGR555))
        return L"GUID_WICPixelFormat16bppBGR555";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat16bppBGR565))
        return L"GUID_WICPixelFormat16bppBGR565";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat16bppBGRA5551))
        return L"GUID_WICPixelFormat16bppBGRA5551";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat16bppGray))
        return L"GUID_WICPixelFormat16bppGray";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat24bppBGR))
        return L"GUID_WICPixelFormat24bppBGR";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat24bppRGB))
        return L"GUID_WICPixelFormat24bppRGB";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppBGR))
        return L"GUID_WICPixelFormat32bppBGR";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppBGRA))
        return L"GUID_WICPixelFormat32bppBGRA";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppPBGRA))
        return L"GUID_WICPixelFormat32bppPBGRA";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppGrayFloat))
        return L"GUID_WICPixelFormat32bppGrayFloat";

    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppRGBA))
        return L"GUID_WICPixelFormat32bppRGBA";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppPRGBA))
        return L"GUID_WICPixelFormat32bppPRGBA";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat48bppRGB))
        return L"GUID_WICPixelFormat48bppRGB";
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat48bppBGR))
        return L"GUID_WICPixelFormat48bppBGR";

    return L"";
}

inline D2D1_PIXEL_FORMAT WicPixelFormatToD2D1(WICPixelFormatGUID guid)
{
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppPRGBA))
        return D2D1::PixelFormat(DXGI_FORMAT_R8G8B8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED);
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppBGR))
        return D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_IGNORE);
    if (InlineIsEqualGUID(guid, GUID_WICPixelFormat32bppPBGRA))
        return D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_PREMULTIPLIED);
    return D2D1::PixelFormat();
}

inline HRESULT ScaledSize(IWICBitmapSource *srcImage, float scale, UINT &width, UINT &height)
{
    HRESULT hr;
    hr = srcImage->GetSize(&width, &height);
    CHECK_return(hr);

    width = static_cast<UINT>(width * scale);
    height = static_cast<UINT>(height * scale);

    return hr;
}

inline HRESULT ScaleImage(IWICImagingFactory *factory, IWICBitmapSource *srcImage, float scale, IWICBitmap **result)
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
    hr = converter->Initialize(scaler, InternalWICFormat, WICBitmapDitherTypeNone, nullptr, 0.f,
                               WICBitmapPaletteTypeMedianCut);
    CHECK_return(hr);

    hr = factory->CreateBitmapFromSource(converter, WICBitmapCacheOnLoad, result);
    CHECK_return(hr);
    return hr;
}

HRESULT ConcatPages(IWICImagingFactory *factory, const Config &config, std::vector<CComPtr<IWICBitmap>> pages,
                    IWICBitmap **target)
{
    HRESULT hr;
    UINT newheight{}, newwidth{};
    for (auto &img : pages)
    {
        UINT width{}, height{};
        hr = img->GetSize(&width, &height);
        CHECK_return(hr);

        if ((newheight + height) > config.imageLimit())
        {
            break;
        }

        newheight += height;
        newwidth = std::max(newwidth, width);

        WICPixelFormatGUID format{};
        hr = img->GetPixelFormat(&format);
        CHECK_return(hr);
        auto formatName = GetPixelFormatName(format);
    }

    hr = factory->CreateBitmap(newwidth, newheight, InternalWICFormat, WICBitmapCacheOnDemand, target);
    CHECK_return(hr);

    CComPtr<ID2D1Factory> pD2D1Factory;
    hr = D2D1CreateFactory(D2D1_FACTORY_TYPE_MULTI_THREADED, &pD2D1Factory);
    CHECK_return(hr);

    CComPtr<ID2D1RenderTarget> pRT;
    auto rtp = D2D1::RenderTargetProperties(D2D1_RENDER_TARGET_TYPE_SOFTWARE);
    hr = pD2D1Factory->CreateWicBitmapRenderTarget(*target, rtp, &pRT);
    CHECK_return(hr);

    CComPtr<ID2D1DeviceContext> pDC;
    hr = pRT.QueryInterface(&pDC);

    pRT->BeginDraw();
    pRT->Clear(D2D1::ColorF(config.backgroundColor()));
    UINT offset{};
    for (auto &img : pages)
    {
        UINT width{}, height{};
        hr = img->GetSize(&width, &height);
        CHECK_return(hr);

        if ((offset + height) > config.imageLimit())
        {
            break;
        }

        CComPtr<ID2D1Bitmap> pD2dBitmap1;
        hr = pRT->CreateBitmapFromWicBitmap(img, &pD2dBitmap1);
        CHECK_return(hr);

        if (pDC && not config.forceD2D1())
        { // Direct2D1.1

            pDC->DrawImage(pD2dBitmap1, D2D1::Point2F(0, offset));
        }
        else
        { // Direct2D1
            pRT->DrawBitmap(pD2dBitmap1, D2D1::RectF(0, offset, width, offset + height), 1.0F,
                            D2D1_BITMAP_INTERPOLATION_MODE_LINEAR);
        }
        offset += height;
    }

    hr = pRT->EndDraw();
    CHECK_return(hr);
    return hr;
}

inline HRESULT GetSupportedExtentions(IWICImagingFactory *factory, std::wstring &result)
{
    CComPtr<IEnumUnknown> pEnum;
    HRESULT hr;
    hr = factory->CreateComponentEnumerator(WICDecoder, WICComponentEnumerateDefault, &pEnum);
    CHECK_return(hr);

    IUnknown *pComponent{};
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

    constexpr auto target{L"."sv};
    constexpr auto replacement{L"*."sv};

    std::string::size_type pos = 0;
    while ((pos = result.find(target, pos)) != std::string::npos)
    {
        result.replace(pos, target.length(), replacement);
        pos += replacement.length();
    }

    return hr;
}

inline HRESULT Load7zDll(HMODULE &h7z)
{
    WCHAR moduleFileName[MAX_PATH]{};
    if (not GetModuleFileName(_Module.GetModuleInstance(), moduleFileName, ARRAYSIZE(moduleFileName)))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }
    PathAppend(moduleFileName, LR"(\..\7z.dll)");
    h7z = LoadLibrary(moduleFileName);
    if (h7z != nullptr)
        return S_OK;

    if (GetModuleFileName(nullptr, moduleFileName, ARRAYSIZE(moduleFileName)))
    {
        // relative path from WinMergeU.exe
        PathAppend(moduleFileName, LR"(\..\Merge7z\7z.dll)");
        h7z = LoadLibrary(moduleFileName);
        if (h7z != nullptr)
            return S_OK;
    }

    return HRESULT_FROM_WIN32(GetLastError());
}

inline HRESULT SafeArrayGetCount(_In_ SAFEARRAY *psa, _Out_ LONG *plLbound, _In_ UINT nDim = 0)
{
    HRESULT hr;
    LONG lUbound;
    hr = SafeArrayGetUBound(psa, nDim, &lUbound);
    CHECK_return(hr);
    LONG lLbound;
    hr = SafeArrayGetLBound(psa, nDim, &lLbound);
    CHECK_return(hr);

    *plLbound = (lUbound - lLbound);

    return hr;
}

class CbzToImage final : IArchiveExtractCallback, ICryptoGetTextPassword2
{
    typedef UINT32(WINAPI *CreateObjectFunc)(REFGUID clsID, REFIID interfaceID, void **outObject);

    struct ItemInfo
    {
        UInt32 index;
        CComBSTR name;
    };

    CComPtr<IWICImagingFactory> m_pWicFactory;

    std::vector<CComPtr<IWICBitmap>> m_pages;

    Config const &m_config;

  public:
    CbzToImage(const CbzToImage &) = delete;
    CbzToImage &operator=(const CbzToImage &) = delete;

    explicit CbzToImage(Config const &config) : m_config(config)
    {
    }
    HRESULT Init()
    {
        HRESULT hr;

        hr = m_pWicFactory.CoCreateInstance(CLSID_WICImagingFactory);
        CHECK_return(hr);

        return hr;
    }

    static REFGUID GetFormatGUID(LPCWSTR path)
    {
        if (PathMatchSpecEx(path, L"*.cbz;*.zip", PMSF_MULTIPLE) == S_OK)
            return CLSID_CFormatZip;
        if (PathMatchSpecEx(path, L"*.cbr;*.rar", PMSF_MULTIPLE) == S_OK)
            return CLSID_CFormatRar;
        if (PathMatchSpecEx(path, L"*.cb7;*.7z", PMSF_MULTIPLE) == S_OK)
            return CLSID_CFormat7z;
        return CLSID_NULL;
    }
    static REFGUID GetFormatGUIDFromSig(const void *sig, size_t size)
    {
        constexpr BYTE magicZip[] = {'P', 'K', '\03', '\04'};
        if (memcmp(sig, "PK\03\04", (std::min(sizeof(magicZip), size))) == 0)
            return CLSID_CFormatZip;
        constexpr BYTE magicRar[] = {'R', 'a', 'r', '!'};
        if (memcmp(sig, magicRar, (std::min(sizeof(magicRar), size))) == 0)
            return CLSID_CFormatRar;
        constexpr BYTE magic7z[] = {'7', 'z', 0xBC, 0xAF, 0x27, 0x1C};
        if (memcmp(sig, magic7z, std::min(sizeof(magic7z), size)) == 0)
            return CLSID_CFormat7z;
        return CLSID_NULL;
    }

    HRESULT ProcessFile(LPCWSTR path, LPCWSTR dest)
    {
        HRESULT hr;

        CComPtr<InFileStream> inputStream;
        hr = InFileStream::OpenPath(path, &inputStream);
        CHECK_return(hr);

        CComPtr<IStream> outStream;
        hr = SHCreateStreamOnFileEx(dest, STGM_WRITE, FILE_ATTRIBUTE_NORMAL, TRUE, nullptr, &outStream);

        std::wstring extension;
        hr = Process(inputStream, GetFormatGUID(path), outStream, extension);
        outStream.Release();

        WCHAR newName[MAX_PATH]{};
        StringCchCopy(newName, ARRAYSIZE(newName), dest);
        PathRenameExtension(newName, extension.c_str());
        if (not MoveFile(dest, newName))
            return HRESULT_FROM_WIN32(GetLastError());

        return hr;
    }

    HRESULT ProcessBuffer(
        /* [in] */ LPSAFEARRAY &pBuffer,
        /* [in] */ INT &pSize)
    {
        HRESULT hr;

        void *sig;
        hr = SafeArrayAccessData(pBuffer, &sig);
        CHECK_return(hr);
        LONG count;
        hr = SafeArrayGetCount(pBuffer, &count);
        CHECK_return(hr);
        auto guid = GetFormatGUIDFromSig(sig, count);
        hr = SafeArrayUnaccessData(pBuffer);

        CComPtr<InSafeArrayStream> inputStream;
        hr = InSafeArrayStream::From(std::move(pBuffer), &inputStream);
        CHECK_return(hr);

        CComPtr<ISafeArrayStream> piStream;
        hr = SafeArrayStream::From(nullptr, &piStream);
        CHECK_return(hr);

        std::wstring extension;
        hr = Process(inputStream, guid, piStream, extension);
        CHECK_return(hr);

        ULONG size;
        piStream->GetSize(&size);
        pSize = size;

        hr = piStream->Detach(&pBuffer);

        return hr;
    }

  private:
    HRESULT Process(IInStream *inputStream, REFGUID guid, IStream *outputStream, std::wstring &extension)
    {
        HRESULT hr;
        if (guid == IID_NULL)
            return ERROR_INVALID_DATA;

        std::wstring supportedExt;

        hr = GetSupportedExtentions(m_pWicFactory, supportedExt);
        CHECK_return(hr);

        HMODULE h7z{};
        hr = Load7zDll(h7z);
        CHECK_return(hr);
        auto deleter{[](HMODULE p) { FreeLibrary(p); }};
        auto _ = std::unique_ptr<decltype(std::remove_pointer_t<HMODULE>()), decltype(deleter)>(h7z);

        auto CreateObject = (CreateObjectFunc)GetProcAddress(h7z, "CreateObject");
        if (CreateObject == nullptr)
            return HRESULT_FROM_WIN32(GetLastError());

        CComPtr<IInArchive> pInArchive;
        hr = CreateObject(guid, IID_IInArchive, IID_PPV_ARGS_Helper(&pInArchive));
        CHECK_return(hr);

        hr = pInArchive->Open(inputStream, nullptr, nullptr);
        CHECK_return(hr);

        PropVarinat vSolid{};
        hr = pInArchive->GetArchiveProperty(kpidSolid, &vSolid);
        if (vSolid.boolVal)
            return ERROR_INVALID_DATA;

        UInt32 itemCount{};
        hr = pInArchive->GetNumberOfItems(&itemCount);
        CHECK_return(hr);
        std::vector<ItemInfo> items;
        items.reserve(itemCount);

        for (UInt32 index = 0; index < itemCount; ++index)
        {
            PropVarinat vPath{};
            hr = pInArchive->GetProperty(index, kpidPath, &vPath);
            CHECK_return(hr);
            CComBSTR name;
            hr = PropVariantToBSTR(vPath, &name);
            CHECK_return(hr);

            if (S_OK == PathMatchSpecExW(name, supportedExt.c_str(), PMSF_MULTIPLE))
            {
                items.emplace_back(index, name);
            }
        }
        SortItems(items);

        std::vector<UInt32> indics;
        indics.reserve(static_cast<size_t>(itemCount));
        for (auto &i : items)
        {
            indics.push_back(i.index);
        }

        hr = pInArchive->Extract(indics.data(), static_cast<UInt32>(indics.size()), FALSE, this);
        CHECK_return(hr);

        pInArchive.Release();
        CComPtr<IWICBitmap> target;
        hr = ConcatPages(m_pWicFactory, m_config, m_pages, &target);
        m_pages.clear();

        hr = WriteToStream(m_pWicFactory, target, outputStream, extension);
        CHECK_return(hr);
        return hr;
    }

    void SortItems(std::vector<ItemInfo> &items)
    {
        auto naturalCompareFunc = [](const ItemInfo &a, const ItemInfo &b) -> bool {
            return StrCmpLogicalW(a.name, b.name) < 0;
        };
        auto ordinalCompareFunc = [](const ItemInfo &a, const ItemInfo &b) -> bool {
            return StrCmpIW(a.name, b.name) < 0;
        };
        auto cultureCompareFunc = [](const ItemInfo &a, const ItemInfo &b) -> bool {
            return StrCmpICW(a.name, b.name) < 0;
        };
#ifndef __INTELLISENSE__ // workaround
        // sort
        std::sort(items.begin(), items.end(),
                  (m_config.fileNameOrdering() == FileNameOrdering::Natural)   ? naturalCompareFunc
                  : (m_config.fileNameOrdering() == FileNameOrdering::Culture) ? cultureCompareFunc
                                                                               : ordinalCompareFunc);
#endif
    }

    IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv) override
    {
        static const QITAB qit[] = {
            {&IID_IArchiveExtractCallback, OFFSETOFCLASS(IArchiveExtractCallback, CbzToImage)},
            {0},
        };
        return QISearch(this, qit, riid, ppv);
    }

    IFACEMETHODIMP_(ULONG) AddRef(void) override
    {
        // On stack only
        return 0;
    }

    IFACEMETHODIMP_(ULONG) Release(void) override
    {
        // On stack only
        return 0;
    }

    // IArchiveExtractCallback
    IFACEMETHODIMP SetTotal(UInt64 totalBytes) override
    {
        return S_OK;
    }

    // IArchiveExtractCallback
    IFACEMETHODIMP SetCompleted(const UInt64 *completeValue) override
    {
        if (completeValue == nullptr)
            return E_INVALIDARG;

        return S_OK;
    }

    CComPtr<IStream> m_lastExtractImageSream;

    // IArchiveExtractCallback
    IFACEMETHODIMP GetStream(UInt32 index, ISequentialOutStream **outStream, Int32 askExtractMode) override
    {
        auto askMode{static_cast<decltype(NArchive::NExtract::NAskMode::kExtract)>(askExtractMode)};
        if (askMode != NArchive::NExtract::NAskMode::kExtract)
        {
            return S_OK;
        }
        CComPtr<IStream> stream;
        HRESULT hr = CreateStreamOnHGlobal(nullptr, TRUE, &stream);
        CHECK_return(hr);
        *outStream = StreamWrapper::Create(stream);
        if (*outStream == nullptr)
            hr = E_OUTOFMEMORY;

        m_lastExtractImageSream = stream;
        return hr;
    }
    // IArchiveExtractCallback
    IFACEMETHODIMP PrepareOperation(Int32 askExtractMode) override
    {
        auto askMode{static_cast<decltype(NArchive::NExtract::NAskMode::kExtract)>(askExtractMode)};
        return S_OK;
    }
    // IArchiveExtractCallback
    IFACEMETHODIMP SetOperationResult(Int32 opRes) override
    {
        if (opRes == NArchive::NExtract::NOperationResult::kOK && m_lastExtractImageSream)
        {
            HRESULT hr;
            hr = m_lastExtractImageSream->Seek({}, STREAM_SEEK_SET, nullptr);
            CHECK_return(hr);

            CComPtr<IWICBitmapDecoder> pDecoder;
            hr = m_pWicFactory->CreateDecoderFromStream(m_lastExtractImageSream, nullptr,
                                                        WICDecodeMetadataCacheOnDemand, &pDecoder);
            CHECK_return(hr);

            CComPtr<IWICBitmapFrameDecode> pFrame;
            hr = pDecoder->GetFrame(0, &pFrame);
            CHECK_return(hr);

            CComPtr<IWICBitmap> page;
            hr = ScaleImage(m_pWicFactory, pFrame, m_config.scale(), &page);
            m_pages.emplace_back(page);

            m_lastExtractImageSream = nullptr;
            return hr;
        }
        return S_FALSE;
    }

    IFACEMETHODIMP CryptoGetTextPassword2(Int32 *passwordIsDefined, BSTR *password) override
    {
        return E_NOTIMPL;
    }
};

IFACEMETHODIMP CWinMergeScript::get_PluginUnpackedFileExtension(/* [retval][out] */ BSTR *pVal)
{
    HRESULT hr;
    CComPtr<IWICImagingFactory> factory;
    hr = factory.CoCreateInstance(CLSID_WICImagingFactory);
    CHECK_return(hr);

    auto &container{*std::begin(containerGuidList)};
    auto &[guid, _1, _2] = container;

    CComPtr<IWICBitmapEncoder> pngEncoder;
    hr = factory->CreateEncoder(guid, nullptr, &pngEncoder);
    CHECK_return(hr);

    std::wstring ext;
    hr = GetExteinsionFromEncoder(pngEncoder, ext);
    CHECK_return(hr);

    *pVal = SysAllocString(ext.c_str());
    return S_OK;
}

IFACEMETHODIMP CWinMergeScript::get_PluginEvent(/*[out, retval]*/ BSTR *pVal)
{
    auto config = Config::Load();
    *pVal = SysAllocString(config->pluginMode() == PluginMode::File     ? L"FILE_PACK_UNPACK"
                           : config->pluginMode() == PluginMode::Folder ? L"FILE_FOLDER_PACK_UNPACK"
                           : config->pluginMode() == PluginMode::Buffer ? L"BUFFER_PACK_UNPACK"
                                                                        : L"");
    return S_OK;
}

IFACEMETHODIMP CWinMergeScript::UnpackFile(
    /* [in] */ BSTR fileSrc,
    /* [in] */ BSTR fileDst, VARIANT_BOOL *pbChanged, INT *pSubcode,
    /* [retval][out] */ VARIANT_BOOL *pbSuccess)
{
    if (fileSrc == nullptr || fileDst == nullptr)
        return E_INVALIDARG;

    *pbSuccess = VARIANT_FALSE;
    auto config = Config::Load();

    HRESULT hr;
    CbzToImage sevenzip{*config};
    hr = sevenzip.Init();
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

IFACEMETHODIMP CWinMergeScript::UnpackBufferA(
    /* [in] */ SAFEARRAY **pBuffer,
    /* [in] */ INT *pSize,
    /* [in] */ VARIANT_BOOL *pbChanged,
    /* [in] */ INT *pSubcode,
    /* [retval][out] */ VARIANT_BOOL *pbSuccess)
{
    if (pBuffer == nullptr || pSize == nullptr)
        return E_INVALIDARG;

    *pbSuccess = VARIANT_FALSE;
    auto config = Config::Load();

    CbzToImage sevenzip{*config};
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

IFACEMETHODIMP CWinMergeScript::UnpackFolder(
    /* [in] */ BSTR fileSrc,
    /* [in] */ BSTR folderDst, VARIANT_BOOL *pbChanged, INT *pSubcode,
    /* [retval][out] */ VARIANT_BOOL *pbSuccess)
{
    WCHAR fileDest[MAX_PATH]{};
    PathCombine(fileDest, folderDst, L"image.png");
    return UnpackFile(fileSrc, fileDest, pbChanged, pSubcode, pbChanged);
}

IFACEMETHODIMP CWinMergeScript::ShowSettingsDialog(VARIANT_BOOL *pbHandled)
{
    auto result =
        DialogBoxParam(_Module.GetModuleInstance(), MAKEINTRESOURCE(IDD_DIALOG1), nullptr, &SettingDlgProc, 0);

    if (result == -1)
        return HRESULT_FROM_WIN32(GetLastError());

    *pbHandled = (result == IDOK) ? VARIANT_TRUE : VARIANT_FALSE;
    return S_OK;
}
