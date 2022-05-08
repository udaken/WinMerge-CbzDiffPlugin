#include "stdafx.h"
#include "PdfToImage.h"

#include <shcore.h>
#include <unknwn.h>
#pragma comment(lib, "windowsapp")

#include "winrt/base.h"
#include "winrt/Windows.Foundation.Collections.h"
#include "winrt/Windows.Foundation.h"
#include "winrt/Windows.System.h"

#include "winrt/Windows.Data.Pdf.h"
#include "winrt/Windows.Storage.Streams.h"
#include "winrt/Windows.Storage.h"
#include "winrt/Windows.UI.ViewManagement.h"

using namespace winrt;
using namespace winrt::Windows::Foundation;
using namespace winrt::Windows::Storage;

#define WINRT_IID_PPV_ARGS(ppType) winrt::guid_of<decltype(ppType)>(), winrt::put_abi(ppType)

// ------------------------------------------------------------------
//
// ------------------------------------------------------------------
IAsyncAction RenderAsync(winrt::com_ptr<IWICImagingFactory> m_imagingFactory,
                         winrt::Windows::Storage::Streams::IRandomAccessStream stream, FLOAT scale,
                         std::vector<winrt::com_ptr<IWICBitmap>> &pages)
{
    using namespace winrt::Windows::Data::Pdf;
    using namespace winrt::Windows::Storage::Streams;
    winrt::apartment_context ui_thread;
    // co_await winrt::resume_background(); // Return control; resume on thread pool.

    auto doc{co_await PdfDocument::LoadFromStreamAsync(stream)};

    UINT pageCount = doc.PageCount();
    for (UINT pageIndex = 0; pageIndex < pageCount; pageIndex++)
    {
        auto page{doc.GetPage(pageIndex)};

        auto size{page.Size()};
        auto rotation{page.Rotation()};

        InMemoryRandomAccessStream buf{};
        PdfPageRenderOptions renderOptions{};
        renderOptions.BitmapEncoderId(CLSID_WICBmpEncoder);
        renderOptions.DestinationWidth(static_cast<UINT>(size.Width * scale));
        renderOptions.DestinationHeight(static_cast<UINT>(size.Height * scale));
        co_await page.RenderToStreamAsync(buf, renderOptions);

        winrt::com_ptr<IStream> s;
        winrt::check_hresult(::CreateStreamOverRandomAccessStream(winrt::get_unknown(buf), WINRT_IID_PPV_ARGS(s)));

        winrt::com_ptr<IWICBitmapDecoder> decoder;
        winrt::check_hresult(m_imagingFactory->CreateDecoderFromStream(s.get(), &GUID_VendorMicrosoft,
                                                                       WICDecodeMetadataCacheOnLoad, decoder.put()));

        winrt::com_ptr<IWICBitmapFrameDecode> frame;
        winrt::check_hresult(decoder->GetFrame(0, frame.put()));

        winrt::com_ptr<IWICFormatConverter> converter;
        winrt::check_hresult(m_imagingFactory->CreateFormatConverter(converter.put()));
        winrt::check_hresult(converter->Initialize(frame.get(), GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherTypeNone,
                                                   nullptr, 0.0, WICBitmapPaletteTypeCustom));

        winrt::com_ptr<IWICBitmap> a;
        winrt::check_hresult(m_imagingFactory->CreateBitmapFromSource(converter.get(), WICBitmapCacheOnLoad, a.put()));

        pages.emplace_back(a);
    }
    // co_await ui_thread; // Switch back to calling context.
}

HRESULT PdfToImage::Process(winrt::Windows::Storage::Streams::IRandomAccessStream stream, IStream *outputStream,
                            std::wstring &extension)
{
    auto imagingFactory = winrt::create_instance<IWICImagingFactory>(CLSID_WICImagingFactory);

    std::vector<winrt::com_ptr<IWICBitmap>> pages;
    RenderAsync(imagingFactory, stream, m_config.scale(), pages).get();
    winrt::com_ptr<IWICBitmap> result;
    winrt::check_hresult(ConcatPages(imagingFactory.get(), m_config, pages, result.put()));

    return WriteToStream(imagingFactory.get(), result.get(), outputStream, extension);
}

HRESULT PdfToImage::ProcessFile(LPCWSTR path, LPCWSTR dest)
{
    try
    {
        WINRT_ASSERT(path);

        WCHAR srcfullpath[MAX_PATH]{};
        if (not GetFullPathNameW(path, ARRAYSIZE(srcfullpath), srcfullpath, nullptr))
            return HRESULT_FROM_WIN32(GetLastError());

        winrt::Windows::Storage::Streams::IRandomAccessStream stream;
        winrt::check_hresult(
            ::CreateRandomAccessStreamOnFile(srcfullpath, (DWORD)FileAccessMode::Read, WINRT_IID_PPV_ARGS(stream)));

        winrt::com_ptr<IStream> out;
        WCHAR destfullpath[MAX_PATH]{};
        if (not GetFullPathNameW(dest, ARRAYSIZE(destfullpath), destfullpath, nullptr))
            return HRESULT_FROM_WIN32(GetLastError());

        winrt::check_hresult(SHCreateStreamOnFileEx(destfullpath, STGM_WRITE | STGM_CREATE, FILE_ATTRIBUTE_NORMAL, TRUE,
                                                    nullptr, out.put()));

        std::wstring extension;
        winrt::check_hresult(Process(stream, out.get(), extension));
        out = {};
        stream = {};

        WCHAR newName[MAX_PATH]{};
        StringCchCopy(newName, ARRAYSIZE(newName), destfullpath);
        PathRenameExtension(newName, extension.c_str());
        if (not MoveFile(destfullpath, newName))
            return HRESULT_FROM_WIN32(GetLastError());
        return S_OK;
    }
    catch (const hresult_error &hr)
    {
        return hr.code();
    }
    catch (...)
    {
        return E_FAIL;
    }
}
