#pragma once
#include "stdafx.h"
#include "Config.h"
#include <string>

HRESULT ConcatPages(IWICImagingFactory *factory, const Config &config, std::vector<winrt::com_ptr<IWICBitmap>> pages,
                    IWICBitmap **target);

HRESULT WriteToStream(IWICImagingFactory *factory, IWICBitmapSource *srcImage, IStream *piStream,
                      std::wstring &extentionName);

class PdfToImage final
{
    Config &m_config;

  public:
    explicit PdfToImage(Config &config) : m_config(config)
    {
    }
    HRESULT ProcessFile(LPCWSTR path, LPCWSTR dest);
    HRESULT Process(winrt::Windows::Storage::Streams::IRandomAccessStream stream, IStream *outputStream,
                    std::wstring &extension);
};
