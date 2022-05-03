#pragma once
#include "stdafx.h"

class InFileStream final :
	public IInStream, public IStreamGetSize
{
	long m_Ref{ 1 };
	CComPtr<IStream> const m_baseStram;
	~InFileStream()
	{
	}
	InFileStream(IStream* stream) : m_baseStram(stream)
	{
	}

public:
	InFileStream(const InFileStream&) = delete;
	InFileStream& operator=(const InFileStream&) = delete;

	static HRESULT OpenPath(LPCWSTR pszFile, InFileStream** pobj) noexcept
	{
		CComPtr<IStream> stream;
		HRESULT hr = SHCreateStreamOnFileEx(pszFile,
			STGM_READ, FILE_ATTRIBUTE_NORMAL, FALSE, nullptr, &stream);
		if (FAILED(hr))
			return hr;
		*pobj = new(std::nothrow) InFileStream(stream);
		if (*pobj == nullptr)
			return E_OUTOFMEMORY;
		return hr;
	}

	STDMETHODIMP QueryInterface(REFIID riid, void** ppv) override
	{
		static const QITAB qit[] =
		{
			{ &IID_IInStream, OFFSETOFCLASS(IInStream, InFileStream)},
			{ &IID_IStreamGetSize, OFFSETOFCLASS(IStreamGetSize, InFileStream)},
			{ 0 },
		};
		return QISearch(this, qit, riid, ppv);
	}

	STDMETHODIMP_(ULONG) AddRef() override
	{
		return InterlockedIncrement(&m_Ref);
	}

	STDMETHODIMP_(ULONG) Release() override
	{
		ULONG count = InterlockedDecrement(&m_Ref);
		if (count == 0)
		{
			delete this;
			return 0;
		}
		return count;
	}

	STDMETHOD(Read)(void* data, UInt32 size, UInt32* processedSize) override
	{
		return m_baseStram->Read(data, size, reinterpret_cast<ULONG*>(processedSize));
	}
	STDMETHOD(Seek)(Int64 offset, UInt32 seekOrigin, UInt64* newPosition) override
	{
		return m_baseStram->Seek(LARGE_INTEGER{ .QuadPart = offset }, seekOrigin, reinterpret_cast<ULARGE_INTEGER*>(newPosition));
	}

	STDMETHOD(GetSize)(UInt64* size) override
	{
		STATSTG stat{};
		HRESULT hr = m_baseStram->Stat(&stat, STATFLAG_NONAME);
		*size = stat.cbSize.QuadPart;
		return hr;
	}

};
