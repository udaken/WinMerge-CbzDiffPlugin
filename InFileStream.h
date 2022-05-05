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

	IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override
	{
		static const QITAB qit[] =
		{
			{ &IID_IInStream, OFFSETOFCLASS(IInStream, InFileStream)},
			{ &IID_IStreamGetSize, OFFSETOFCLASS(IStreamGetSize, InFileStream)},
			{ 0 },
		};
		return QISearch(this, qit, riid, ppv);
	}

	IFACEMETHODIMP_(ULONG) AddRef() override
	{
		return InterlockedIncrement(&m_Ref);
	}

	IFACEMETHODIMP_(ULONG) Release() override
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

class InSafeArrayStream final :
	public IInStream, public IStreamGetSize
{
	long m_Ref{ 1 };
	ULONG m_offset{ 0 };
	CComSafeArray<BYTE> m_baseStram;
	explicit InSafeArrayStream(LPSAFEARRAY&& stream) : m_baseStram()
	{
		m_baseStram.Attach(stream);
		stream = nullptr;
	}

	~InSafeArrayStream()
	{
	}
public:

	static HRESULT From(LPSAFEARRAY&& stream, InSafeArrayStream** pobj)
	{
		*pobj = new(std::nothrow) InSafeArrayStream(std::move(stream));
		if (*pobj == nullptr)
			return E_OUTOFMEMORY;
		return S_OK;
	}

	InSafeArrayStream(const InSafeArrayStream&) = delete;
	InSafeArrayStream& operator=(const InSafeArrayStream&) = delete;

	IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override
	{
		static const QITAB qit[] =
		{
			{ &IID_IInStream, OFFSETOFCLASS(IInStream, InFileStream)},
			{ &IID_IStreamGetSize, OFFSETOFCLASS(IStreamGetSize, InFileStream)},
			{ 0 },
		};
		return QISearch(this, qit, riid, ppv);
	}

	IFACEMETHODIMP_(ULONG) AddRef() override
	{
		return InterlockedIncrement(&m_Ref);
	}

	IFACEMETHODIMP_(ULONG) Release() override
	{
		ULONG count = InterlockedDecrement(&m_Ref);
		if (count == 0)
		{
			delete this;
			return 0;
		}
		return count;
	}

	IFACEMETHODIMP Read(void* data, UInt32 size, UInt32* processedSize) override
	{
		auto readCount = min(size, m_baseStram.GetCount() - m_offset);
		if (readCount == 0)
			return S_FALSE;

		memcpy(data, &m_baseStram[(LONG)m_offset], readCount);
		m_offset += readCount;
		*processedSize = readCount;
		return S_OK;
	}
	IFACEMETHODIMP Seek(Int64 offset, UInt32 seekOrigin, UInt64* newPosition) override
	{
		Int64 newPos = 0;
		switch (seekOrigin)
		{
		case STREAM_SEEK_SET:
			newPos = offset;
			break;
		case STREAM_SEEK_CUR:
			newPos = m_offset + offset;
			break;
		case STREAM_SEEK_END:
			newPos = m_baseStram.GetCount() + offset;
			break;
		default:
			return E_INVALIDARG;
		}
		m_offset = static_cast<ULONG>(std::clamp(newPos, 0LL, static_cast<Int64>(m_baseStram.GetCount())));;

		if (newPosition)
			*newPosition = m_offset;

		return S_OK;
	}

	IFACEMETHODIMP GetSize(UInt64* size) override
	{
		*size = m_baseStram.GetCount();
		return S_OK;
	}
};

DECLARE_INTERFACE_(ISafeArrayStream, IStream)
{
	virtual HRESULT GetSize(ULONG * pSize) PURE;
	virtual HRESULT Detach(LPSAFEARRAY * ppsa) PURE;
};

class SafeArrayStream final :
	public ISafeArrayStream
{
	long m_Ref{ 1 };
	ULONG m_offset{ 0 };
	CComSafeArray<BYTE> m_safeArray;
	explicit SafeArrayStream(LPSAFEARRAY stream) : m_safeArray()
	{
		if (stream)
			m_safeArray.Attach(stream);
		else
			m_safeArray.Create();
	}

	~SafeArrayStream()
	{
	}
public:

	static HRESULT From(
		_In_opt_ LPSAFEARRAY psa,
		__RPC__deref_out_opt ISafeArrayStream** ppstrm)
	{
		*ppstrm = new(std::nothrow) SafeArrayStream(psa);
		return *ppstrm ? S_OK : E_OUTOFMEMORY;
	}

	SafeArrayStream(const SafeArrayStream&) = delete;
	SafeArrayStream& operator=(const SafeArrayStream&) = delete;

	IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv) override
	{
		static const QITAB qit[] =
		{
			{ &IID_IInStream, OFFSETOFCLASS(IInStream, InFileStream)},
			{ &IID_IStreamGetSize, OFFSETOFCLASS(IStreamGetSize, InFileStream)},
			{ 0 },
		};
		return QISearch(this, qit, riid, ppv);
	}

	IFACEMETHODIMP_(ULONG) AddRef() override
	{
		return InterlockedIncrement(&m_Ref);
	}

	IFACEMETHODIMP_(ULONG) Release() override
	{
		ULONG count = InterlockedDecrement(&m_Ref);
		if (count == 0)
		{
			delete this;
			return 0;
		}
		return count;
	}

	IFACEMETHODIMP Seek(
		/* [in] */ LARGE_INTEGER dlibMove,
		/* [in] */ DWORD dwOrigin,
		/* [annotation] */
		_Out_opt_  ULARGE_INTEGER* plibNewPosition) override
	{
		Int64 newPos = 0;
		switch (dwOrigin)
		{
		case STREAM_SEEK_SET:
			newPos = dlibMove.QuadPart;
			break;
		case STREAM_SEEK_CUR:
			newPos = m_offset + dlibMove.QuadPart;
			break;
		case STREAM_SEEK_END:
			newPos = m_safeArray.GetCount() + dlibMove.QuadPart;
			break;
		default:
			return E_INVALIDARG;

		}
		m_offset = static_cast<ULONG>(std::clamp(newPos, 0LL, static_cast<Int64>(m_safeArray.GetCount())));;
		if (plibNewPosition)
			plibNewPosition->QuadPart = m_offset;

		return S_OK;
	}

	IFACEMETHODIMP SetSize(
		/* [in] */ ULARGE_INTEGER libNewSize) override
	{
		if (libNewSize.HighPart)
			return E_INVALIDARG;

		return m_safeArray.Resize(libNewSize.LowPart);
	}

	IFACEMETHODIMP CopyTo(
		/* [annotation][unique][in] */
		_In_  IStream* pstm,
		/* [in] */ ULARGE_INTEGER cb,
		/* [annotation] */
		_Out_opt_  ULARGE_INTEGER* pcbRead,
		/* [annotation] */
		_Out_opt_  ULARGE_INTEGER* pcbWritten) override
	{
		auto readCount = static_cast<ULONG>(min(cb.QuadPart, m_safeArray.GetCount() - m_offset));

		ULONG written{};
		HRESULT hr = pstm->Write(&m_safeArray[(LONG)m_offset], readCount, &written);
		if (FAILED(hr)) return hr;
		m_offset += readCount;
		if (pcbRead) pcbRead->QuadPart = readCount;
		if (pcbWritten) pcbWritten->QuadPart = written;
		return hr;
	}

	IFACEMETHODIMP Commit(
		/* [in] */ DWORD grfCommitFlags) override
	{
		return S_OK;
	}

	IFACEMETHODIMP Revert(void) override
	{
		return E_NOTIMPL;
	}

	IFACEMETHODIMP LockRegion(
		/* [in] */ ULARGE_INTEGER libOffset,
		/* [in] */ ULARGE_INTEGER cb,
		/* [in] */ DWORD dwLockType) override
	{
		return E_NOTIMPL;
	}

	IFACEMETHODIMP UnlockRegion(
		/* [in] */ ULARGE_INTEGER libOffset,
		/* [in] */ ULARGE_INTEGER cb,
		/* [in] */ DWORD dwLockType)  override
	{
		return E_NOTIMPL;
	}

	IFACEMETHODIMP Stat(
		/* [out] */ __RPC__out STATSTG* pstatstg,
		/* [in] */ DWORD grfStatFlag)  override
	{
		*pstatstg = {};
		pstatstg->type = STGTY_STREAM;
		pstatstg->clsid = IID_NULL;
		pstatstg->cbSize.QuadPart = m_safeArray.GetCount();
		if (grfStatFlag == STATFLAG_NONAME)
		{
			return S_OK;
		}
		else
			return E_INVALIDARG;
	}

	IFACEMETHODIMP Clone(
		/* [out] */ __RPC__deref_out_opt IStream** ppstm) override
	{
		LPSAFEARRAY newArray{};
		HRESULT hr = m_safeArray.CopyTo(&newArray);
		if (FAILED(hr)) return hr;
		*ppstm = new (std::nothrow) SafeArrayStream(newArray);
		return *ppstm ? S_OK : E_OUTOFMEMORY;
	}

	IFACEMETHODIMP Read(
		/* [annotation] */
		_Out_writes_bytes_to_(cb, *pcbRead)  void* pv,
		/* [annotation][in] */
		_In_  ULONG cb,
		/* [annotation] */
		_Out_opt_  ULONG* pcbRead) override
	{
		auto readCount = min(cb, m_safeArray.GetCount() - m_offset);
		memcpy(pv, &m_safeArray[(LONG)m_offset], readCount);
		m_offset += readCount;
		if (pcbRead)
			*pcbRead = readCount;
		return S_OK;
	}

	IFACEMETHODIMP Write(
		/* [annotation] */
		_In_reads_bytes_(cb)  const void* pv,
		/* [annotation][in] */
		_In_  ULONG cb,
		/* [annotation] */
		_Out_opt_  ULONG* pcbWritten)  override
	{
		HRESULT hr = m_safeArray.Add(cb, reinterpret_cast<const BYTE*>(pv));
		m_offset += cb;

		if (pcbWritten)
		{
			*pcbWritten = cb;
		}
		return hr;
	}

	IFACEMETHODIMP GetSize(ULONG* pSize) override
	{
		*pSize = m_safeArray.GetCount();
		return S_OK;
	}

	IFACEMETHODIMP Detach(LPSAFEARRAY* ppsa) override
	{
		*ppsa = m_safeArray.Detach();
		return S_OK;
	}

};
