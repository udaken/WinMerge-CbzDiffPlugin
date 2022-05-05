#pragma once
#include "stdafx.h"

#include <memory>

enum class FileNameOrdering : DWORD
{
	Natural,
	Ordinal,
	Culture,
};

enum class Direction : DWORD
{
	TopToDown,
	RightToLeft,
	LeftToRight,
};

class Config final
{
	DWORD m_fileMode{ 0 };
	DWORD m_forceD2D1{ 0 };
	DWORD m_scale{ 10 };
	DWORD m_fileNameOrdering{ static_cast<DWORD>(FileNameOrdering::Natural) };


public:

	static constexpr LPCWSTR key_FileMode{ L"FileMode" };
	static constexpr LPCWSTR key_Scale{ L"Scale" };
	static constexpr LPCWSTR key_ForceD2D1{ L"ForceD2D1" };
	static constexpr LPCWSTR key_FileNameOrdering{ L"FileNameOrdering" };

	static std::unique_ptr<Config> Load();

	Config(const Config&) = delete;
	Config& operator=(const Config&) = delete;
	Config(Config&&) = default;
	Config() noexcept {}
	Config& operator=(Config&&) = default;

	bool fileMode() const { return m_fileMode == 1; }
	bool forceD2D1() const { return m_forceD2D1 == 1; }
	FLOAT scale() const { return  scaleInt() / 100.f; }
	DWORD scaleInt() const { return m_scale; }
	FileNameOrdering fileNameOrdering() const {
		return static_cast<FileNameOrdering>(m_fileNameOrdering);
	}
};
