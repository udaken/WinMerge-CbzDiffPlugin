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

enum class PluginMode : DWORD
{
	Buffer,
	File,
	Folder,
};

class Config final
{
	DWORD m_Mode{ static_cast<DWORD>(PluginMode::Folder) };
	DWORD m_forceD2D1{ 0 };
	DWORD m_scale{ 10 };
	DWORD m_fileNameOrdering{ static_cast<DWORD>(FileNameOrdering::Natural) };
	DWORD m_backgroundColor{ 0x0FF };
	DWORD m_imageLimit{ 0xFFFFFFFF };

public:

	static constexpr LPCWSTR key_FileMode{ L"PluginMode" };
	static constexpr LPCWSTR key_Scale{ L"Scale" };
	static constexpr LPCWSTR key_ForceD2D1{ L"ForceD2D1" };
	static constexpr LPCWSTR key_FileNameOrdering{ L"FileNameOrdering" };
	static constexpr LPCWSTR key_BackgroundColor{ L"BackgroundColor" };

	static std::unique_ptr<Config> Load();

	Config(const Config&) = delete;
	Config& operator=(const Config&) = delete;
	Config(Config&&) = default;
	Config() noexcept {}
	Config& operator=(Config&&) = default;

	PluginMode pluginMode() const { return static_cast<PluginMode>(m_Mode); }
	bool fileMode() const { return m_Mode == static_cast<DWORD>(PluginMode::File); }
	bool forceD2D1() const { return m_forceD2D1 == 1; }
	FLOAT scale() const { return  scaleInt() / 100.f; }
	DWORD scaleInt() const { return m_scale; }
	FileNameOrdering fileNameOrdering() const {
		return static_cast<FileNameOrdering>(m_fileNameOrdering);
	}
	UINT backgroundColor() const { return m_backgroundColor & 0xFFFFFF; }
	UINT imageLimit() const { return m_imageLimit; }
};
