#pragma once

#include <string>
#include <vector>
#include <PathCch.h>
#include <Windows.h>
#pragma comment(lib, "Pathcch.lib")

class IniConfigManager
{
private:
	std::wstring m_configFilePath; // 配置文件
public:
	~IniConfigManager() {

	}
	IniConfigManager(std::wstring fileName, bool bFileNameContainPath = false)
		: m_configFilePath(bFileNameContainPath ? fileName : getConfigAbsPath(fileName))
	{

	}

	void setConfigFilePath(std::wstring filePath) {
		m_configFilePath = filePath;
	}
	std::wstring getConfigFilePath() const {
		return m_configFilePath;
	}

	/**
	 *写入配置文件
	 * @param section 配置区域
	 * @param key 键
	 * @param value 值
	 * @return true成功，false失败
	 */
	bool writeConfigString(std::wstring section, std::wstring key, std::wstring value) {
		return WritePrivateProfileString(section.c_str(), key.c_str(), value.c_str(), m_configFilePath.c_str());
	}

	template <typename T>
	bool writeConfigNumber(std::wstring section, std::wstring key, T value) {
		return WritePrivateProfileString(section.c_str(), key.c_str(), std::to_wstring(value).c_str(), m_configFilePath.c_str());
	}

	/**
	 *写入配置文件
	 * @param section 配置区域
	 * @param key 键
	 * @param default 如果没有则使用默认值返回
	 * @return 对应的配置值
	 */
	std::wstring readConfigString(std::wstring section, std::wstring key, std::wstring default) {
		if (section.empty() || key.empty())
			return default;
		size_t n = MAX_PATH;
		std::vector<wchar_t> buf(n);
		auto&& path = m_configFilePath.c_str();
		// 返回 n - 1个字符说明可能缓冲区不够长，进行增长测试
		for (size_t got = 0; (got = GetPrivateProfileString(section.c_str(), key.c_str(), default.c_str(), buf.data(), n, path)) == n - 1;)
		{
			n *= 2;
			buf.resize(n);
		}
		return std::wstring(buf.data());
	}
	int readConfigInt(std::wstring section, std::wstring key, int default) {
		return GetPrivateProfileInt(section.c_str(), key.c_str(), default, m_configFilePath.c_str());
	}

	// 明面是bool，实际是int，算了不误人子弟了
	//bool readConfigBool(std::wstring section, std::wstring key, bool default) {
	//	return readConfigInt(section, key, default);
	//}

	std::wstring getAppAbsPath()
	{
		std::wstring rst;
		size_t n = MAX_PATH;
		TCHAR* p = new TCHAR[n];
		size_t got = 0;
		for (; (got = GetModuleFileName(NULL, p, n)) == n;)
		{
			delete[] p;
			n *= 2;
			p = new TCHAR[n];
		}
		if (!got)
			return rst;
		PathRemoveFileSpec(p);
		rst = p;
		delete[] p;
		return rst;
	}

	std::wstring pathAppend(std::wstring pathBase, std::wstring pathMore) {
		auto len1 = pathBase.size();
		auto len2 = pathMore.size();
		auto newLen = len1 + len2;
		auto bufSize = newLen + 10;
		std::vector<wchar_t> buffer(bufSize);
		memcpy(buffer.data(), pathBase.data(), len1 * sizeof(wchar_t));
		buffer[len1] = 0;
		HRESULT hr = S_OK;
		for (; FAILED(hr = PathCchAppendEx(buffer.data(), bufSize, pathMore.c_str(), PATHCCH_ALLOW_LONG_PATHS));)
		{
			bufSize *= 2;
			buffer.resize(bufSize);
			memcpy(buffer.data(), pathBase.data(), len1 * sizeof(wchar_t));
			buffer[len1] = 0;
		}
		return std::wstring(buffer.data());
	}

	std::wstring getConfigAbsPath(std::wstring configFileName)
	{
		auto appPath = getAppAbsPath();
		return pathAppend(appPath, configFileName);
	}



};

