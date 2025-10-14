#pragma once
#include <fstream>
#include <string>
#include <sstream>

#include "StringProcess.h"
#include <chrono>
#include <iomanip>

// 自带防日志文件打开失败的日志记录
class FileLogger
{
private:
	std::ofstream fs;
	class FileLoggerStream
	{
	private:
		std::wstringstream buffer;
		FileLogger& logger;
		template <typename T>
		FileLoggerStream(FileLogger& logger, T&& text) : logger(logger) {
			buffer << std::forward<T>(text);
		}
		friend class FileLogger;
	public:
		FileLoggerStream(FileLoggerStream&) = delete; // 禁止拷贝
		FileLoggerStream& operator=(const FileLoggerStream&) = delete;
		FileLoggerStream(FileLoggerStream&& other) noexcept
			: logger(other.logger), buffer(std::move(other.buffer))
		{
		}
		FileLoggerStream& operator=(FileLoggerStream&& other) noexcept {
			if (this != &other) {
				// logger 是引用，不能重新绑定
				buffer = std::move(other.buffer);
			}
			return *this;
		}
		~FileLoggerStream() {
			// 析构时候添加信息并写入日志
			auto& fs = logger.fs;
			if (fs.is_open())
			{
				auto now = std::chrono::system_clock::now();
				std::time_t time = std::chrono::system_clock::to_time_t(now);
				fs << "[" << std::put_time(std::localtime(&time), "%Y-%m-%d %H:%M:%S") << "] "// 加上时间信息
					<< wstr2str_2UTF8(buffer.str()) << std::endl; // << 系列语句结束自动添加换行符
			}
		}

		template <typename T>
		FileLoggerStream& operator<<(T&& text) {
			buffer << std::forward<T>(text); // 完美转发
			return *this;
		}

		FileLoggerStream& operator<<(std::string text) {
			buffer << str2wstr_2UTF8(text);
			return *this;
		}
		// 支持 std::endl 等流操作符
		FileLoggerStream& operator<<(std::ostream& (*s)(std::ostream&)) {
			buffer << s;
			return *this;
		}
	};
public:
	~FileLogger() {
		close();
	}
	FileLogger(std::wstring filePath, bool clearBefore = true) {
		open(filePath, clearBefore);
	}
	FileLogger(std::string filePath, bool clearBefore = true) {
		open(filePath, clearBefore);
	}
	FileLogger() {
		// 默认构造
	}
	bool open(std::wstring filePath, bool clearBefore = true) {
		close();
		fs.open(filePath, std::ofstream::out | (clearBefore ? std::ofstream::trunc : std::ofstream::app));
		return isOpen();
	}
	bool open(std::string filePath, bool clearBefore = true) {
		close();
		fs.open(filePath, std::ofstream::out | (clearBefore ? std::ofstream::trunc : std::ofstream::app));
		return isOpen();
	}
	bool isOpen() const {
		return fs.is_open();
	}
	void close() {
		if (isOpen())
			fs.close();
	}
	template <typename T>
	FileLoggerStream operator<<(T&& text) {
		return FileLoggerStream(*this, std::forward<T>(text));
	}

};

