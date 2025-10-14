#pragma once
#include <fstream>
#include <string>
#include <sstream>

#include "StringProcess.h"
#include <chrono>
#include <iomanip>

// �Դ�����־�ļ���ʧ�ܵ���־��¼
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
		FileLoggerStream(FileLoggerStream&) = delete; // ��ֹ����
		FileLoggerStream& operator=(const FileLoggerStream&) = delete;
		FileLoggerStream(FileLoggerStream&& other) noexcept
			: logger(other.logger), buffer(std::move(other.buffer))
		{
		}
		FileLoggerStream& operator=(FileLoggerStream&& other) noexcept {
			if (this != &other) {
				// logger �����ã��������°�
				buffer = std::move(other.buffer);
			}
			return *this;
		}
		~FileLoggerStream() {
			// ����ʱ�������Ϣ��д����־
			auto& fs = logger.fs;
			if (fs.is_open())
			{
				auto now = std::chrono::system_clock::now();
				std::time_t time = std::chrono::system_clock::to_time_t(now);
				fs << "[" << std::put_time(std::localtime(&time), "%Y-%m-%d %H:%M:%S") << "] "// ����ʱ����Ϣ
					<< wstr2str_2UTF8(buffer.str()) << std::endl; // << ϵ���������Զ���ӻ��з�
			}
		}

		template <typename T>
		FileLoggerStream& operator<<(T&& text) {
			buffer << std::forward<T>(text); // ����ת��
			return *this;
		}

		FileLoggerStream& operator<<(std::string text) {
			buffer << str2wstr_2UTF8(text);
			return *this;
		}
		// ֧�� std::endl ����������
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
		// Ĭ�Ϲ���
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

