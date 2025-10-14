#pragma once
#include <string>
#include <fstream>
#include <vector>
//#include <optional>

class TextFileReader
{
	typedef size_t LineCountType;
private:
	std::wstring filePath; // �ļ�·�������ļ�����
	std::ifstream ifs;
	std::vector<std::wstring> lineVector;
	std::wstring fullText;

public:
	~TextFileReader() {
		if (ifs.is_open())
			ifs.close();
	}
	TextFileReader(std::wstring fileWithPath = L"");
	void setFilePath(std::wstring filePath) {
		this->filePath = filePath;
	}
	std::wstring getFilePath() const {
		return filePath;
	}
	bool open();

	bool readLines(LineCountType readLineCount = 0);

	std::wstring getFullText() const {
		return fullText;
	}

	std::vector<std::wstring> getLinesVector() const {
		return lineVector;
	}

	const std::wstring& getFullTextConstReference() { // ���ٿ���
		return fullText;
	}
	const std::vector<std::wstring>& getLinesVectorConstReference() { // ���ٿ���
		return lineVector;
	}
};

