#pragma once
#include <string>
#include <fstream>
#include <vector>
//#include <optional>

class TextFileReader
{
	typedef size_t LineCountType;
private:
	std::wstring filePath; // 文件路径（含文件名）
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

	const std::wstring& getFullTextConstReference() { // 减少拷贝
		return fullText;
	}
	const std::vector<std::wstring>& getLinesVectorConstReference() { // 减少拷贝
		return lineVector;
	}
};

