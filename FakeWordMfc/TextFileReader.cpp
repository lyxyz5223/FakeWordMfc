#include "pch.h"
#include "TextFileReader.h"
#include "StringProcess.h"


TextFileReader::TextFileReader(std::wstring filePath)
{
	this->filePath = filePath;
}

bool TextFileReader::open()
{
	if (ifs.is_open())
		ifs.close();
	ifs.open(filePath);
	if (ifs.is_open())
		return true;
	else
		return false;
}

bool TextFileReader::readLines(LineCountType readLineCount)
{
	if (!ifs.is_open())
		return false;
	fullText.clear();
	std::string tmpLine;
	if (readLineCount == 0)
		while (std::getline(ifs, tmpLine))
		{
			auto tmpWStr = str2wstr_2UTF8(tmpLine);
			lineVector.emplace_back(tmpWStr);
			fullText += tmpWStr + L"\n";
		}
	else
		for (LineCountType i = 0; i < readLineCount && std::getline(ifs, tmpLine); i++)
		{
			auto tmpWStr = str2wstr_2UTF8(tmpLine);
			lineVector.emplace_back(tmpWStr);
			fullText += tmpWStr + L"\n";
		}
	return true;
}