#pragma once
#include <string>
#include <unordered_map>

enum class SupportedFileType {
	Word,
	Txt
};

struct SupportedFile {
	std::wstring filePath;
	std::wstring fileName;
	std::wstring fileExtension;
	SupportedFileType fileType = SupportedFileType::Txt;
};
struct ReplaceMode {
	enum _ReplaceMode {
		SlackOff,
		Work
	} mode;
	static const int replaceModeCount = 2;

#define ReplaceMode_MapPair(str) { str, L#str }
	std::unordered_map<_ReplaceMode, std::wstring> modeWStringMap{
		ReplaceMode_MapPair(SlackOff),
		ReplaceMode_MapPair(Work)
	};
#undef ReplaceMode_MapPair
#define ReplaceMode_MapPair(str) { str, #str }
	std::unordered_map<_ReplaceMode, std::string> modeStringMap{
		ReplaceMode_MapPair(SlackOff),
		ReplaceMode_MapPair(Work)
	};
#undef ReplaceMode_MapPair

	ReplaceMode(const ReplaceMode& mode)
		: mode(mode.mode)
	{

	}
	ReplaceMode(const _ReplaceMode& mode)
		: mode(mode)
	{

	}
	_ReplaceMode next() {
		return (_ReplaceMode)((mode + 1) % replaceModeCount);
	}

	template <typename T>
	static ReplaceMode fromNumber(T number) {
		return ReplaceMode(number % count);
	}

	int toInt() {
		return static_cast<int>(mode);
	}

	std::wstring toWString() {
		return modeWStringMap[mode];
	}
	std::string toString() {
		return modeStringMap[mode];
	}
	ReplaceMode& operator=(const _ReplaceMode& mode) {
		this->mode = mode;
		return *this;
	}
	ReplaceMode& operator=(const ReplaceMode& mode) {
		this->mode = mode.mode;
		return *this;
	}

	operator _ReplaceMode() {
		return mode;
	}

	operator std::wstring() {
		return toWString();
	}
	operator std::string() {
		return toString();
	}
};

union CtrlHotKey {
	DWORD hotKey;
	struct {
		WORD preserved;
		unsigned char modifiers;
		unsigned char virtualKeyCode;
	};
};

struct HotKey {
	// 虚拟密钥代码采用低序字节，修饰符标志位于高阶字节中
	UINT virtualKeyCode; // 虚拟密钥代码
	UINT modifiers; // 修饰符标志

	HotKey(UINT virtualKeyCode, UINT modifiers, bool bNoRepeat = true)
		: virtualKeyCode(virtualKeyCode),
		modifiers(modifiers | (bNoRepeat ? MOD_NOREPEAT : 0x0)) {
	}
	CtrlHotKey toCtrlHotKey() const {
		CtrlHotKey chk{ 0 };
		chk.virtualKeyCode = virtualKeyCode;
		if (modifiers & MOD_SHIFT)
			chk.modifiers |= HOTKEYF_SHIFT;
		if (modifiers & MOD_CONTROL)
			chk.modifiers |= HOTKEYF_CONTROL;
		if (modifiers & MOD_ALT)
			chk.modifiers |= HOTKEYF_ALT;
		return chk;
	}
	static HotKey fromCtrlHotKey(DWORD ctrlHotKey, bool bNoRepeat = true) {
		HotKey hk{ 0, 0 };
		hk.virtualKeyCode = (ctrlHotKey & 0xFF);
		WORD ctrlMods = ((ctrlHotKey >> 8) & 0xFF);
		if (ctrlMods & HOTKEYF_SHIFT)
			hk.modifiers |= MOD_SHIFT;
		if (ctrlMods & HOTKEYF_CONTROL)
			hk.modifiers |= MOD_CONTROL;
		if (ctrlMods & HOTKEYF_ALT)
			hk.modifiers |= MOD_ALT;
		if (bNoRepeat)
			hk.modifiers |= MOD_NOREPEAT;
		return hk;
	}
};

//HotKey MakeHotKey(DWORD hotKey) {
//	return HotKey{ hotKey };
//}
//
//HotKey MakeHotKey(WORD virtualKeyCode, WORD modifiers) {
//	return HotKey{ virtualKeyCode, modifiers };
//}
struct WordDetectMode {
	enum _WordDetectMode {
		AppEvent,
		LoopCheck,
		HotKey,
		DllInject
	} mode;
	static const int modeCount = 4;

#define WordDetectMode_MapPair(str) { str, L#str }
	std::unordered_map<_WordDetectMode, std::wstring> modeWStringMap{
		WordDetectMode_MapPair(AppEvent),
		WordDetectMode_MapPair(LoopCheck),
		WordDetectMode_MapPair(HotKey),
		WordDetectMode_MapPair(DllInject)
	};
#undef WordDetectMode_MapPair
#define WordDetectMode_MapPair(str) { str, #str }
	std::unordered_map<_WordDetectMode, std::string> modeStringMap{
		WordDetectMode_MapPair(AppEvent),
		WordDetectMode_MapPair(LoopCheck),
		WordDetectMode_MapPair(HotKey),
		WordDetectMode_MapPair(DllInject)
	};
#undef WordDetectMode_MapPair

	WordDetectMode(const WordDetectMode& mode)
		: mode(mode.mode)
	{

	}
	WordDetectMode(const _WordDetectMode& mode)
		: mode(mode)
	{

	}
	_WordDetectMode next() {
		return static_cast<_WordDetectMode>((mode + 1) % modeCount);
	}

	template <typename T>
	static WordDetectMode fromNumber(T number) {
		return WordDetectMode(static_cast<_WordDetectMode>(static_cast<int>(number % modeCount)));
	}

	int toInt() {
		return static_cast<int>(mode);
	}

	std::wstring toWString() {
		return modeWStringMap[mode];
	}
	std::string toString() {
		return modeStringMap[mode];
	}
	WordDetectMode& operator=(const _WordDetectMode& mode) {
		this->mode = mode;
		return *this;
	}
	WordDetectMode& operator=(const WordDetectMode& mode) {
		this->mode = mode.mode;
		return *this;
	}

	operator _WordDetectMode() {
		return mode;
	}

	operator std::wstring() {
		return toWString();
	}
	operator std::string() {
		return toString();
	}


};


#define ID_HOTKEY_DEFAULT_SWITCH_REPLACE_MODE 1037
#define DEFAULT_HOTKEY_SWITCHREPLACEMODE HotKey('S', MOD_NOREPEAT | MOD_CONTROL | MOD_ALT | MOD_SHIFT)
#define DEFAULT_OIRATIO 2
#define DEFAULT_WORDDETECTMODE WordDetectMode::AppEvent