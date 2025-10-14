#pragma once


#define COM_CATCH_NORESULT(tip) catch (_com_error err) { \
        std::wstring errMsg = err.ErrorMessage(); \
        std::wstring msg = tip; \
        msg += '\n'; \
        msg += errMsg; \
		MessageBeep(MB_ICONERROR); \
        ::MessageBox(0, msg.c_str(), L"Error", MB_ICONERROR); \
    }
#define COM_CATCH_ALL_NORESULT(tip) catch (...) { \
        std::wstring msg = tip; \
        msg += L"\nUnknown Error."; \
		MessageBeep(MB_ICONERROR); \
        ::MessageBox(0, msg.c_str(), L"Error", MB_ICONERROR); \
    }
#define COM_CATCH(tip, result) catch (_com_error err) { \
        std::wstring errMsg = err.ErrorMessage(); \
        std::wstring msg = tip; \
        msg += '\n'; \
        msg += errMsg; \
		MessageBeep(MB_ICONERROR); \
        ::MessageBox(0, msg.c_str(), L"Error", MB_ICONERROR); \
        result = ErrorResult<ExtractValueType<decltype(result)>::type>(errMsg); \
    }
#define COM_CATCH_ALL(tip, result) catch (...) { \
        std::wstring msg = tip; \
        msg += L"\nUnknown Error."; \
		MessageBeep(MB_ICONERROR); \
        ::MessageBox(0, msg.c_str(), L"Error", MB_ICONERROR); \
        result = ErrorResult<ExtractValueType<decltype(result)>::type>(L"Unknown Error."); \
    }
#define COM_CATCH_NOMSGBOX(tip, result) catch (_com_error err) { \
        std::wstring errMsg = err.ErrorMessage(); \
        std::wstring msg = tip; \
        msg += '\n'; \
        msg += errMsg; \
        result = ErrorResult<ExtractValueType<decltype(result)>::type>(errMsg); \
    }
#define COM_CATCH_ALL_NOMSGBOX(tip, result) catch (...) { \
        std::wstring msg = tip; \
        msg += L"\nUnknown Error."; \
        result = ErrorResult<ExtractValueType<decltype(result)>::type>(L"Unknown Error."); \
    }

template <typename DataType>
class Result
{
public:
    typedef long long CodeType;
    typedef std::wstring MsgType;
private:
    bool success = true;
    CodeType code = 0;
    MsgType msg = L"";
    DataType data;

public:
    Result(bool success, CodeType code, DataType data, MsgType msg) {
        this->success = success;
        this->code = code;
        this->data = data;
        this->msg = msg;
    }
    Result(bool success, CodeType code, DataType data) {
        this->success = success;
        this->code = code;
        this->data = data;
    }
    Result(bool success, CodeType code, MsgType msg) {
        this->success = success;
        this->code = code;
        this->msg = msg;
    }

    Result(bool success, CodeType code) {
        this->success = success;
        this->code = code;
    }
    Result(bool success) {
        this->success = success;
    }
    static Result<DataType> CreateFromHRESULT(HRESULT hr, DataType data) {
        if (SUCCEEDED(hr))
        {
            return SuccessResult(hr, data);
        }
        else
        {
            _com_error err(hr);
            std::wstring msg = err.ErrorMessage();
            //auto msg = HResultToWString(hr);
            if (hr == E_NOINTERFACE)
                msg = L"No such interface";
            return ErrorResult<DataType>(hr, msg);
        }
    }
    const auto& getSuccess() const noexcept {
        return success;
    }
    const auto& getCode() const noexcept {
        return code;
    }

    const auto& getMessage() const noexcept {
        return msg;
    }

    const auto& getData() const noexcept {
        return data;
    }

    operator bool() {
        return success;
    }
};


template <typename DataType>
class ErrorResult : public Result<DataType>
{
public:
    ErrorResult(Result::CodeType code, Result::MsgType msg) : Result(false, code, msg) {}
    ErrorResult(Result::MsgType msg) : Result(false, 0, msg) {}
    ErrorResult(Result::CodeType code) : Result(false, code, Result::MsgType()) {}
    ErrorResult() : Result(false, 0, Result::MsgType()) {}
};

template <typename DataType>
class SuccessResult : public Result<DataType>
{
public:
    SuccessResult(Result::CodeType code, DataType data) : Result(true, code, data) {}
    SuccessResult(DataType data) : Result(true, 0, data) {}
    SuccessResult(Result::CodeType code) : Result(true, code, DataType()) {}
    SuccessResult() : Result(true, 0, DataType()) {}
};


// 类型萃取
// 类型萃取模板
template<typename T>
struct ExtractValueType {
    using type = T;
};

template<typename T>
struct ExtractValueType<Result<T>> {
    using type = T;
};
