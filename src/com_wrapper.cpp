#include "com_wrapper.hpp"
#include <sstream>
#include <iomanip>

com_initializer::com_initializer() : initialized{false} {
	HRESULT hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
	if (SUCCEEDED(hr)) initialized = true;
	else if (hr != RPC_E_CHANGED_MODE) throw com_exception(hr, "Failed to initialize COM");
}

com_initializer::~com_initializer() {
	if (initialized) CoUninitialize();
}

com_exception::com_exception(HRESULT hr_, const std::string& message) : std::runtime_error(format_error(hr_, message)), hr{hr_} {}

std::string com_exception::format_error(HRESULT hr, const std::string& message) {
	std::ostringstream oss;
	if (!message.empty()) oss << message << ": ";
	oss << "HRESULT 0x" << std::hex << std::uppercase << hr;
	_com_error error(hr);
	if (error.ErrorMessage()) oss << " (" << static_cast<const char*>(_bstr_t(error.ErrorMessage())) << ")";
	return oss.str();
}

void check_hresult(HRESULT hr, const std::string& context) {
	if (FAILED(hr)) throw com_exception(hr, context);
}

std::string bstr_to_string(BSTR bstr) {
	if (!bstr) return "";
	_bstr_t wrapper(bstr, false);
	return static_cast<const char*>(wrapper);
}
