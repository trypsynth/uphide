#pragma once
#include <windows.h>
#include <comdef.h>
#include <memory>
#include <stdexcept>

class com_initializer {
public:
	com_initializer();
	~com_initializer();
	com_initializer(const com_initializer&) = delete;
	com_initializer& operator=(const com_initializer&) = delete;
	com_initializer(com_initializer&&) = delete;
	com_initializer& operator=(com_initializer&&) = delete;

private:
	bool initialized;
};

class com_exception : public std::runtime_error {
public:
	explicit com_exception(HRESULT hr, const std::string& message = "");

	HRESULT get_hresult() const noexcept { return hr; }

private:
	HRESULT hr;

	static std::string format_error(HRESULT hr, const std::string& message);
};

template<typename T>
class com_ptr {
public:
	com_ptr() = default;

	explicit com_ptr(T* ptr_) : ptr{ptr_} {
		if (ptr) ptr->AddRef();
	}

	com_ptr(const com_ptr& other) : ptr{other.ptr} {
		if (ptr) ptr->AddRef();
	}

	com_ptr(com_ptr&& other) noexcept : ptr{other.ptr} {
		other.ptr = nullptr;
	}

	~com_ptr() {
		if (ptr) ptr->Release();
	}

	com_ptr& operator=(const com_ptr& other) {
		if (this != &other) {
			reset();
			ptr = other.ptr;
			if (ptr) ptr->AddRef();
		}
		return *this;
	}

	com_ptr& operator=(com_ptr&& other) noexcept {
		if (this != &other) {
			reset();
			ptr = other.ptr;
			other.ptr = nullptr;
		}
		return *this;
	}

	T* get() const noexcept { return ptr; }

	T* operator->() const noexcept { return ptr; }

	T& operator*() const noexcept { return *ptr; }

	explicit operator bool() const noexcept { return ptr != nullptr; }

	void reset(T* ptr_ = nullptr) {
		if (ptr) ptr->Release();
		ptr = ptr_;
		if (ptr) ptr->AddRef();
	}

	T** address_of() {
		reset();
		return &ptr;
	}

	T* release() noexcept {
		T* temp = ptr;
		ptr = nullptr;
		return temp;
	}

	template<typename U>
	HRESULT query_interface(U** out) const {
		if (!ptr) return E_POINTER;
		return ptr->QueryInterface(__uuidof(U), reinterpret_cast<void**>(out));
	}

private:
	T* ptr = nullptr;
};

void check_hresult(HRESULT hr, const std::string& context = "");
std::string bstr_to_string(BSTR bstr);
