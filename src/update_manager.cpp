#include "update_manager.hpp"
#include <iostream>
#include <iomanip>

WindowsUpdateManager::WindowsUpdateManager() {
	create_update_session();
	create_update_searcher();
}

void WindowsUpdateManager::create_update_session() {
	HRESULT hr = CoCreateInstance(CLSID_UpdateSession, nullptr, CLSCTX_INPROC_SERVER, IID_IUpdateSession, reinterpret_cast<void**>(session.address_of()));
	check_hresult(hr, "Failed to create Update Session");
}

void WindowsUpdateManager::create_update_searcher() {
	HRESULT hr = session->CreateUpdateSearcher(searcher.address_of());
	check_hresult(hr, "Failed to create Update Searcher");
	_bstr_t client_id(L"uphide");
	hr = searcher->put_ClientApplicationID(client_id);
	check_hresult(hr, "Failed to set client application ID");
}

void WindowsUpdateManager::search_for_updates() {
	std::cout << "Checking for updates...\n";
	perform_search();
	process_search_results();
}

void WindowsUpdateManager::perform_search() {
	_bstr_t search_criteria(L"IsInstalled=0 And IsHidden=0");
	HRESULT hr = searcher->Search(search_criteria, search_result.address_of());
	check_hresult(hr, "Failed to search for updates");
}

void WindowsUpdateManager::process_search_results() {
	HRESULT hr = search_result->get_Updates(update_collection.address_of());
	check_hresult(hr, "Failed to get update collection");
	LONG count = 0;
	hr = update_collection->get_Count(&count);
	check_hresult(hr, "Failed to get update count");
	updates.clear();
	for (LONG i = 0; i < count; ++i) {
		com_ptr<IUpdate> update;
		hr = update_collection->get_Item(i, update.address_of());
		if (SUCCEEDED(hr) && update) {
			std::string title = safe_get_update_title(update.get());
			updates.emplace_back(title, i + 1, update.get());
		}
	}
}

std::string WindowsUpdateManager::safe_get_update_title(IUpdate* update) const {
	try {
		BSTR title_bstr = nullptr;
		HRESULT hr = update->get_Title(&title_bstr);
		if (SUCCEEDED(hr) && title_bstr) {
			std::string result = bstr_to_string(title_bstr);
			SysFreeString(title_bstr);
			return result;
		}
	} catch (...) {
	}
	return "<Title unavailable>";
}

void WindowsUpdateManager::display_available_updates() const {
	if (updates.empty()) {
		std::cout << "No updates available.\n";
		return;
	}
	std::cout << "Available updates:\n";
	for (const auto& update : updates)
		std::cout << update.index << ": " << update.title << "\n";
}

void WindowsUpdateManager::hide_selected_updates(const std::vector<int>& indices) {
	if (indices.empty()) return;
	int hidden_count = 0;
	for (int index : indices) {
		size_t array_index = static_cast<size_t>(index - 1);
		if (array_index >= updates.size()) {
			std::cout << "Warning: Update index " << index << " is out of range. Skipping.\n";
			continue;
		}
		try {
			const auto& update_info = updates[array_index];
			HRESULT hr = update_info.update_ptr->put_IsHidden(VARIANT_TRUE);
			if (SUCCEEDED(hr)) {
				std::cout << "Hid " << update_info.title << "\n";
				++hidden_count;
			} else
				std::cout << "Failed to hide " << update_info.title << ". (HRESULT: 0x" << std::hex << hr << ")\n";
		} catch (const com_exception& e) {
			std::cout << "Error hiding update " << index << ": " << e.what() << "\n";
		} catch (const std::exception& e) {
			std::cout << "Unexpected error hiding update " << index << ": " << e.what() << "\n";
		}
	}
}
