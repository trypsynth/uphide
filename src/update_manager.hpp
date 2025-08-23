#pragma once
#include "com_wrapper.hpp"
#include <wuapi.h>
#include <vector>
#include <string>

struct UpdateInfo {
	std::string title;
	long index;
	com_ptr<IUpdate> update_ptr;

	UpdateInfo(const std::string& t, long idx, IUpdate* ptr) : title{t}, index{idx}, update_ptr{ptr} {}
};

class WindowsUpdateManager {
public:
	WindowsUpdateManager();
	~WindowsUpdateManager() = default;
	WindowsUpdateManager(const WindowsUpdateManager&) = delete;
	WindowsUpdateManager& operator=(const WindowsUpdateManager&) = delete;
	WindowsUpdateManager(WindowsUpdateManager&&) = delete;
	WindowsUpdateManager& operator=(WindowsUpdateManager&&) = delete;
	void search_for_updates();
	void display_available_updates() const;
	void hide_selected_updates(const std::vector<int>& indices);

	size_t get_update_count() const { return updates.size(); }

	bool has_updates() const { return !updates.empty(); }

private:
	com_initializer com_init;
	com_ptr<IUpdateSession> session;
	com_ptr<IUpdateSearcher> searcher;
	com_ptr<ISearchResult> search_result;
	com_ptr<IUpdateCollection> update_collection;
	std::vector<UpdateInfo> updates;

	void create_update_session();
	void create_update_searcher();
	void perform_search();
	void process_search_results();
	std::string safe_get_update_title(IUpdate* update) const;
};
