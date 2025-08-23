#pragma once
#include "com_wrapper.hpp"
#include <wuapi.h>
#include <vector>
#include <string>

struct update_info {
	std::string title;
	long index;
	com_ptr<IUpdate> update_ptr;

	update_info(const std::string& t, long idx, IUpdate* ptr) : title{t}, index{idx}, update_ptr{ptr} {}
};

class windows_update_manager {
public:
	windows_update_manager();
	~windows_update_manager() = default;
	windows_update_manager(const windows_update_manager&) = delete;
	windows_update_manager& operator=(const windows_update_manager&) = delete;
	windows_update_manager(windows_update_manager&&) = delete;
	windows_update_manager& operator=(windows_update_manager&&) = delete;

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
	std::vector<update_info> updates;

	void create_update_session();
	void create_update_searcher();
	void perform_search();
	void process_search_results();
	std::string get_update_title(IUpdate* update) const;
};
