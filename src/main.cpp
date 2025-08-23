#include "update_manager.hpp"
#include <iostream>
#include <sstream>
#include <vector>
#include <string>
#include <algorithm>
#include <cctype>

std::vector<int> parse_user_input(const std::string& input) {
	std::vector<int> indices;
	std::istringstream iss(input);
	std::string token;
	while (iss >> token) {
		try {
			int index = std::stoi(token);
			if (index > 0) indices.push_back(index);
		} catch (const std::exception&) {
			std::cout << "Warning: '" << token << "' is not a valid number. Skipping.\n";
		}
	}
	return indices;
}

std::string get_user_input() {
	std::string input;
	std::getline(std::cin, input);
	// Trim whitespace from both sides.
	input.erase(input.begin(), std::find_if(input.begin(), input.end(), [](unsigned char ch) {
		return !std::isspace(ch);
	}));
	input.erase(std::find_if(input.rbegin(), input.rend(), [](unsigned char ch) {
		return !std::isspace(ch);
	}).base(), input.end());
	return input;
}

void print_instructions() {
	std::cout << "Enter the numbers of the updates you want to hide, separated by spaces.\n";
	std::cout << "Leave blank and press Enter to exit.\n";
}

int main() {
	try {
		windows_update_manager manager;
		manager.search_for_updates();
		if (!manager.has_updates()) {
			std::cout << "No updates available.\nPress Enter to exit...";
			std::cin.get();
			return 0;
		}
		manager.display_available_updates();
		print_instructions();
		std::string user_input = get_user_input();
		if (user_input.empty()) return 0;
		std::vector<int> selected_indices = parse_user_input(user_input);
		if (selected_indices.empty()) return 0;
		std::sort(selected_indices.begin(), selected_indices.end());
		selected_indices.erase(std::unique(selected_indices.begin(), selected_indices.end()), selected_indices.end());
		std::cout << "\nProceeding to hide selected updates...\n\n";
		manager.hide_selected_updates(selected_indices);
	} catch (const com_exception& e) {
		std::cerr << "COM Error: " << e.what() << "\n";
		std::cerr << "This program requires Windows Update services to be running.\n";
		return 1;
	} catch (const std::exception& e) {
		std::cerr << "Error: " << e.what() << "\n";
		return 1;
	} catch (...) {
		std::cerr << "An unknown error occurred.\n";
		return 1;
	}
	std::cout << "\nPress Enter to exit...";
	std::cin.get();
	return 0;
}
