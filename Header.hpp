/**
 * @file   Header.hpp
 * @brief
 *
 * @date   2022-11-29
 */
 //note:col|, row-

#pragma once

#define TEST

#include <iostream>
#include <wchar.h>
#include <Windows.h>
#include <vector>
#include <fstream>
#include <io.h>
#include <string>
#include <list>
#include <xlnt/xlnt.hpp>
#include <direct.h>

using namespace std;
using namespace xlnt;

constexpr auto INFO_COUT = "[info]\t";
constexpr auto WARNING_COUT = "[warning]\t";
constexpr auto ERROR_COUT = "[error]\t";

/**
 * @brief all mode used in program
 */
enum class MODE
{
	FIND_CELL_ALL_MATCH,
	FIND_CELL_PART_MATCH,
	LIST2STR_BRIEF,
	LIST2STR_DETAIL
};

/**
 * @brief error type definition
 */
enum class MY_ERROR_TYPE
{
	CLASS_NODE_CONSTRUCT_ERROR,
	CLASS_NODE_CODE_CONSTRUCT_ERROR
};

/**
 * @brief node_code class, code like "Xxxx-xx-xx"
 */
class node_code
{
public:
	node_code() { loc = "";	number = ""; }
	explicit node_code(string code);
	~node_code() = default;
	void set_code(string set_loc, string set_number);
	string get_code() const;
	string get_loc() const { return loc; }
	string get_number() const { return number; }
	int get_row_idx() const;
	int get_col_idx() const;
	bool is_empty() const { return (loc.empty() || number.empty()) ? true : false; }
	bool operator==(const node_code& other) const { return (this->loc == other.loc) && (this->number == other.number); }
	bool operator!=(const node_code& other) { return (this->loc != other.loc) || (this->number != other.number); }
private:
	string loc;
	string number;
};

/**
 * .
 */
class node
{
public:
	node();
	node(node_code& code, vector<string>& file_dir_list);
	~node() = default;
	bool has_up_code();
	bool has_down_code();
	node_code& get_cur_code() { return cur_code; }
	vector<node_code>& get_up_code() { return up_code; }
	vector<node_code>& get_down_code() { return down_code; }
	string get_cur_name() const { return cur_name; }
	vector<string> get_up_name() const { return up_name; }

	/**
	 * @return down name in vector<string>
	 */
	vector<string> get_down_name() const { return down_name; }

	/**
	 * @return comment in string
	 */
	string get_comment() const { return comment; }
	void swap_up_down();
	bool operator==(node const& other) const {
		return (this->cur_code == other.cur_code) &&
			(this->cur_name == other.cur_name) &&
			(this->down_name == other.down_name) &&
			(this->up_name == other.up_name);
	}
private:
	/**
	 * @brief .
	 *
	 * @param ws
	 * @param rg_merged
	 * @param col
	 * @param row
	 * @return
	 */
	string get_cell_value(worksheet& ws, vector<range_reference>& rg_merged, column_t& col, row_t& row) const;
	string workbook_name;
	string worksheet_name;
	string cur_name;
	node_code cur_code;
	vector<string> up_name;
	vector<node_code> up_code;
	vector<string> down_name;
	vector<node_code> down_code;
	string comment;
	string mod_date;
};

void get_folder_file(string path, vector<string>& files);

void copy_folder(string source, string dest);

bool is_node_code(string& str);

string utf2str(const string& str);

vector<cell_reference> find_cell(worksheet& ws, string str, MODE mode);

vector<cell_reference> find_cell(worksheet& ws, range_reference& range, string str, MODE mode);

vector<string> find_filename(vector<string>& filename_list, string str);

list<node> get_single_node_list(node_code& code, vector<string>& files_list);

vector<list<node>> get_node_lists(node_code& code, vector<string>& files_list);

bool is_equal_node_list(list<node>& node_list_1, list<node>& node_list_2);

bool is_found_node_in_node_list(node_code& code, list<node>& node_list);

string list2str(list<node>& node_list, MODE mode);

