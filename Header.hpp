/**
 * @file   Header.hpp
 * @brief
 *
 * @date   2022-11-29
 */
 //note:col|, row-

#pragma once

#define TEST

#define INFO "[info]\t"
#define WARNING "[warning]\t"
#define ERROR "[error]\t"

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

enum MODE
{
	FIND_CELL_ALL_MATCH,
	FIND_CELL_PART_MATCH
};

enum ERROR_TYPE
{
	CLASS_NODE_CONSTRUCT_ERROR
};

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
	bool operator==(const node_code& other) { return (this->loc == other.loc) && (this->number == other.number); }
private:
	string loc;
	string number;
};

class node
{
public:
	node();
	node(node_code& code, vector<string>& file_dir_list);
	~node() = default;
	bool has_up_code() { return up_code.is_empty() ? false : true; }
	bool has_down_code() { return down_code.is_empty() ? false : true; }
	node_code& get_cur_code() { return cur_code; }
	node_code& get_up_code() { return up_code; }
	node_code& get_down_code() { return down_code; }
	string get_cur_name() { return cur_name; }
	string get_up_name() { return up_name; }
	string get_down_name() { return down_name; }
	void swap_up_down();
private:
	string get_cell_value(worksheet& ws, vector<range_reference>& rg_merged, column_t& col, row_t& row) const;
	string workbook_name;
	string worksheet_name;
	string cur_name;
	node_code cur_code;
	string up_name;
	node_code up_code;
	string down_name;
	node_code down_code;
	string comment;
};

void get_folder_file(string path, vector<string>& files);

void copy_folder(string source, string dest);

string utf2str(const string& str);

vector<cell_reference> find_cell(worksheet& ws, string str, MODE mode);

vector<cell_reference> find_cell(worksheet& ws, range_reference& range, string str, MODE mode);

vector<string> find_filename(vector<string>& filename_list, string str);

list<node> get_node_list(node_code& code, vector<string>& file_dir_list);