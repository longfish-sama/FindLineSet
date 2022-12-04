/**
 * @file   Header.hpp
 * @brief
 *
 * @date   2022-11-29
 */
 //col|, row-

#pragma once

#include <iostream>
#include <wchar.h>
#include <Windows.h>
#include <vector>
#include <fstream>
#include <io.h>
#include <string>
#include <list>
#include <xlnt/xlnt.hpp>

using namespace std;
using namespace xlnt;

class node_code
{
public:
	node_code() { loc = "";	number = ""; }
	node_code(string code);
	~node_code();
	void set_code(string set_loc, string set_number);
	string get_code();
	string get_loc() { return loc; }
	string get_number() { return number; }
	int get_tens_digit() { return(atoi(number.c_str()) / 10); }
	int get_units_digits() { return(atoi(number.c_str()) % 10); }
private:
	string loc;
	string number;
};

class node
{
public:
	node();
	node(node_code code, vector<string>& file_list);
	~node();
private:
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

void get_folder_file(std::string path, std::vector<std::string>& files);

string utf2str(const std::string& str);

vector<cell_reference> find_cell(worksheet& ws, string str, int mode = 0);

vector<cell_reference> find_cell(worksheet& ws, range_reference& range, string str, int mode = 0);

vector<string> find_filename(vector<string>& filename_list, string str);

node get_node(worksheet& ws, cell_reference& cr);