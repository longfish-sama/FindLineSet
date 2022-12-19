﻿/**
 * @file   Header.cpp
 * @brief
 *
 * @date   2022-11-29
 */

#include "Header.hpp"

 /**
  * @brief .
  *
  * @param folder: folder name without '/'or'\\'
  * @param files: vector to hold results
  */
void get_folder_file(string folder, vector<string>& files)
{
	intptr_t hFile = 0;
	struct _finddata_t fileinfo;
	string p;
	if ((hFile = _findfirst(p.assign(folder).append("/*").c_str(), &fileinfo)) != -1)
	{
		do
		{
			if (fileinfo.attrib & _A_SUBDIR)
			{
				if (strcmp(fileinfo.name, ".") != 0 && strcmp(fileinfo.name, "..") != 0)
				{
					get_folder_file(p.assign(folder).append("/").append(fileinfo.name), files);
				}
			}
			else
			{
				files.push_back(p.assign(folder).append("/").append(fileinfo.name));
			}

		} while (_findnext(hFile, &fileinfo) == 0);
		_findclose(hFile);
	}
}

void copy_folder(string source_folder, string dest_folder)//todo: delete then copy
{
	if (_access(dest_folder.c_str(), 0) != 0 && _mkdir(dest_folder.c_str()) != 0)
	{
		return;
	}

	vector<string> source_files;
	vector<string> dest_files;
	get_folder_file(source_folder, source_files);
	dest_files = source_files;
	for (auto& item : dest_files)
	{
		if (item.find(".xlsx") != string::npos && item.find("旧") == string::npos && item.find("副本") == string::npos)
		{
			string tmp = "";
			for (auto c : item)
			{
				if (c == '/' || c == ':')
				{
					c = '-';
				}
				if (c >= 0 && c <= 127)
				{
					tmp = tmp + c;
				}
			}
			item = dest_folder + "/" + tmp + ".zip";
		}
		else
		{
			item.clear();
		}
#ifdef TEST
		cout << INFO_COUT + item << endl;
#endif
	}

	for (size_t i = 0; i < dest_files.size(); i++)
	{
		if (!dest_files.at(i).empty())
		{
			ifstream infile(source_files[i], ios::in | ios::binary);
			if (!infile.is_open() || infile.bad() || infile.fail())
			{
				cout << WARNING_COUT + source_files[i] + " fail to open" << endl;
				infile.close();
				continue;
			}
			ofstream outfile(dest_files[i], ios::out | ios::binary);
			if (!outfile.is_open() || outfile.bad() || outfile.fail())
			{
				cout << WARNING_COUT + dest_files[i] + " fail to open" << endl;
				infile.close();
				outfile.close();
				continue;
			}
			outfile << infile.rdbuf();
			outfile.close();
			infile.close();
		}
	}

	return;
}

bool is_node_code(string& str)
{
	size_t index_1 = str.find_first_of("-");
	size_t index_2 = str.find_last_of("-");
	if (index_1 < str.length() && index_1 < index_2 && index_2 < str.length() &&
		((str[0] >= 0x41 && str[0] <= 0x5a) || (str[0] >= 0x61 && str[0] <= 0x7a)))
	{
		return true;
	}
	else
	{
		return false;
	}
}

/**
 * @brief .
 *
 * @param str
 * @return
 */
string utf2str(const string& str)
{
	int nwLen = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, nullptr, 0);

	auto* pwBuf = new wchar_t[nwLen + 1];
	memset(pwBuf, 0, nwLen * 2 + 2);

	MultiByteToWideChar(CP_UTF8, 0, str.c_str(), str.length(), pwBuf, nwLen);

	int nLen = WideCharToMultiByte(CP_ACP, 0, pwBuf, -1, nullptr, NULL, nullptr, nullptr);

	auto* pBuf = new char[nLen + 1];
	memset(pBuf, 0, nLen + 1);

	WideCharToMultiByte(CP_ACP, 0, pwBuf, nwLen, pBuf, nLen, nullptr, nullptr);

	string retStr = pBuf;

	delete[] pBuf;
	delete[] pwBuf;

	pBuf = nullptr;
	pwBuf = nullptr;

	return retStr;
}

/**
 * @brief .
 *
 * @param ws
 * @param str
 * @param find_mode
 * @return
 */
vector<cell_reference> find_cell(worksheet& ws, string str,
	MODE find_mode = MODE::FIND_CELL_PART_MATCH,
	MODE return_mode = MODE::FIND_CELL_RET_VALUED_CELL)
{
	vector<cell_reference> cr;
	string tmp_str = "";

	switch (find_mode)
	{
	case MODE::FIND_CELL_PART_MATCH:
		for (auto row : ws.rows(false))
		{
			for (auto val : row)
			{
				tmp_str = utf2str(val.to_string());
				if (tmp_str.find(str) != string::npos)
				{
					switch (return_mode)
					{
					case MODE::FIND_CELL_RET_ALL_CELL:
						if (val.is_merged())
						{
							vector<cell_reference>tmp_cr = get_cells_in_merged_range(ws, val.reference());
							for (auto cr_val : tmp_cr)
							{
								cr.push_back(cr_val);
							}
						}
						else
						{
							cr.push_back(val.reference());
						}
						break;
					case MODE::FIND_CELL_RET_VALUED_CELL:
						cr.push_back(val.reference());
						break;
					default:
						break;
					}
				}
			}
		}
		break;

	case MODE::FIND_CELL_ALL_MATCH:
		for (auto row : ws.rows(false))
		{
			for (auto val : row)
			{
				tmp_str = utf2str(val.to_string());
				if (tmp_str == str)
				{
					switch (return_mode)
					{
					case MODE::FIND_CELL_RET_ALL_CELL:
						if (val.is_merged())
						{
							vector<cell_reference>tmp_cr = get_cells_in_merged_range(ws, val.reference());
							for (auto cr_val : tmp_cr)
							{
								cr.push_back(cr_val);
							}
						}
						else
						{
							cr.push_back(val.reference());
						}
						break;
					case MODE::FIND_CELL_RET_VALUED_CELL:
						cr.push_back(val.reference());
						break;
					default:
						break;
					}
				}
			}
		}
		break;

	default:
		break;
	}

	return cr;
}

vector<cell_reference> find_cell(worksheet& ws, range_reference& rg, string str,
	MODE find_mode = MODE::FIND_CELL_PART_MATCH) //todo: change to private function
{
	vector<cell_reference> cr;
	string tmp;

	switch (find_mode)
	{
	case MODE::FIND_CELL_PART_MATCH:
		for (row_t row = rg.top_left().row(); row <= rg.bottom_right().row(); row++)
		{
			for (column_t col = rg.top_left().column(); col <= rg.bottom_right().column(); col++)
			{
				tmp = utf2str(ws.cell(col, row).to_string());
				if (tmp.find(str) != string::npos)
				{
					cr.push_back(ws.cell(col, row).reference());
				}
			}
		}
		break;

	case MODE::FIND_CELL_ALL_MATCH:
		for (row_t row = rg.top_left().row(); row <= rg.bottom_right().row(); row++)
		{
			for (column_t col = rg.top_left().column(); col <= rg.bottom_right().column(); col++)
			{
				tmp = utf2str(ws.cell(col, row).to_string());
				if (tmp == str)
				{
					cr.push_back(ws.cell(col, row).reference());
				}
			}
		}
		break;

	default:
		break;
	}

	return cr;
}

vector<cell_reference> get_cells_in_merged_range(worksheet& ws, const cell_reference& tl_cell)
{
	vector<cell_reference> ret;
	if (ws.cell(tl_cell).is_merged())
	{
		vector<range_reference>rg_merged = ws.merged_ranges();
		for (auto const& val : rg_merged)
		{
			if (val.top_left() == tl_cell)
			{
				for (size_t r = 0; r < val.height(); r++)
				{
					for (size_t c = 0; c < val.width(); c++)
					{
						ret.push_back(cell_reference(tl_cell.column() + c, tl_cell.row() + r));
					}
				}
				break;
			}
		}
	}
	else
	{
		ret.push_back(tl_cell);
	}
	return ret;
}

string get_merged_cell_value(worksheet& ws, vector<range_reference>& rg_merged, const column_t& col, const row_t& row)
{
	if (ws.cell(col, row).is_merged() && (!ws.cell(col, row).has_value()))
	{
		for (size_t i = 0; i < rg_merged.size(); i++)
		{
			if (rg_merged.at(i).top_left().column() <= col && col <= rg_merged.at(i).top_right().column() &&
				rg_merged.at(i).top_left().row() <= row && row <= rg_merged.at(i).bottom_left().row())
			{
				return utf2str(ws.cell(rg_merged.at(i).top_left()).to_string());
			}
		}
	}
	else
	{
		return utf2str(ws.cell(col, row).to_string());
	}

	return "";
}



vector<string> find_filename(vector<string>& filename_list, string str)
{
	vector<string> res;
	for (size_t i = 0; i < filename_list.size(); i++)
	{
		size_t j = filename_list[i].find(str);
		if (j != string::npos && filename_list[i].find("~$") == string::npos &&
			(filename_list.at(i)[j + str.length()] < 0x30 || filename_list.at(i)[j + str.length()] > 0x39))
		{
			res.push_back(filename_list[i]);
		}
	}
	return res;
}

/**
 * @brief .
 *
 * @param code: node code of 1st node object
 * @param files_list
 * @return
 */
list<node> get_single_node_list(node_code& code, vector<string>& files_list)
{
	if (code.is_empty())
	{
		return list<node>();
	}
	else
	{
		list<node> ret;
		node n0(code, files_list);
		node n1 = n0;
		ret.push_back(n0);

		while (n0.has_up_code())
		{
			node n_up(n0.get_up_code().at(0), files_list);
			ret.push_front(n_up);
			for (auto const& val : n_up.get_up_code())
			{

				if (val == n0.get_cur_code())
				{
					n_up.swap_up_down();
					break;
				}
			}
			n0 = n_up;
		}

		while (n1.has_down_code())
		{
			node n_down(n1.get_down_code().at(0), files_list);
			ret.push_back(n_down);
			for (auto const& val : n_down.get_down_code())
			{
				if (val == n1.get_cur_code())
				{
					n_down.swap_up_down();
					break;
				}
			}
			n1 = n_down;
		}

		return ret;
	}
}

vector<list<node>> get_node_lists(node_code& code, vector<string>& files_list)
{
	vector<list<node>> ret;
	list<node_code> codes_unused;
	vector<node_code> codes_used;
	if (!code.is_empty())
	{
		codes_unused.push_back(code);
	}
	while (!codes_unused.empty())
	{
		list<node> tmp = get_single_node_list(codes_unused.front(), files_list); //remember to pop front
		codes_used.push_back(codes_unused.front()); //remember to pop front

		bool is_unique_list = none_of(ret.begin(), ret.end(), [&tmp](auto& val) {
			return tmp == val;
			});
		if (is_unique_list)
		{
			ret.push_back(tmp);

			for (auto val : tmp) //traverse current temp node list to find multi code node
			{
				if (val.get_down_code().size() > 1) //find a node with multi down code
				{
					for (auto val1 : val.get_down_code()) //traverse node's down code
					{
						if (val1.is_empty())
						{
							continue;
						}
						bool is_unique_code = true; //unused code flag
						if (has_node_in_node_list(val1, tmp)) //traverse current temp list
						{
							is_unique_code = false; //false if code is appeared in current node list
						}
						else
						{
							is_unique_code = none_of(codes_used.begin(), codes_used.end(), [&val1](auto& val2) {
								return val2 == val1;
								}); //false if code is used
						}
						if (is_unique_code)
						{
							codes_unused.push_back(val1);
						}
					}
				}
				if (val.get_up_code().size() > 1) //find a node with multi up code
				{
					for (auto val1 : val.get_up_code()) //traverse node's up code
					{
						if (val1.is_empty())
						{
							continue;
						}
						bool is_unique_code = true; //unused code flag
						if (has_node_in_node_list(val1, tmp)) //traverse current temp list
						{
							is_unique_code = false; //false if code is appeared in current node list
						}
						else
						{
							is_unique_code = none_of(codes_used.begin(), codes_used.end(), [&val1](auto& val2) {
								return val2 == val1;
								}); //false if code is used
						}
						if (is_unique_code)
						{
							codes_unused.push_back(val1);
						}
					}
				}

			}
		}

		codes_unused.pop_front(); //remember to pop front
	}
	return ret;
}

bool is_equal_node_list(list<node>& node_list_1, list<node>& node_list_2)
{
	if (node_list_1.size() != node_list_2.size())
	{
		return false;
	}
	else
	{
		bool ret_front = true;
		bool ret_back = true;
		vector<node> v1;
		vector<node> v2;
		for (auto const& val : node_list_1)
		{
			v1.push_back(val); //convert list to vector
		}
		for (auto const& val : node_list_2)
		{
			v2.push_back(val);
		}
		for (size_t i = 0; i < v1.size(); i++)
		{
			ret_front = ret_front && (v1.at(i) == v2.at(i));
			ret_back = ret_back && (v1.at(i) == v2.at(v1.size() - 1 - i));
		}
		return ret_front || ret_back;
	}
}

bool has_node_in_node_list(node_code& code, list<node>& node_list)
{
	return any_of(node_list.begin(), node_list.end(), [&code](auto node) {
		return code == node.get_cur_code();
		});
}

/**
 * @brief .
 *
 * @param node_list
 * @param mode
 * @return
 */
string list2str(list<node>& node_list, MODE mode = MODE::LIST2STR_BRIEF)
{
	string ret = "";
	switch (mode)
	{
	case MODE::LIST2STR_BRIEF:
		for (auto val : node_list)
		{
			ret = ret + val.get_cur_name() + " | " +
				val.get_cur_code().get_code() + " | " +
				val.get_comment() + "\n";
		}
		break;

	case MODE::LIST2STR_DETAIL:
		for (auto val : node_list)
		{
			string up = "";
			string down = "";
			for (auto const& up_str : val.get_up_name())
			{
				up = up + "↑" + up_str + " ";
			}
			for (auto const& down_str : val.get_down_name())
			{
				down = down + "↓" + down_str + " ";
			}
			ret = ret + val.get_cur_name() + " | " +
				val.get_cur_code().get_code() + " | " +
				val.get_comment() + " | " +
				up + "| " +
				down + "\n";
		}
		break;

	default:
		break;
	}
	return ret + "\n";
}

string list2str(vector<list<node>>& node_lists, MODE mode = MODE::LIST2STR_BRIEF)
{
	string retval = "";
	switch (mode)
	{
	case MODE::LIST2STR_BRIEF:
		for (size_t i = 0; i < node_lists.size(); i++)
		{
			retval = retval + "line" + to_string(i + 1) + "/" + to_string(node_lists.size()) + "\n";
			for (auto val : node_lists.at(i))
			{
				retval = retval + val.get_cur_name() + " | " +
					val.get_cur_code().get_code() + " | " +
					val.get_comment() + "\n";
			}
			retval = retval + "\n";
		}
		break;

	case MODE::LIST2STR_DETAIL:
		for (size_t i = 0; i < node_lists.size(); i++)
		{

			retval = retval + "line" + to_string(i + 1) + "/" + to_string(node_lists.size()) + "\n";
			for (auto val : node_lists.at(i))
			{
				string up = "";
				string down = "";
				for (auto const& up_str : val.get_up_name())
				{
					up = up + "↑" + up_str + " ";
				}
				for (auto const& down_str : val.get_down_name())
				{
					down = down + "↓" + down_str + " ";
				}
				retval = retval + val.get_cur_name() + " | " +
					val.get_cur_code().get_code() + " | " +
					val.get_comment() + " | " +
					up + "| " +
					down + "\n";
			}
			retval = retval + "\n";
		}
		break;

	default:
		break;
	}

	return retval;
}

node_code::node_code(string code)
{
	size_t index_1 = code.find_first_of("-");
	size_t index_2 = code.find_last_of("-");
	if (index_1 < code.length() && index_1 < index_2 && index_2 < code.length())
	{
		loc = code.substr(0, index_2);
		number = code.substr(index_2 + 1);
		if (number.back() < 0x30 || number.back() > 0x39)
		{
			number = number.substr(0, number.size() - 1); //ignore last none-number char
		}
		//for (size_t i = 0; i < loc.length(); i++)
		//{
		//	loc[i] = toupper(loc[i]);
		//}
	}
	else
	{
		loc = "";
		number = "";
#ifdef TEST
		cout << "node_code construct error" << endl;
#endif 
	}
}

node_code::node_code(worksheet& ws, cell_reference& cr, vector<string>& files_list)
{
	if ((ws.cell(cr).has_value() || ws.cell(cr).is_merged()) &&
		(ws.cell(1, cr.row()).has_value() || ws.cell(1, cr.row()).is_merged()))
	{
		vector<range_reference>rg_merged = ws.merged_ranges();

		//get row index
		int row_idx = 0;
		try
		{
			row_idx = stoi(get_merged_cell_value(ws, rg_merged, 1, cr.row()));
		}
		catch (const invalid_argument& ex) //error handle
		{
			cout << "node_code construct invalid_argument: " << ex.what() << endl;
			row_idx = 0;
			throw MY_ERROR_TYPE::CLASS_NODE_CODE_CONSTRUCT_ERROR;
		}
		catch (const out_of_range& ex) //error handle
		{
			cout << "node_code construct out_of_range: " << ex.what() << endl;
			row_idx = 0;
			throw MY_ERROR_TYPE::CLASS_NODE_CODE_CONSTRUCT_ERROR;
		}

		//get column index
		int row_decrement = 0;
		string tmp_row = "";
		int col_idx = 0;
		do
		{
			tmp_row = utf2str(ws.cell(cr.column(), cr.row() - row_decrement).to_string());
			row_decrement++;
		} while (tmp_row != "1" && tmp_row != "2" && tmp_row != "3" && tmp_row != "4" && tmp_row != "5" &&
			tmp_row != "6" && tmp_row != "7" && tmp_row != "8" && tmp_row != "9" && tmp_row != "10" &&
			cr.row() - row_decrement > 0);
		if (!tmp_row.empty())
		{
			try
			{
			col_idx = stoi(tmp_row);
			}
			catch (const invalid_argument& ex) //error handle
			{
				cout << "node_code construct invalid_argument: " << ex.what() << endl;
				col_idx = 0;
				throw MY_ERROR_TYPE::CLASS_NODE_CODE_CONSTRUCT_ERROR;
			}
			catch (const out_of_range& ex) //error handle
			{
				cout << "node_code construct out_of_range: " << ex.what() << endl;
				col_idx = 0;
				throw MY_ERROR_TYPE::CLASS_NODE_CODE_CONSTRUCT_ERROR;
			}

		}
		else
		{
			col_idx = 0;
		}

		//get number part
		if (row_idx == 0 || col_idx == 0)
		{
			node_code::number = "";
		}
		else
		{
			node_code::number = to_string((row_idx - 1) * 10 + col_idx);
		}

		//get loc part
		string title = utf2str(ws.cell("A1").to_string());
		if (size_t idx1 = title.find_first_of(" "); idx1 < title.length())
		{
			title = title.substr(0, idx1);
		}
		else
		{
			title = "";
		}
		node_code::loc = title;
	}
}


void node_code::set_code(string set_loc, string set_number)
{
	loc = set_loc;
	number = set_number;
	for (size_t i = 0; i < loc.length(); i++)
	{
		loc[i] = toupper(loc[i]);
	}
}

string node_code::get_code() const
{
	if (loc.empty() || number.empty())
	{
		return "";
	}
	else
	{
		return loc + "-" + number;
	}
}

int node_code::get_row_idx() const
{
	int tmp = 0;

	try
	{
		tmp = stoi(number);
	}
	catch (const invalid_argument& ex) //error handle
	{
		cout << "node_code::get_row_idx() invalid_argument: " << ex.what() << endl;
		throw MY_ERROR_TYPE::CLASS_NODE_CODE_MEMB_FUNC_ERROR;
		return 0;
	}
	catch (const out_of_range& ex) //error handle
	{
		cout << "node_code::get_row_idx() out_of_range: " << ex.what() << endl;
		throw MY_ERROR_TYPE::CLASS_NODE_CODE_MEMB_FUNC_ERROR;
		return 0;
	}

	if (tmp % 10 == 0)
	{
		return tmp / 10;
	}
	else
	{
		return tmp / 10 + 1;
	}
}

int node_code::get_col_idx() const
{
	int tmp = 0;

	try
	{
		tmp = stoi(number);
	}
	catch (const invalid_argument& ex) //error handle
	{
		cout << "node_code::get_row_idx() invalid_argument: " << ex.what() << endl;
		throw MY_ERROR_TYPE::CLASS_NODE_CODE_MEMB_FUNC_ERROR;
		return 0;
	}
	catch (const out_of_range& ex) //error handle
	{
		cout << "node_code::get_row_idx() out_of_range: " << ex.what() << endl;
		throw MY_ERROR_TYPE::CLASS_NODE_CODE_MEMB_FUNC_ERROR;
		return 0;
	}

	if (tmp % 10 == 0)
	{
		return 10;
	}
	else
	{
		return tmp % 10;
	}
}


node::node(node_code& code, vector<string>& files_list)
{
	//use node code to find file path
	vector<string> file = find_filename(files_list, code.get_loc());
	if (file.size() > 1) //error handle
	{
		cerr << "node construct: multi file found: " << code.get_loc() << endl;
		for (auto val : file)
		{
			cerr << val << endl;
		}
		throw MY_ERROR_TYPE::CLASS_NODE_CONSTRUCT_ERROR;
	}
	if (file.empty()) //error handle
	{
		cerr << "node construct: none file found: " << code.get_loc() << endl;
		throw MY_ERROR_TYPE::CLASS_NODE_CONSTRUCT_ERROR;
	}

	//open file
	workbook wb;
	try
	{
		wb.load(file[0]);
	}
	catch (const xlnt::exception& exp) //error handle
	{
		cerr << exp.what() << endl;
		cerr << "node construct: cannot load workbook: " << file.at(0) << endl;
		throw MY_ERROR_TYPE::CLASS_NODE_CONSTRUCT_ERROR;
	}
	auto ws = wb.active_sheet();
	auto rg_merged = ws.merged_ranges();

	//find row index in first column
	range_reference rg_col(1, 2, 1, ws.highest_row_or_props());
	vector<cell_reference> row_idx = find_cell(ws, rg_col, to_string(code.get_row_idx()), MODE::FIND_CELL_ALL_MATCH);
	vector<cell_reference> row_idx_next = find_cell(ws, rg_col, to_string(code.get_row_idx() + 1), MODE::FIND_CELL_ALL_MATCH);
	if (row_idx.size() > 1) //error handle
	{
		cerr << "node construct: multi row_idx found: " << code.get_code() << endl;
		throw MY_ERROR_TYPE::CLASS_NODE_CONSTRUCT_ERROR;
	}
	if (row_idx.empty()) //error handle
	{
		cerr << "node construct: none row_idx found: " << code.get_code() << endl;
		throw MY_ERROR_TYPE::CLASS_NODE_CONSTRUCT_ERROR;
	}

	//find column index in selected row
	range_reference rg_row(row_idx[0].column_index() + 1, row_idx[0].row(), ws.highest_column_or_props(), row_idx[0].row());
	vector<cell_reference> col_idx = find_cell(ws, rg_row, to_string(code.get_col_idx()), MODE::FIND_CELL_ALL_MATCH);
	if (col_idx.size() > 1) //error handle
	{
		cerr << "node construct: multi column_idx found: " << code.get_code() << endl;
		throw MY_ERROR_TYPE::CLASS_NODE_CONSTRUCT_ERROR;
	}
	if (col_idx.empty()) //error handle
	{
		cerr << "node construct: none column_idx found: " << code.get_code() << endl;
		throw MY_ERROR_TYPE::CLASS_NODE_CONSTRUCT_ERROR;
	}

	//get node info
	row_t rw = col_idx[0].row() + 1;
	column_t cl = col_idx[0].column();
	string tmp = "";
	while (rw <= ws.highest_row() && (row_idx_next.empty() ? true : rw <= row_idx_next[0].row()))
	{
		tmp = utf2str(ws.cell(2, rw).to_string());

		if (tmp == "号")
		{
			string value = get_merged_cell_value(ws, rg_merged, cl, rw);
			for_each(value.begin(), value.end(), [](char& c) {
				if (c == '\n') { c = '+'; }});
			node::cur_name = node::cur_name + " " + value;
		}
		else if (tmp == "上")
		{
			string value = get_merged_cell_value(ws, rg_merged, cl, rw);
			if (value != "")
			{
				for_each(value.begin(), value.end(), [](char& c) {
					if (c == '\n') { c = '+'; }});
				node::up_name.push_back(value);
				node::up_code.push_back(node_code(value));
			}
		}
		else if (tmp == "下")
		{
			string value = get_merged_cell_value(ws, rg_merged, cl, rw);
			if (value != "")
			{
				for_each(value.begin(), value.end(), [](char& c) {
					if (c == '\n') { c = '+'; }});
				node::down_name.push_back(value);
				node::down_code.push_back(node_code(value));
			}
		}
		else if (tmp == "注")
		{
			string value = get_merged_cell_value(ws, rg_merged, cl, rw);
			for_each(value.begin(), value.end(), [](char& c) {
				if (c == '\n') { c = '+'; }});
			node::comment = node::comment + " " + value;
		}

		rw++;
	}
	node::workbook_name = file[0];
	node::worksheet_name = ws.title();
	node::cur_code = code;
}

bool node::has_up_code()
{
	return any_of(up_code.begin(), up_code.end(), [](auto& val) {
		return !val.is_empty();
		});
}

bool node::has_down_code()
{
	return any_of(down_code.begin(), down_code.end(), [](auto& val) {
		return !val.is_empty();
		});
}

void node::swap_up_down()
{
	swap(up_code, down_code);
	swap(up_name, down_name);
}

