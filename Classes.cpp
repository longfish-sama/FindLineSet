/**
 * @file   Class.cpp
 * @brief  
 * 
 * @date   2022-12-20
 */

#include "Header.hpp"

/**
 * @brief 
 * 
 * @param code: 
 */
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

/**
 * @brief .
 * 
 * @param ws
 * @param cr
 * @param files_list
 */
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

		if (tmp == "╨е")
		{
			string value = get_merged_cell_value(ws, rg_merged, cl, rw);
			for_each(value.begin(), value.end(), [](char& c) {
				if (c == '\n') { c = '+'; }});
			node::cur_name = node::cur_name + " " + value;
		}
		else if (tmp == "ио")
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
		else if (tmp == "об")
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
		else if (tmp == "в╒")
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

