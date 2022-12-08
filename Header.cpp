/**
 * @file   Header.cpp
 * @brief
 *
 * @date   2022-11-29
 */

#include "Header.hpp"

 /**
  * @brief .
  *
  * @param folder: folder name without '/'or'\'
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

void copy_folder(string source_folder, string dest_folder)
{
	if (_access(dest_folder.c_str(), 0) != 0 && _mkdir(dest_folder.c_str()) != 0)
	{
		return;
	}

	vector<string> source_files, dest_files;
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
		cout << INFO + item << endl;
#endif
	}

	for (size_t i = 0; i < dest_files.size(); i++)
	{
		if (!dest_files.at(i).empty())
		{
			ifstream infile(source_files[i], ios::in | ios::binary);
			if (!infile.is_open() || infile.bad() || infile.fail())
			{
				cout << WARNING + source_files[i] + " fail to open" << endl;
				infile.close();
				continue;
			}
			ofstream outfile(dest_files[i], ios::out | ios::binary);
			if (!outfile.is_open() || outfile.bad() || outfile.fail())
			{
				cout << WARNING + dest_files[i] + " fail to open" << endl;
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
 * @param mode
 * @return
 */
vector<cell_reference> find_cell(worksheet& ws, string str, MODE mode = FIND_CELL_PART_MATCH)
{
	vector<cell_reference> cr;
	string tmp = "";

	switch (mode)
	{

	case FIND_CELL_PART_MATCH:
		for (auto row : ws.rows(false))
		{
			for (auto cell : row)
			{
				tmp = utf2str(cell.to_string());
				if (tmp.find(str) != string::npos)
				{
					cr.push_back(cell.reference());
				}
			}
		}
		break;

	case FIND_CELL_ALL_MATCH:
		for (auto row : ws.rows(false))
		{
			for (auto cell : row)
			{
				tmp = utf2str(cell.to_string());
				if (tmp == str)
				{
					cr.push_back(cell.reference());
				}
			}
		}
		break;

	default:
		break;
	}

	return cr;
}

vector<cell_reference> find_cell(worksheet& ws, range_reference& rg, string str, MODE mode = FIND_CELL_PART_MATCH)
{
	vector<cell_reference> cr;
	string tmp;

	switch (mode)
	{
	case FIND_CELL_PART_MATCH:
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

	case FIND_CELL_ALL_MATCH:
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

vector<string> find_filename(vector<string>& filename_list, string str)
{
	vector<string> res;
	for (size_t i = 0; i < filename_list.size(); i++)
	{
		if (filename_list[i].find(str) != string::npos && filename_list[i].find("~$") == string::npos)
		{
			res.push_back(filename_list[i]);
		}
	}
	return res;
}

list<node> get_node_list(node_code& code, vector<string>& file_dir_list)
{
	if (code.is_empty())
	{
		return list<node>();
	}
	else
	{
		list<node> ret;
		node n0(code, file_dir_list);
		node n1 = n0;
		ret.push_back(n0);

		while (n0.has_up_code())
		{
			node n_up(n0.get_up_code(), file_dir_list);
			ret.push_front(n_up);
			if (n_up.get_up_code() == n0.get_cur_code())
			{
				n_up.swap_up_down();
			}
			n0 = n_up;
		}

		while (n1.has_down_code())
		{
			node n_down(n1.get_down_code(), file_dir_list);
			ret.push_back(n_down);
			if (n_down.get_down_code() == n1.get_cur_code())
			{
				n_down.swap_up_down();
			}
			n1 = n_down;
		}
		return ret;
	}
}

node_code::node_code(string code)
{
	size_t index_1 = code.find_first_of("-");
	size_t index_2 = code.find_last_of("-");
	if (index_1 < code.length() && index_1 < index_2 && index_2 < code.length())
	{
		loc = code.substr(0, index_2);
		number = code.substr(index_2 + 1);
		for (size_t i = 0; i < loc.length(); i++)
		{
			loc[i] = toupper(loc[i]);
		}
	}
	else
	{
		loc = "";
		number = "";
		cout << "node_code construct error" << endl;
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
	int tmp = atoi(number.c_str());
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
	int tmp = atoi(number.c_str());
	if (tmp % 10 == 0)
	{
		return 10;
	}
	else
	{
		return tmp % 10;
	}

}


node::node(node_code& code, vector<string>& file_dir_list)
{
	//use node code to find file path
	vector<string> file_dir = find_filename(file_dir_list, code.get_loc());
	if (file_dir.empty() || file_dir.size() > 1)
	{
		cout << "multi or none file find: " << code.get_loc() << endl;
		for (size_t i = 0; i < file_dir.size(); i++)
		{
			cout << file_dir[i] << endl;
		}
		throw CLASS_NODE_CONSTRUCT_ERROR;
	}
	//open file
	workbook wb;
	try
	{
		wb.load(file_dir[0]);//fixme: range error
	}
	catch (const xlnt::exception& exp)
	{
		cerr << exp.what() << endl;
	}
	auto ws = wb.active_sheet();
	auto rg_merged = ws.merged_ranges();

	//find row index in first column
	range_reference rg_col(1, 2, 1, ws.highest_row_or_props());
	vector<cell_reference> row_idx = find_cell(ws, rg_col, to_string(code.get_row_idx()), FIND_CELL_ALL_MATCH);
	vector<cell_reference> row_idx_next = find_cell(ws, rg_col, to_string(code.get_row_idx() + 1), FIND_CELL_ALL_MATCH);
	if (row_idx.empty() || row_idx.size() > 1)
	{
		cout << "multi or none cell find: " << code.get_row_idx() << endl;
	}

	//find column index in selected row
	range_reference rg_row(row_idx[0].column_index(), row_idx[0].row(), ws.highest_column_or_props(), row_idx[0].row());
	vector<cell_reference> col_idx = find_cell(ws, rg_row, to_string(code.get_col_idx()), FIND_CELL_ALL_MATCH);
	if (col_idx.empty() || col_idx.size() > 1)
	{
		cout << "multi or none cell find: " << code.get_col_idx() << endl;
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
			string value = node::get_cell_value(ws, rg_merged, cl, rw);
			for_each(value.begin(), value.end(), [](char& c) {
				if (c == '\n') { c = '+'; }});
			node::cur_name = value;
		}
		else if (tmp == "上")
		{
			string value = node::get_cell_value(ws, rg_merged, cl, rw);
			for_each(value.begin(), value.end(), [](char& c) {
				if (c == '\n') { c = '+'; }});
			node::up_name = value;
			node::up_code = node_code(node::up_name);
		}
		else if (tmp == "下")
		{
			string value = node::get_cell_value(ws, rg_merged, cl, rw);
			for_each(value.begin(), value.end(), [](char& c) {
				if (c == '\n') { c = '+'; }});
			node::down_name = value;
			node::down_code = node_code(node::down_name);
		}
		else if (tmp == "注")
		{
			string value = node::get_cell_value(ws, rg_merged, cl, rw);
			for_each(value.begin(), value.end(), [](char& c) {
				if (c == '\n') { c = '+'; }});
			node::comment = value;
		}

		rw++;
	}
	node::workbook_name = file_dir[0];
	node::worksheet_name = ws.title();
	node::cur_code = code;
}

void node::swap_up_down()
{
	swap(up_code, down_code);
	swap(up_name, down_name);
}

string node::get_cell_value(worksheet& ws, vector<range_reference>& rg_merged, column_t& col, row_t& row) const
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
