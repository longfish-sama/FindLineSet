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

/**
 * @brief .
 *
 * @param str
 * @return
 */
string utf2str(const string& str)
{
	int nwLen = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), -1, NULL, 0);

	wchar_t* pwBuf = new wchar_t[nwLen + 1];
	memset(pwBuf, 0, nwLen * 2 + 2);

	MultiByteToWideChar(CP_UTF8, 0, str.c_str(), str.length(), pwBuf, nwLen);

	int nLen = WideCharToMultiByte(CP_ACP, 0, pwBuf, -1, NULL, NULL, NULL, NULL);

	char* pBuf = new char[nLen + 1];
	memset(pBuf, 0, nLen + 1);

	WideCharToMultiByte(CP_ACP, 0, pwBuf, nwLen, pBuf, nLen, NULL, NULL);

	string retStr = pBuf;

	delete[] pBuf;
	delete[] pwBuf;

	pBuf = NULL;
	pwBuf = NULL;

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
vector<cell_reference> find_cell(worksheet& ws, string str, int mode = 0)
{
	vector<cell_reference> c;
	string temp_str = "";

	for (auto row : ws.rows(false))
	{
		for (auto cell : row)
		{
			temp_str = utf2str(cell.to_string());
			if (temp_str.find(str) != string::npos)
			{
				c.push_back(cell.reference());
			}
		}
	}
	return c;
}

vector<cell_reference> find_cell(worksheet& ws, range_reference& range, string str, int mode = 0)
{
	return vector<cell_reference>();
}

vector<string> find_filename(vector<string>& filename_list, string str)
{
	vector<string> res;
	for (size_t i = 0; i < filename_list.size(); i++)
	{
		if (filename_list[i].find(str) != string::npos)
		{
			res.push_back(filename_list[i]);
		}
	}
	return res;
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

node_code::~node_code()
{
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

string node_code::get_code()
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

node::node(node_code code, vector<string>& file_list)
{
	vector<string> file_dir = find_filename(file_list, code.get_loc());
	if (file_dir.empty() || file_dir.size() > 1)
	{
		cout << "multi or none file find: " << code.get_loc() << endl;
		for (size_t i = 0; i < file_dir.size(); i++)
		{
			cout << file_dir[i] << endl;
		}
	}
	workbook wb;
	wb.load(file_dir[0]);
	worksheet ws = wb.active_sheet();

	range_reference rg_col(1, 1, 1, ws.highest_row_or_props());
	vector<cell_reference> cr_col = find_cell(ws, rg_col, to_string(code.get_tens_digit() + 1));
	if (cr_col.empty() || cr_col.size() > 1)
	{
		cout << "multi or none cell find: " << code.get_tens_digit() + 1 << endl;
	}

	range_reference rg_row(cr_col[0].column_index(), cr_col[0].row(), ws.highest_column_or_props(), cr_col[0].row());
	vector<cell_reference> cr_row = find_cell(ws, rg_row, to_string(code.get_units_digits()));
	if (cr_row.empty() || cr_row.size() > 1)
	{
		cout << "multi or none cell find: " << code.get_tens_digit() + 1 << endl;
	}
	style s = ws.cell(cr_row[0]).style();
}
