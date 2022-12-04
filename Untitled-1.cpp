/**
 * @file   Untitled-1.cpp
 * @brief
 *
 * @date   2022-11-29
 */

#include "Header.hpp"

using namespace std;
using namespace xlnt;

int main()
{
	workbook wb;
	wb.load("data/E413-36.xlsx");
	auto ws = wb.active_sheet();
	vector<cell_reference> res = find_cell(ws, "Î÷Áë", 0);
	for (size_t i = 0; i < res.size(); i++)
	{
		cout << utf2str(ws.cell(res[i]).to_string()) << endl;
	}
	vector<string> filelist;
	get_folder_file("data", filelist);
	cout << ws.cell("D36").to_string() << endl;

	node_code code("e418-y3-1");
	cout << code.get_code() << endl;
	node n(code, filelist);
	system("pause");
}