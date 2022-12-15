/**
 * @file   Untitled-1.cpp
 * @brief
 *
 * @date   2022-11-29
 */

#include "Header.hpp"

using namespace std;
using namespace xlnt;

void test()
{
	vector<string> filelist;
	string work_folder = "data";
	string source_folder = "D:/chend/OneDrive/工/202211培训/04-线架"; //todo: change here
	get_folder_file(work_folder, filelist);

	//copy_folder(source_folder, work_folder);

	workbook wb;
	worksheet ws;
	try
	{
		vector<string>f = find_filename(filelist, "E418-R1");
		cout << f.at(0) << endl;
		wb.load(f.at(0));
		ws = wb.active_sheet();
	}
	catch (const xlnt::exception& exp)
	{
		cerr << exp.what() << endl;
		exit(-1);
	}


	vector<cell_reference> res = find_cell(ws, "西岭", MODE::FIND_CELL_PART_MATCH,MODE::FIND_CELL_RET_ALL_CELL);
	for (size_t i = 0; i < res.size(); i++)
	{
		cout << utf2str(ws.cell(res[i]).to_string()) << endl;
	}
	cout << ws.cell("D36").to_string() << endl;

	node_code code("e418-r1-44");
	cout << code.get_code() << endl;
	if (code.is_empty())return;
	node n(code, filelist);

	list<node> l = get_single_node_list(code, filelist);
	cout << list2str(l, MODE::LIST2STR_DETAIL) << endl;
	get_node_lists(code, filelist);

}

int main(int argc, char* argv[])
{

	test();
	system("pause");
}