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
	get_folder_file("data", filelist);

	//copy_folder("D:/chend/OneDrive/工/202211培训/04-线架", "data");

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


	vector<cell_reference> res = find_cell(ws, "西岭", MODE::FIND_CELL_PART_MATCH);
	for (size_t i = 0; i < res.size(); i++)
	{
		cout << utf2str(ws.cell(res[i]).to_string()) << endl;
	}
	cout << ws.cell("D36").to_string() << endl;

	node_code code("e418-r1-44");
	cout << code.get_code() << endl;
	if (code.is_empty())return;
	node n(code, filelist);

	list<node> l = get_node_list(code, filelist);
	cout << list2str(l) << endl;

}

int main()
{

	test();
	system("pause");
}