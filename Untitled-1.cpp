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

	copy_folder(source_folder, work_folder);

	workbook wb;
	worksheet ws;
	try
	{
		vector<string>f = find_filename(filelist, "E418-Q3");
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

	node_code code("E418-R1-44");
	cout << code.get_code() << endl;
	if (code.is_empty())return;
	node n(code, filelist);

	list<node> l = get_single_node_list(code, filelist);
	cout << list2str(l, MODE::LIST2STR_DETAIL) << endl;
	get_node_lists(code, filelist);
	for (auto c : res)
	{
		node_code nc(ws, c, filelist);
		cout << nc.get_code() << endl;
		node(nc, filelist);
		vector<list<node>>vln = get_node_lists(nc, filelist);
		cout << list2str(vln, MODE::LIST2STR_DETAIL) << endl;
	}
}

int main(int argc, char* argv[])
{

	test();
	system("pause");

	vector<string> codes_to_find;
	vector<string> names_to_find;
	MODE find_cell_match_mode = MODE::FIND_CELL_PART_MATCH;
	MODE find_cell_return_mode = MODE::FIND_CELL_RET_ALL_CELL;
	MODE list2str_mode = MODE::LIST2STR_BRIEF;
	string source_data_folder;
	string work_data_folder = "data";
	bool require_refresh = false;

	auto cli = (
		(clipp::required("-code")&clipp::values("codes to find",codes_to_find)|
		clipp::required("-name")&clipp::values("names to find",names_to_find)),
		clipp::option("-out")&
			(clipp::required("brief").set(list2str_mode,MODE::LIST2STR_BRIEF)|
			clipp::required("detail").set(list2str_mode,MODE::LIST2STR_DETAIL)),
		clipp::option("-refresh").set(require_refresh),
		clipp::option("-source")&clipp::value("source_folder",source_data_folder)
		);
	if (!clipp::parse(argc,argv,cli))
	{
		cout << clipp::make_man_page(cli, argv[0]) << endl;
	}
}