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

	copy_folder(source_folder, work_folder);

	get_folder_file(work_folder, filelist);

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


	vector<cell_reference> res = find_cell(ws, "西岭", MODE::FIND_CELL_PART_MATCH, MODE::FIND_CELL_RET_ALL_CELL);
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
		node_code nc(ws, c);
		cout << nc.get_code() << endl;
		node(nc, filelist);
		vector<list<node>>vln = get_node_lists(nc, filelist);
		cout << list2str(vln, MODE::LIST2STR_DETAIL) << endl;
	}
}

int main(int argc, char* argv[])
{
#ifdef TEST
	test();
	system("pause");
	cout << INFO_COUT << "in debug mode" << endl;
#endif

	vector<string> codes_to_find;
	vector<string> names_to_find;
	MODE find_cell_match_mode = MODE::FIND_CELL_PART_MATCH;
	MODE list2str_mode = MODE::LIST2STR_BRIEF;
	string source_data_folder;
	string work_data_folder = "data";
	bool require_refresh = false;

	auto cli = (
		clipp::required("-code") & clipp::values("codes to find", codes_to_find) |
		clipp::required("-name") & clipp::values("names to find", names_to_find),
		clipp::option("-out") &
		(clipp::required("brief").set(list2str_mode, MODE::LIST2STR_BRIEF) |
			clipp::required("detail").set(list2str_mode, MODE::LIST2STR_DETAIL)),
		clipp::option("-refresh").set(require_refresh),
		clipp::option("-source") & clipp::value("source_folder", source_data_folder)
		);
	if (!clipp::parse(argc, argv, cli))
	{
		cout << clipp::make_man_page(cli, argv[0]) << endl;
	}

	if (require_refresh)
	{
		copy_folder(source_data_folder, work_data_folder);
	}

	vector<string> files_list;
	get_folder_file(work_data_folder, files_list);

	for (auto code : codes_to_find)
	{
		if (is_node_code(code))
		{
			vector<list<node>> tmp_codes_result;
			node_code tmp_code(code);
			tmp_codes_result = get_node_lists(tmp_code, files_list);
			cout << tmp_code.get_code() << endl;
			cout << list2str(tmp_codes_result, list2str_mode) << endl;
		}
	}

	for (auto const& name : names_to_find)
	{
		for (auto const& file : files_list)
		{
			workbook wb;
			worksheet ws;
			vector<cell_reference> tmp_find_cell_results;
			try
			{
				wb.load(file);
				ws = wb.active_sheet();
			}
			catch (const xlnt::exception& ex)
			{
				cerr << ex.what() << "fail to load file: " << file << endl;
				continue;
			}
			tmp_find_cell_results = find_cell(ws, name, MODE::FIND_CELL_PART_MATCH, MODE::FIND_CELL_RET_ALL_CELL);
			for (auto val : tmp_find_cell_results)
			{
				node_code tmp_code(ws, val);
				if (!tmp_code.is_empty())
				{
					cout << tmp_code.get_code() << endl;
					vector<list<node>>tmp_name_result = get_node_lists(tmp_code, files_list);
					cout << list2str(tmp_name_result, list2str_mode) << endl;
				}
			}
		}
	}
}