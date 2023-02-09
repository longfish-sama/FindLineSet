/**
 * @file   Untitled-1.cpp
 * @brief
 *
 * @date   2022-11-29
 */

#include "Header.hpp"
#include "clipp.h"
#include "operate_config.h"
#include <chrono>
#include <cstdlib>
#include <cstdio>
#include <iostream>
#include <string>

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
	MODE list2str_mode = MODE::LIST2STR_BRIEF;
	string work_data_folder = "data";
	const char config_filepath[] = "config";
	bool require_refresh = false;
	const time_t cur_update_time = chrono::system_clock::to_time_t(chrono::system_clock::now());
	//var define in file or by user
	string source_data_folder = "";
	string separate_1to3_tmp = "";
	vector<string> separate_1to3_list;
	time_t last_update_time = 0;
	unsigned int duration_hour = 6;

	//read config file
	try
	{
		ConfigHandle.init(config_filepath);
	}
	catch (const operatorconfig::File_not_found& ex)
	{
		cerr << ERROR_COUT << "config file not exist: " << ex.filename << endl;
		return -1;
	}

	//set parameter base on config file
	//lower priority
	try
	{
		source_data_folder = ConfigHandle.read<string>("source_data_folder");
		separate_1to3_tmp = ConfigHandle.read<string>("separate_1to3");
		last_update_time = ConfigHandle.read<time_t>("last_update_time");
		duration_hour = ConfigHandle.read<int>("duration");
	}
	catch (const operatorconfig::Key_not_found& ex)
	{
		cerr << WARNING_COUT << "key pair not exist: " << ex.key << endl;
	}
	//update time
	ConfigHandle.add("last_update_time", cur_update_time);
	ofstream out_config;
	out_config.open(config_filepath);
	out_config << ConfigHandle;
	out_config.close();
	//cal duration
	const time_t duration_sec = duration_hour * 60 * 60;
	if (cur_update_time - last_update_time >= duration_sec)
	{
		require_refresh = true;
	}
	//get separator list
	istringstream iss(separate_1to3_tmp);
	string tmp = "";
	while (getline(iss, tmp, '+'))
	{
		separate_1to3_list.push_back(tmp);
	}

	//paste parameter
	//higher priority
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
		return -1;
	}

	//process start
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