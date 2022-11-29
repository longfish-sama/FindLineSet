#pragma once

#include <iostream>
#include <wchar.h>
#include <Windows.h>
#include <vector>
#include <fstream>
#include <io.h>
#include <string>
#include <xlnt/xlnt.hpp>

using namespace std;
using namespace xlnt;

struct MyStruct
{

};

void get_file(std::string path, std::vector<std::string>& files);

string utf2str(const std::string& str);

vector<cell_reference> find_str(worksheet& ws, string str, int mode);