#include <xlnt/xlnt.hpp>
#include "Header.hpp"

using namespace std;
using namespace xlnt;

int main()
{
	workbook wb;
	wb.load("data/E413-36.xlsx");
	auto ws = wb.active_sheet();
	vector<cell_reference> res = find_str(ws, "Î÷Áë", 0);
	for (size_t i = 0; i < res.size(); i++)
	{
		cout << utf2str(ws.cell(res[i]).to_string()) << endl;
	}
	//system("pause");
}