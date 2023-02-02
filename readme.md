# 开发配置

0. VS2019 Windows10 C++17

1. 引用库下载并解压：[clipp](https://github.com/muellan/clipp/releases/tag/v1.2.3)，[xlnt](https://github.com/tfussell/xlnt/releases/tag/v1.5.0)

2. 修改 xlnt\include\xlnt\cell\cell.hpp、xlnt\source\cell\cell.cpp：[
xlnt彻底解决中文问题
](https://blog.csdn.net/pipi0714/article/details/104757233)

3. 修改 xlnt\include\xlnt\cell\rich_text.hpp，xlnt\source\cell\rich_text.cpp，xlnt\source\detail\serialization\xlsx_consumer.cpp：[Qt使用xlnt操作Excel（二）：导入Excel](https://blog.csdn.net/baidu_30570701/article/details/89437242)

4. 进入VS2019，目录树 Untitled-1 右击-属性 进入.![exp1](https://img-blog.csdnimg.cn/20181127221452959.png)选择 VC++目录-常规-包含目录 加入两个引用库的文件夹位置，VC++目录-常规-库目录 添加[xlntd.lib](xlntd.lib)所在的文件夹位置，链接器-输入-附加依赖项 添加[xlntd.lib](xlntd.lib)![exp2](https://img-blog.csdnimg.cn/20181127221935359.png)参考[Windows C++ 项目属性页参考](https://learn.microsoft.com/zh-cn/cpp/build/reference/property-pages-visual-cpp?view=msvc-160)

5. [xlntd.dll](xlntd.dll)应在[Untitled-1.cpp](Untitled-1.cpp)旁边