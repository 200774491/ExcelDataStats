# ExcelDataStats
用于导出Excel表中数据做数据统计

当程序执行时会自动读取程序目录的config.ini 文件内容如果没有则需要自己输入路径信息
[config示例]
[Paths]
ExcelFile = D:\文档\数据.xlsm
SheetNames = 头饰,甲胄,靴子
OutputFile = D:\war3Project\文档\物品\output.txt
SelectedHeaders = 随机防御
WriteFirstNonEmptyOnly = False



* WriteFirstNonEmptyOnly 表示 遍历 SelectedHeaders 列下的全部数据,或只执行一次就跳过
