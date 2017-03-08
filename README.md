# PySymTool
PySymTool is a Python command line tool which is used to symbolicate the iOS crash stack in batch

## Input
1. Excel file contains crash stack information that need to symbolicate;
2. Python script file;
3. symbolicatecrash tool;
4. Python modules(xlrd: read from excel, xlsxwriter: write to excel)

## Output
1. Excel file contains all symblicated stack information

## 注意事项

### 1.命名

需将 `dSYM` 文件和 `app` 文件放在同一目录，并且两个文件名需一致，举个例子：

* app 文件名为: tztHuaTaiZLMobile.app.1494.app
* dSYM 文件名则必须为: tztHuaTaiZLMobile.app.1494.dSYM 

### 2.symbolicatecrash 版本

如果 `app` 是由 `Xcode 7` 打出来的，则必须使用 `Xcode 7` 版本的 `symbolicatecrash` 来进行符号化。

使用如下命令可以查找 `symbolicatecrash` 所在路径：

```shell
find /Applications/Xcode.app -name symbolicatecrash -type f
``` 


