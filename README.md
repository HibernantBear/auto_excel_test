# Auto Excel Test / 自动Excel试卷批改

## Introduction / 介绍

This small tool aims for helping teachers handle repeating test paper correction and sum these scores into new excel, but this tool now is still imperfect and only for objective test, any response or coding is appreciate.
___
这个小工具主要是为了减轻教师批改试卷的负担，但是目前只能用于客观题，欢迎star和反馈

## Install / 安装

### What you need / 你需要:
- [Python](https://www.python.org/)
- [xlrd](https://pypi.org/project/xlrd/), [xlwt](https://pypi.org/project/xlwt/)

### Prepare work / 准备工作
1. Install Python<br>
   ___
   安装Python
2. Install xlrd and xlwt library: `pip install xlrd xlwt`
   ___
   安装 xlrd and xlwt 库: `pip install xlrd xlwt`

### Use this tool / 使用
1. Clone this repository: `git clone https://github.com/HibernantBear/auto_excel_test.git`
   ___
   克隆这个仓库: `git clone https://github.com/HibernantBear/auto_excel_test.git`
2. Prepare your test paper, answer and template files
   ___
   准备你的试卷，答案和模板文件
3. Open a terminal and run: `python ./main.py`
   ___
   打开终端运行: `python ./main.py`
4. Check the output.xls file in your folder
   ___
   在文件夹中查看 output.xls 文件
5. Star this project if you like, this will be the encouragement for me to update this project(of course if you want donate, that will be very kind of you and I do appreciate. [DONATE ME]())
   ___
   如果你喜欢的话请star这个项目，这对于我的开发将会是极大的鼓励(当然如果您想捐赠的话，我会非常感谢您[DONATE ME]())

## Template / 模板

Template file is a excel file contains two sheets named as "input" and "output". Input sheet defined the template of test paper and anser. Output sheet defined the template of the output file's structure
___
模板文件里面有两个表格，"input"的表格定义了试卷和答案的模板，"output"定义了输出Excel的模板，下面是两个模板的解释

### Input / 输入模板
- K_[num]: refer to "KEY" keyword, the right cell of this is the text you wanna show in the output file.
   ___
   K_[num]: 这个是值的变量名，右边的单元格就是输出文件里面显示的文本<br>
   ___
   e.g. / 例如

    |   |    A    |    B    |
    |:-:|:-------:|:-------:|
    | 1 |   K_0   |   Name  |
    | 2 |   K_1   |   Grade |
- Q_[num]: refer to "QUESTION" keyword, the rigth cell of this is the score of this question.
   ___
   Q_[num]: 是问题的变量名，右边的单元格是该题目的分数<br>
   ___
   e.g. / 例如

    |   |    A    |    B    |
    |:-:|:-------:|:-------:|
    | 1 |   Q_0   |    3    |
    | 2 |   Q_1   |    2    |
- SCORE: the rigth cell of this is text of the score, for localization.
   ___
   SCORE: 右边单元格表示了分数在输出文件的标题是什么，用于自定义和本地化

### Output / 输出模板

Keywords are the same of the input template, you can change the order to what you like, the output file will arrange in the same order.
___
和输入模板的关键词一样，你可以自定义顺序和那些要显示在输出文件中

## Demo / 例子

For more detail of the template, answer and test paper files, please refer to the demo folder in the root of this project
___
参阅根目录下的demo文件夹
