# 报告自动化工具介绍
## 概述
工具主要用于检测/检验报告的自动化生成，由报告的模板文件和原始记录表格文件两大部分组成。 <br>
检验报告模板的格式和内容完全分离。<br>
报告模板内置于自动化工具中，控制着报告的格式。<br>
原始记录采用Excel形式，控制内容。<br>
报告模板和原始记录模板后期可以分开维护和管理。<br>
原始记录测试完成后，报告一键自动化生成，基本不需要人工录入和修改。<br>

## 其他事项
支持命令行 和 GUI 两种方式运行， 控制开关是 reportAardio.py 中 AARDIO = False; <br>
使用aardio打包出windows exe可执行程序;<br>
图形化界面采用 aardio 制作，具体内容存放在 `Aardio_AutoReport`文件夹下。<br>
原始记录的Excel模板在 `Template_examples` 文件夹下。 <br>

## 联系人：
liugang@caict.ac.cn


