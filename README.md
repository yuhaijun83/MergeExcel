# MergeExcel
Merge multiple Excel files

# Python版本
Python 3.13.11

# 克隆完成后，进入到MergeExcel文件夹
cd ./MergeExcel
# 创建虚拟环境
python -m venv .venv
# 激活虚拟环境
.venv\Scripts\activate.bat
# 升级pip
python -m pip install --upgrade pip
# 安装依赖包
pip install -r requirements.txt

# 各文件说明
│ &nbsp;&nbsp; .gitignore <br>
│ &nbsp;&nbsp; Activate.bat                     &nbsp;&nbsp;&nbsp;&nbsp;-> # 快捷激活环境 <br>
│ &nbsp;&nbsp; AutoRun.bat                      &nbsp;&nbsp;&nbsp;&nbsp;-> # 自动启动合并工具（调用main.py，需要有python环境） <br>
│ &nbsp;&nbsp; Excel合并工具.exe                 &nbsp;&nbsp;&nbsp;&nbsp;-> # 打包好的可执行程序，无需安装python环境 <br>
│ &nbsp;&nbsp; LICENSE <br>
│ &nbsp;&nbsp; main.py                          &nbsp;&nbsp;&nbsp;&nbsp;-> # 合并工具源程序 <br>
│ &nbsp;&nbsp; MergeExcel.spec
│ &nbsp;&nbsp; PyInstaller.bat                  &nbsp;&nbsp;&nbsp;&nbsp;-> # py源码打包成用脚本，在无python环境下运行<br>
│ &nbsp;&nbsp; README.md <br>
│ &nbsp;&nbsp; requirements.txt                 &nbsp;&nbsp;&nbsp;&nbsp;-> # 依赖包文件 <br>
│ <br>
├─.idea                             &nbsp;&nbsp;&nbsp;&nbsp;-> # pycharm工程配置文件夹 <br>
│ <br>
├─.venv                             &nbsp;&nbsp;&nbsp;&nbsp;-> # 虚拟环境目录 <br>
│ <br>
├─dist
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;MergeExcel.exe               &nbsp;&nbsp;&nbsp;&nbsp;-> # PyInstaller.bat打包生成的可执行程序，在无python环境下运行<br>
│ <br>
├─input                             &nbsp;&nbsp;&nbsp;&nbsp;-> # Excel 输入目录（待合并Excel文件） <br>
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;input_a.xlsx <br>
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;input_b.xlsx <br>
│ <br>
└─output                            &nbsp;&nbsp;&nbsp;&nbsp;-> # Excel 输出目录（合并后Excel文件） <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Excel合并结果_20260123_105846.xlsx <br>
  <br>

# 使用方法1
把要合并的excel放到input文件夹下，运行 AutoRun.bat ，会自动在 output 文件夹下生成合并结果<br>

# 使用方法2
把 dist文件夹下的 MergeExcel.exe 拷贝出来，放到跟 input 文件夹同级文件夹下，运行 MergeExcel.exe ，会自动在 output 文件夹下生成合并结果<br>

# Excel文件合并后的结果例
input_a.xlxs<br>
a b c<br>
1 2 3<br>
<br>
input_b.xlxs<br>
a b d<br>
5 6 7<br>
<br>
Excel合并结果_20260123_105846.xlsx<br>
来源文件&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;a b c d<br>
input_a.xlsx 1 2 3<br>
input_b.xlsx 5 6 &nbsp;&nbsp;  7<br>

