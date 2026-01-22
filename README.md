# MergeExcel
Merge multiple Excel files

# Python版本
Python 3.13.11

# 创建虚拟环境
python -m venv .venv
# 升级pip
python -m pip install --upgrade pip
# 安装依赖包
pip install -r requirements.txt

# 各文件说明
│  .gitignore <br>
│  Activate.bat                     &nbsp;&nbsp;&nbsp;&nbsp;-> # 快捷激活环境 <br>
│  AutoRun.bat                      &nbsp;&nbsp;&nbsp;&nbsp;-> # 自动启动合并工具（调用main.py，需要有python环境） <br>
│  Excel合并工具.exe                 &nbsp;&nbsp;&nbsp;&nbsp;-> # 打包好的可执行程序，无需安装python环境 <br>
│  LICENSE <br>
│  main.py                          &nbsp;&nbsp;&nbsp;&nbsp;-> # 合并工具源程序 <br>
│  PyInstaller.bat                  &nbsp;&nbsp;&nbsp;&nbsp;-> # py源码打包成可执行程序，在无python环境下运行<br>
│  README.md <br>
│  requirements.txt                 &nbsp;&nbsp;&nbsp;&nbsp;-> # 依赖包文件 <br>
│ <br>
├─.idea                             &nbsp;&nbsp;&nbsp;&nbsp;-> # pycharm工程配置文件夹 <br>
│  │  misc.xml <br>
│  │  modules.xml <br>
│  │  Process_Excel.iml <br>
│  │  vcs.xml <br>
│  │  workspace.xml <br>
│  │ <br>
│  └─inspectionProfiles <br>
│          profiles_settings.xml <br>
│          Project_Default.xml <br>
│ <br>
├─.venv                             &nbsp;&nbsp;&nbsp;&nbsp;-> # 虚拟环境目录 <br>
├─input                             &nbsp;&nbsp;&nbsp;&nbsp;-> # Excel 输入目录（待合并Excel文件） <br>
└─output                            &nbsp;&nbsp;&nbsp;&nbsp;-> # Excel 输出目录（合并后Excel文件） <br>
 <br>

# Excel文件合并例
input_a.xlxs<br>
a b c<br>
1 2 3<br>
<br>
input_b.xlxs<br>
a b d<br>
5 6 7<br>
<br>
output.xlxs<br>
a b c d<br>
1 2 3<br>
5 6 &nbsp;&nbsp;  7<br>

