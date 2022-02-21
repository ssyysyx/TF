## 操作指南

#### 1. 运行环境

---

* windows系统
* Anaconda--安装时已勾选环境变量
* wind金融终端--已修复python接口且在登录状态



#### 2. 简介

---

* 债券持仓情况集中度
  * 把握**当前交易日**我方债券类的持仓情况，集合**交易所**和**银行间**所持债券，重点关注**房地产**行业持有债券，汇总**分行业**债券类持仓**集中度**，呈现**持仓总量变化**折线图
  * 一般运行时段：交易日**10:30-15:00** 任一时刻
* 定增投后管理
  * 截至**前一个交易日**，追踪定向增发投资产品的风险及损益，维护各项指标
  * 一般运行时段：交易日**14:30-15:00** 任一时刻
* 产品投后管理
  * 截至**前一个交易日**，追踪基金投资产品的风险及损益，维护各项指标
  * 一般运行时段：交易日**15:00-15:30** 任一时刻
* 浮动收益成交
  * 以**前一个交易日**浮动收益市值法为基础，融入**当日的**持仓地产债境内成交跟踪，扩增四列得到成交价的收益，计算隐形风险
  * 一般运行时段：交易日**15:10-15:40** 任一时刻



#### 3. 脚本执行方法

---


* 脚本文件是`execute.py`
* 键入`win`+`r` 进入运行
  ![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/1.png)
* 键入`cmd`+`enter` 进入终端
  ![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/2.png)
* 键入`cd `+拖入execute.py所在的目录+`enter` 进入运行目录
  ![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/3.png)
* 键入`python execute.py`+`enter`开始执行脚本，将自动输出任务日志
  ![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/4.png)
  ![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/5.png)
  ![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/6.png)
  ![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/7.png)
* 输出文件的路径展示在任务日志的最后面
  ![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/8.png)

#### 4.脚本的输入文件

---


![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/9.png)

* 输入文件在./data/下
  【`td`: 当前交易日 `td-1`: 前一个交易日 `td-2`: 前前个交易日】 
  
  * 债券持仓情况集中度：3个
    * `./result`中`td-1`的债券持仓情况集中度.xlsx
    * `O32-新综合信息查询-组合证券`中`td`的债券.xls
    * `Comstar`中`td`的收益风险评估_投资组合.xls
  * 定增投后管理：4个
    * `./result`中`td-2`的定增投后管理.xlsx
    * `mail财通基金君享天成单一资产管理计划`的`td-1`君享天成.xls
    * `mail国泰君安收益互换盯市日报`的`td-1`盯市日报表.xlsx
    * `mail收益互换日报表&交易回执`的`td-1`收益呼唤日报表.xls
  * 产品投后管理：6个
    * `./result`中`td-2`的产品投后管理.xlsx
    * `mail资产净值公告_SQS889_凡二量化对冲7号1期私募证券投资基金A`的`td-1`资产净值公告_SQS889_凡二量化对冲7号1期私募证券投资基金A.xls
    * `mail衍复天禄1000指增一号私募证券投资基金_五矿证券FOF11号单一资产管理计划_虚拟计提净值表`的`td-1`衍复天禄1000指增一号私募证券投资基金_五矿证券FOF11号单一资产管理计划_TA虚拟计提后净值表.xlsx
    * `mail财通基金君享丰硕定增量化对冲单一资产管理计划`的`td-1`君享丰硕定增量化对冲_Z.xls
    * `mail赫富尊享十九号私募证券投资基金_五矿证券FOF11号单一资产管理计划_虚拟计提净值表`的`td-1`赫富尊享十九号私募证券投资基金_五矿证券FOF11号单一资产管理计划_TA虚拟计提后净值表.xlsx
    * `mail衍复天禄灵活对冲三号私募证券投资基金_五矿证券FOF11号单一资产管理计划_虚拟计提净值表`的`td-1`衍复天禄灵活对冲三号私募证券投资基金_五矿证券FOF11号单一资产管理计划_TA虚拟计提后净值表.xlsx
  * 浮动收益成交：2个
    * `from张老师`的`td-1`浮动收益.xlsx
    * `from张老师`的`td`持仓地产债境内成交跟踪.xlsx

#### 5.提交注意事项

---

* 债券持仓情况集中度
  * 跑完进入输出文件债券类持仓sheet，如果行业列完整，不用更改；如果存在有行业列是空的，则到wind找到其行业，并复制整行，粘贴到输入文件的隐藏工作表`债券类持有列表`的最后一行，保存，进入终端重新跑脚本；
  * 发送**输出文件**和其`sheet占比`的**截图**到**跨境风险监控群**
  * 附：吕总，这是今天的债券类持仓监控
  ![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/10.png)
  ![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/11.png)
* 定增投后管理
  * 跑完进入输出文件的`sheet2129`，修改最后一行的`B列：单元资产净值`，数据源自`O32-新综合信息查询-单元资产-查询条件td-1-基金编号2129-取单元资产净值（净价）总和` ，保存。
  * 发送**输出文件**到**戚老师个人微信**、**跨境风险监控群**、**敞口定增跨境产品群**
  * 附：
    * 戚老师，这是`td-1`的定增投后管理
    * 吕总，这是`td-1`的定增投后管理
    * 各位老师，这是`td-1`的定增投后管理
* 产品投后管理
  * 跑完进入输出文件的`sheet天行者`，查询天行者1号2号3号的单位净值 [天行者单位净值查询](https://www.mintrust.com/informationDisclosure/netWorthDisclosure) ，如果有更新，则填入并更新对应日期及之后的行；无更新，则不必更改。【此操作一般每周只需进行一遍】
  * 发送**输出文件**到**跨境风险监控群、敞口定增跨境产品群**
  * 附：
    * 吕总，这是td-1的产品投后管理
    * 各位老师，这是td-1的产品投后管理
* 浮动收益成交
  * 跑完打开输出文件`sheet市值法`检查浮动净价收益和成交价的收益差距是否过大，L列数据是否合理，必要时进行修改并保存。
  * 将输出文件发给张老师，张老师检查回复ok后，再发至**跨境风险监控群**
  * 附：
    * 吕总，这是td-1的浮动成交收益表

![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/12.png)

#### 6.不定时改动文件

---

* 定增投后管理的持仓sheet改动较多
	* ![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/13.png)
  	* 改动**定增投后管理sheet持仓**为下面的状态
  	![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/14.png)
  * ![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/15.png)
		* 改动**定增投后管理sheet持仓**为下面的状态
		![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/16.png)
  * 不定期查wind：wind右下角输入ECM并回车，左下角搜定向增发，代码搜索股票存在空缺的，日期选近两个月，提取数据，**发行**日期**上市**日期**解禁**日期 
  ![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/17.png)
  * 财通-君享天成 国君互换-盯市 银河-收益互换：申购金额