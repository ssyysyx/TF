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
  * 一般运行时段：**10:30-15:00** 任一时刻
* 定增投后管理
  * 截至**前一个交易日**，追踪定向增发投资产品的风险及损益，维护各项指标
  * 一般运行时段：**14:30-15:00** 任一时刻
* 产品投后管理
  * 截至**前一个交易日**，追踪基金投资产品的风险及损益，维护各项指标
  * 一般运行时段：**15:00-15:30** 任一时刻
* 浮动收益成交
  * 以**前一个交易日**浮动收益市值法为基础，融入**当日的**持仓地产债境内成交跟踪，扩增四列得到成交价的收益，计算隐形风险
  * 一般运行时段：**15:10-15:40** 任一时刻



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
* 输入文件在./data/下
  * 债券持仓情况集中度：3个
    * `./result`中前一个交易日的债券持仓情况集中度.xlsx
    * `O32-新综合信息查询-组合证券`中当前交易日的债券.xls
    * `Comstar`中当前交易日的收益风险评估_投资组合.xls
  * 定增投后管理：4个
    * `./result`中前前交易日的定增投后管理.xlsx
    * `mail财通基金君享天成单一资产管理计划`的君享天成.xls
    * `mail国泰君安收益互换盯市日报`的盯市日报表.xlsx
    * `mail收益互换日报表&交易回执`的收益呼唤日报表.xls
  * 产品投后管理：6个
    * `./result`中前前交易日的产品投后管理.xlsx
    * `mail资产净值公告_SQS889_凡二量化对冲7号1期私募证券投资基金A`的资产净值公告_SQS889_凡二量化对冲7号1期私募证券投资基金A.xls
    * `mail衍复天禄1000指增一号私募证券投资基金_五矿证券FOF11号单一资产管理计划_虚拟计提净值表`的衍复天禄1000指增一号私募证券投资基金_五矿证券FOF11号单一资产管理计划_TA虚拟计提后净值表.xlsx
    * `mail财通基金君享丰硕定增量化对冲单一资产管理计划`的君享丰硕定增量化对冲_Z.xls
    * `mail赫富尊享十九号私募证券投资基金_五矿证券FOF11号单一资产管理计划_虚拟计提净值表`的赫富尊享十九号私募证券投资基金_五矿证券FOF11号单一资产管理计划_TA虚拟计提后净值表.xlsx
    * `mail衍复天禄灵活对冲三号私募证券投资基金_五矿证券FOF11号单一资产管理计划_虚拟计提净值表`的衍复天禄灵活对冲三号私募证券投资基金_五矿证券FOF11号单一资产管理计划_TA虚拟计提后净值表.xlsx
  * 浮动收益成交：2个
    * `from张老师`的浮动收益.xlsx
    * `from张老师`的持仓地产债境内成交跟踪.xlsx

![win+r](https://gitee.com/oushisyx318/TF/raw/master/readme_files/8.png)
