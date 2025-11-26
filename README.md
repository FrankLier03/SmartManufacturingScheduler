# SmartManufacturingScheduler
VBA
✅ 如何使用（Usage）
1. 准备 Excel 模板
本项目需要三个工作表（名称必须一致）：
📄 1）Orders
用于放置工单数据，示例结构：
OrderID	Qty	LineType	DemandDate
A001	120	RW	2025/11/18
A002	80	BTT	2025/11/25
...	...	...	...
📄 2）Config
用于存放基础参数及产能设置：
Key	Value
RW_日常产能	400
97/98/06	200
BTT常产能	50
预留产能	0
StartDate	2025/11/21
EndDate	2025/12/31
组装天数	3
包装天数	1
注意：第 7 行 Key 必须是：焊接产能
因为系统内部会用这个做合并产能控件。

焊接阶段甘特图
组装阶段甘特图
包装阶段甘特图
状态颜色图例
订单甘特汇总图
