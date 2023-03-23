## 📣 简介

该项目审核团队环境控制分析的建模数据集，主要用于监控选定供应商Division 21，22，51中因为Foreign Particulate的原因出现的投诉和返工。

## 📝 数据读取

### 采购数据

遍历目标文件夹，concat每月采购数据。

### 投诉数据

读取投诉明细数据，Foreign Particulate投诉的判断方式为：

- 'If Manufacturing Complaint'] == 'Y'
- ['Short Text For Cause Code'] == 'MFG: Environmental Controls'

然后按'Vendor Name','Year','Month','Division'聚合

### 验货数据

只识别需要返工的Foreign Particulate。是否为返工，在之前的数据集已经进行了判断。

- ['Reject Code'] == 'Foreign Particulate'

然后按'Vendor Name','Year','Month','Division'聚合

## 🔰数据建模

由于Tableau的数据连接特点，有并集（Union），连接（Join），混合（Blend）和关系（Relationship）。

将使用数据关系进行建模分析，具体的优点这里不赘述。由于数据关系的特点，度量值会完成保留，即有度量参与的表一定是Outer。

但是，3个数据源需要full join才能保证维度不为null。

通过多级索引创建笛卡尔积，子列分别为：

- vendor_list
- Year
- Month
- Division

合并数据源，并去除度量求和为0的行。
