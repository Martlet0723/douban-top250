# DataFrame 学习笔记

## 一、DataFrame 是什么

DataFrame 是 pandas 的核心数据结构，本质是一个**带标签的二维表格**（类似 Excel 表或 SQL 表）。

```python
import pandas as pd

df = pd.DataFrame({
    "电影名称": ["肖申克的救赎", "霸王别姬", "阿甘正传"],
    "评分":     [9.7,           9.6,      9.5],
    "评价人数":  [3293412,       2234521,  2456789],
})
```

结果：

| 行索引 | 电影名称     | 评分 | 评价人数 |
|--------|-------------|------|---------|
| 0      | 肖申克的救赎 | 9.7  | 3293412 |
| 1      | 霸王别姬     | 9.6  | 2234521 |
| 2      | 阿甘正传     | 9.5  | 2456789 |

- **行索引**（`df.index`）：默认是 0, 1, 2... 的整数序列，可以自定义行名
- **列索引**（`df.columns`）：列名（"电影名称"、"评分" 等）

---

## 二、df[...] 的真正机制：类型分派，不是常规索引

### 2.1 核心原则

`df[...]` 是 pandas 重载后的 `__getitem__`，**根据传入 key 的类型自动分派**到不同逻辑：

| 你传了什么 | pandas 的理解 | 取的是 |
|-----------|-------------|--------|
| `df["评分"]`（字符串） | 单列名 | 这一**列** |
| `df[["评分", "片名"]]`（字符串 list） | 多列名 | 这几**列** |
| `df[bool_series]`（bool Series） | 布尔掩码 | 筛**行**，True 保留 |
| `df[0:5]`（切片） | 行号切片 | 这几**行**（较少使用） |

**关键点**：`df[行标签]` 想取一整行是行不通的，pandas 会优先把它当列名去匹配。

```python
df[0]          # ❌ 找名为 0 的列，不是第一行
df.loc[0]      # ✅ 按标签取第 0 行
df.iloc[0]     # ✅ 按位置取第 0 行
```

### 2.2 这不是 Python 原生语法，是 pandas 劫持了 []

Python 原生 list 中：

```python
a = [1, 2, 3]
a[True]   # TypeError，索引只能是整数
a >= 2    # TypeError，list 不支持 >=
```

但 pandas 重载了 `[]`（`__getitem__`）和 `>=`（`__ge__`），让它们支持向量操作。

---

## 三、布尔索引（筛选行）：两步走

```python
df = df[df["评分"] >= 9.0]
```

这条语句分两步执行：

**第 1 步**：`df["评分"] >= 9.0` —— 向量化比较

| 行号 | df["评分"] | df["评分"] >= 9.0 |
|------|-----------|-------------------|
| 0    | 9.7       | True              |
| 1    | 9.6       | True              |
| 2    | 8.7       | False             |
| 3    | 9.2       | True              |

返回的是一个**等长的 bool Series**，不是单个 True/False。这是 pandas 重载了 `>=` 运算符的结果——它自动对列里的每个元素逐一比较（向量化操作）。

**第 2 步**：`df[bool_series]` —— 布尔索引筛行

pandas 把 bool Series 和 df 按**行索引对齐**，True 的行保留，False 的行丢弃。索引对齐是这件事能成立的关键——bool Series 和 df 必须拥有相同的行索引。

---

## 四、向量化操作（Vectorized Operations）

### 4.1 含义

语法上看起来像标量运算，但实际上是对整列/整个 DataFrame 的**批量 C 级操作**，不写循环，效率高。

```python
# 这些都是向量化操作，对整列执行
df["评分"].mean()         # 求平均
df["评分"].max()          # 最大值
df.sort_values("排名")     # 按列排序
df.groupby("评分区间").size()  # 分组统计
df["年份"].astype(int)     # 类型转换
df["电影名称"].str.strip()  # 字符串处理
```

### 4.2 你的项目中的实例

```python
# 布尔索引筛选（douban_top250.py 中的概念，不是实际代码）：
df = df[df["评分"] >= 9.0]     # 保留评分 >= 9 的行

# 排序：
df = df.sort_values("排名", ignore_index=True)

# 字符串列批量清洗：
df[col] = df[col].astype(str).str.strip()

# 分桶统计：
df["评分区间"] = pd.cut(df["评分"], bins=bins, labels=labels)
stats = df.groupby("评分区间", observed=False).size()

# 缺失值检查：
df["排名"].isna().any()

# 聚合：
df["评分"].mean()   # 平均分
df["评分"].max()    # 最高分
```

---

## 五、行索引 vs 列索引

| 需求 | 写法 | 说明 |
|------|------|------|
| 取单列 | `df["评分"]` | 返回 Series |
| 取多列 | `df[["评分", "电影名称"]]` | 返回 DataFrame |
| 按行标签取行 | `df.loc["行标签"]` | 标签索引 |
| 按行位置取行 | `df.iloc[0]` | 位置索引（0-based） |
| 筛行（布尔） | `df[df["评分"] >= 9.0]` | 布尔索引 |
| 取单个值 | `df.loc[行标签, "列名"]` | `df.loc[3, "评分"]` |
| 取单个值（位置） | `df.iloc[行号, 列号]` | `df.iloc[2, 0]` |

**常见错误**：`df[行标签]` 不会取行，会优先匹配列名。取行请用 `df.loc` 或 `df.iloc`。

---

## 六、常用操作速查

```python
# 基本信息
df.shape          # (行数, 列数)
df.columns        # 列名列表
df.dtypes         # 每列的数据类型
df.head(10)       # 前 N 行预览
df.info()         # 完整信息摘要

# 统计
df["评分"].mean()       # 均值
df["评分"].max()        # 最大值（搭配 idxmax 取对应行）
df["评分"].min()        # 最小值（搭配 idxmin 取对应行）
df["评价人数"].sum()    # 求和

# 筛选
df[df["评分"] >= 9.0]                      # 单一条件
df[(df["评分"] >= 9.0) & (df["年份"] >= 2000)]  # 多条件用 &（不是 and）
df[df["电影名称"].str.contains("爱")]       # 字符串包含

# 排序
df.sort_values("评分", ascending=False)     # 按列排序
df.sort_values("排名", ignore_index=True)   # 排序后重置索引

# 导出
df.to_excel("output.xlsx", index=False)     # 导出 Excel（不写行号）
df.to_csv("output.csv", index=False)        # 导出 CSV
```

---

## 七、关键理解

1. **`df[...]` 不是常规 Python 索引**——pandas 根据输入类型自动分派到列存取、行筛选还是行切片
2. **`>=` 等比较符被重载**——对列操作不是返回单个 True/False，而是逐元素比较返回 bool Series
3. **布尔索引筛选行的本质是索引对齐**——bool Series 的行索引和 df 的行索引必须一致
4. **向量化 = 不用写循环**——整列操作在 C 层面完成，语法像标量运算但实际是批量操作
5. **取行用 `.loc` / `.iloc`，取列用 `["列名"]`**——不要用 `df[行标签]` 取行
