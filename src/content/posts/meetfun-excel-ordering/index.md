---
title: Meetfun订货excel表使用手册
published: 2026-10-18
pinned: false
description: 阐述订货表格的组成、原理、每周工作流程
tags: [User manual]
category: others
licenseName: CC BY-SA 4.0
author: qvper
sourceLink: https://github.com/qvper/Mizuki-Blog
draft: false
date: 2026-10-16
image: './images/cover.webp'
pubDate: 2026-10-18
permalink: meetfun-excel-ordering
---

## 一、表格的组成
>表格由4个工作表+4个按钮+2组VBA宏代码组成

### 工作表

|  工作表名称     | 作用                 |
| -------- | ------------------ |
| **产品目录** | 存储所有产品的信息       |
| **盘点表**  | 按分类列出产品，填写实际库存 |
| **订货表**  | 汇总库存与销量，自动算出订货量    |
| **订货历史** | 存档每周的期末库存          |

### 按钮

| 按钮名称     | 按钮位置  | 作用                       |
| ------ | --- | ------------------------ |
| 生成盘点表  | 盘点表 | 重新按产品目录生成盘点表             |
| 生成订货表  | 订货表 | 重新生成订货表（**会清空期初/进货/报废**） |
| 读取上周数据 | 订货表 | 把历史里的期末库存填到期初库存          |
| 保存本周数据 | 订货表 | 把本周期末库存存入订货历史            |

### 宏
:::caution
打开excel表格时请务必允许宏运行（或点击确认 “启用内容”）
否则会导致自动化无法工作
:::
- 打开文件时会自动更新盘点表与产品目录一致
- 若打开文件时Excel弹出「启用内容」提示，**一定要点「启用内容」**，否则会导致自动化无法工作

## 二、表格的结构

### 产品目录
>产品目录存储了重要的产品信息，并以颜色区分不同类型的产品
- 第 1 行：标题
- 第 2 行：列标题（不要删、不要改顺序）
- 第 3 行起：产品数据，一个产品一行

| 列   | 内容    | 说明                  |
| --- | ----- | ------------------- |
| B   | 品号    | 产品的唯一编号（如 SC-A-001） |
| C   | 分类    | 产品的大致类型 |
| D   | 产品名称  |                     |
| E   | 单件容量  |                     |
| F   | 容量单位  |                     |
| G   | 装箱数   | 一箱（整体单位）装多少件，**必须 ≥1**    |
| H   | 盘点单位  | 盒/瓶/袋/罐/包/个         |
| I   | 保质期数值 |                     |
| J   | 保质期单位 |                     |
| K   | 储存条件  |                     |

截至目前，产品共有7个大类：

`乳制品`、`茶叶咖啡`、`果汁果酱`、`小料`、`装饰辅料`、`糖浆`、`清洁用品`

:::tip
当添加了新的产品时，表格会基于产品大类自动渲染其对应的颜色
:::

### 盘点表
- 按分类分成不同个区块，便于统一盘点一个类型的产品
- 列：B=品号、C=产品名称、D=盘点单位、E=当前库存
- 只有 **E 列（当前库存）** 需要填，其他列会自动生成，不要随意改动

### 订货表

营业额区域
| 单元格 | 内容      | 说明                    |
| --- | ------- | --------------------- |
| B1  | 使用周期营业额 | **过去**一个周期的实际营业额（如上周） |
| D1  | 订货周期营业额 | **未来**一个周期的预计营业额（如下周） |
| F1  | 业绩系数    |基于前两个数据自动计算，若两个都没填会按 1 算  |

订货表结构

| 列   | 内容    | 谁来填           | 公式                           |
| --- | ----- | ------------- | ---------------------------- |
| A   | 品号    | 自动            |                              |
| B   | 产品名称  | 自动            |                              |
| C   | 装箱数   | 自动            |                              |
| D   | 盘点单位  | 自动            |                              |
| E   | 期初库存  | 自动（读取上周）或手动   |                              |
| F   | 进货量   | **手动填**       |                              |
| G   | 报废    | **手动填**       |                              |
| H   | 期末库存  | **自动**（取自盘点表） | 按品号查找盘点表                     |
| I   | 本周使用量 | 自动            | 期初 + 进货 − 报废 − 期末            |
| J   | 业绩系数  | 自动            | = F1                         |
| K   | 需求量   | 自动            | 使用量 × 业绩系数                   |
| L   | 订货量   | 自动            | （需求量 − 期末）÷ 装箱数，向上取整；库存够时为 0 |
| M   | 订货单位  | 自动            | 装箱数 > 1 显示"箱"，否则只显示盘点单位       |

### 订货历史
>由「保存本周数据」自动追加；同一天同一品号只保留最新一条（自动去重）。


## 三、表格的原理

### 盘点表的由来

- 打开总文件时会自动执行宏，按产品目录的 7 个分类重建盘点表
- **行数 = 产品数量**，产品增减后重新打开文件即可**自动**适配
- **按品号保存库存**：店员填好的数字保留在对应品号上，产品变化不会错位
:::tip 
E 列已经添加了限制：**只允许 ≥0 的数字类型**
:::

### 订货表的计算逻辑
1. **期末库存（H）**：自动从盘点表按品号查找「当前库存」
2. **本周使用量（I）** = 期初 + 进货 − 报废 − 期末
3. **需求量（K）** = 本周使用量 × 业绩系数（F1）
4. **订货量（L）** = （需求量 − 期末库存）÷ 装箱数，向上取整
   - 若库存 ≥ 需求量 → 订货量为 **0**（不会显示负数）
   - 若盘点没填（期末为空）→ 订货量为 **0**
5. **订货单位（M）**：会自动根据装箱数换算订货单位

### 数据保护
- 盘点表「当前库存」：只允许 ≥0 数字类型
- 订货表「营业额」（B1/D1）、「期初库存/进货量/报废」（E/F/G）：只允许 ≥0 的数字类型
- 订货表「期末库存」（H）：**禁止手动填写**，若想更改数据请改盘点表

### 业绩系数

- 若不填 B1/D1，系数会锁定在 1 防止误操作
- 填了 B1 和 D1 → 系数 = D1/B1（例：上周营业额 10000，预计下周 12000 → 系数 1.2，需求放大 20%）
- **只填 1 个数据不生效**，系数会强制锁定在 1 防止误操作

:::note
没有在订货表内添加安全订货系数，例如安全系数为 1 ,则会确保门店永远会保留一个订货周期的库存量。

但是会与业绩系数重合

如果想多定货就把下一个周期的预计营业额写高一点
:::

## 四、每周的工作流程

### 盘点库存

1. 打开表格，点「启用内容」（允许宏运行）
2. 切到「盘点表」
3. 在 **当前库存** 逐行填写实际数量
4. 保存、关闭

或：
1. 单独导出盘点表交给伙伴填写
2. 在 **当前库存** 逐行填写实际数量
3. 收回盘点表后使用vlookup导入数据到主表

> 盘点表打开时会自动生成，


### 订货流程
1. 打开表格（启用宏）
2. 在「订货表」顶部：
   - 填写 B1：上一个周期的实际营业额
   - 填写 D1：下个周期的预计营业额
3. 点 **「读取上周数据」** → 期初库存的数据会自动拉取录入
4. 手动填写本周期的 **「进货量」** 和 **「报废」** 两列
5. 「期末库存」「本周使用量」「需求量」「订货量」全部自动计算
6. 按「订货量」和「订货单位」列下单订货
7. 点 **「保存本周数据」** → 将期末库存存入订货历史，以便下周使用。


### 下周循环
- 下周点「读取上周数据」时，本周保存的期末库存就是下周的期初库存
- 伙伴再盘点 → 重复订货步骤

## 五、注意事项
1. **切记要启用宏**：打开文件点「启用内容」。若不启用宏，自动化会无法使用，按钮也会失效
2. 所有产品信息只改产品目录。盘点表、订货表的数据来源均来自产品目录
3. **分类名需要完全一致**：`乳制品` 就是 `乳制品`，不能写成"乳制品 "、"奶制品"等。写错的产品会在盘点表/订货表直接消失。（填写正确的话会自动染色的）
4. 订货表的「期末库存」列的自动填写数据来源是盘点表，已经做了放呆设计，手动填写会被拦截，要改就改盘点表
5. **「生成订货表」按钮会清空期初/进货/报废**：点之前想清楚（有确认弹窗），点了之后要重新「读取上周数据」
6. **新增产品后**：盘点表打开自动更新；**订货表需要手动点「生成订货表」或重新打开文件**才会出现新产品。
7. 订货表期末库存为空、订货量 0 —— 这是提醒你还没盘点，不是错误
8. **保存数据前检查**：确认盘点表已填、订货表计算正常，再点「保存本周数据」，因为下周的期初库存来自这次保存

## 常见问题 FAQ

**Q：打开文件弹警告信息「是否允许启用内容」，要不要点？**

A：必须点「启用内容」，这是宏安全提示，不点的话自动同步和宏按钮都不工作

**Q：盘点表打开后是空的/不全？**

A：确认产品目录里产品的「分类」是否正确，改完后重新打开文件或点击按钮重新生成

**Q：订货量显示 0？**

A：检查：①盘点表是否填了该产品的当前库存（没填=0）；②业绩系数（B1/D1 是否至少一个为 1 的状态）；③需求是否小于等于库存

**Q：订货量显示负数？**

A：函数公式已经杜绝了负数产生，不可能出现此情况。库存充足时订货量会显示为 0

**Q：新产品订货表里没有？**

A：点「生成订货表」重新生成（注意会清空期初/进货/报废，生成后先「读取上周数据」）

**Q：历史表越来越长怎么办？**

A：目前自动保存不设置清理，因为即使每周上两个新品，也不会影响表格使用，如需只保留最近 N 周可以请联系qvper加清理功能

**Q：换电脑/文件名改了会怎样？**

A：按钮绑定的是宏名，改名不影响宏；但宏引用的是工作表名（盘点表/订货表等），**不要改工作表（Sheets）名称**

## VBA 代码详解

### 一、代码总览

文件里的 VBA 代码分布在 3 个位置

| 位置（VBA 工程中的名字） | 内容 | 行数 |
|---|---|---|
| **模块1** | 订货表相关：生成订货表、RegenOrder、读取上周数据、保存本周数据 | 262 行 |
| **盘点表同步** | 盘点表相关：生成盘点表、SyncInventory | 126 行 |
| **ThisWorkbook** | Workbook_Open 事件（打开文件时自动同步盘点表） | 6 行 |

#### 宏清单

| 宏名 | 所在模块 | 触发方式 | 作用 |
|---|---|---|---|
| `生成订货表` | 模块1 | 程序调用（RegenOrder 转发） | 按产品目录重建订货表 |
| `RegenOrder` | 模块1 | 订货表「生成订货表」按钮 | 弹确认框 → 调生成订货表 |
| `读取上周数据` | 模块1 | 订货表「读取上周数据」按钮 | 历史期末库存 → 期初库存 |
| `保存本周数据` | 模块1 | 订货表「保存本周数据」按钮 | 期末库存 → 订货历史 |
| `生成盘点表` | 盘点表同步 | 程序调用（SyncInventory / 打开文件） | 按产品目录重建盘点表 |
| `SyncInventory` | 盘点表同步 | 盘点表按钮 | 弹窗版生成盘点表（英文别名） |
| `Workbook_Open` | ThisWorkbook | 打开文件时自动 | 静默同步盘点表 |

#### 宏的调用关系

```
打开文件
  └─ Workbook_Open ──→ 生成盘点表(False)  [静默，不弹窗]

点盘点表按钮
  └─ SyncInventory ──→ 生成盘点表(True)    [弹「盘点表已生成！」]

点订货表「生成订货表」按钮
  └─ RegenOrder（确认框）
        └─ 点「是」──→ 生成订货表           [弹「订货表已生成！」]

点订货表「读取上周数据」/「保存本周数据」
  └─ 直接执行对应 Sub
```

### 二、宏代码拆解

#### 1. `生成订货表(Optional ByVal showMsg As Boolean = True)`

**作用**：清空订货表，按「产品目录」的 7 个分类重新生成整张订货表（品号/名称/装箱数/单位 + 8 个公式列 + 数据验证）。

**参数 `showMsg`**：
- `True`（默认）：完成后弹「订货表已生成！」
- `False`：静默（自动化测试用）
- **执行流程**：

```vba
Set wsOrder = Sheets("订货表")       ' 目标工作表
Set wsCat = Sheets("产品目录")        ' 数据来源
wsOrder.Cells.Clear                   ' ① 清空整张订货表（含已填数据！）
```

##### **① 第 1 行：业绩系数区域**

```vba
wsOrder.Cells(1, 1).Value = "使用周期营业额"
wsOrder.Cells(1, 2).Value = 0
wsOrder.Cells(1, 3).Value = "订货周期营业额"
wsOrder.Cells(1, 4).Value = 0
wsOrder.Cells(1, 5).Value = "业绩系数"
wsOrder.Cells(1, 6).Formula = "=IF(OR(B1=0,D1=0),1,D1/B1)"
```

- A1~E1 写文字标签，B1/D1 初始值为 0（营业额，用户填）
- **F1 是业绩系数**：`=IF(OR(B1=0,D1=0),1,D1/B1)`
  - 如果 B1 或 D1 任一为 0（没填）→ 系数 = 1（不调整）
  - 两个都填了 → 系数 = D1/B1（未来预计 ÷ 过去实际）

**② 营业额数据验证**（B1、D1 只允许 ≥0 数字）

```vba
With wsOrder.Range("B1")
    .Validation.Delete                    ' 先删旧验证，避免重复添加报错
    .Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
        Operator:=xlGreaterEqual, Formula1:="0"
    .Validation.IgnoreBlank = True        ' 允许留空
End With
```

**③ 分类块循环**（7 个分类）

```vba
categories = Array("乳制品", "茶叶咖啡", "果汁果酱", "小料", "装饰辅料", "糖浆", "清洁用品")
currentRow = 3
For i = LBound(categories) To UBound(categories)
```

每个分类做三件事：

**(a) 分类标题行**（合并 A:M，浅蓝底加粗）

```vba
wsOrder.Cells(currentRow, 1).Value = categories(i)
wsOrder.Range(wsOrder.Cells(currentRow, 1), wsOrder.Cells(currentRow, 13)).Merge
```

**(b) 列标题行**（品号/产品名称/…/订货单位，13 列）

```vba
wsOrder.Cells(currentRow, 1).Value = "品号"
... 共 13 个标题 ...
```

**(c) 产品行**（遍历产品目录，分类匹配才写入）

```vba
lastRow = wsCat.Cells(wsCat.Rows.count, 2).End(xlUp).Row   ' 产品目录最后一行
For r = 3 To lastRow
    If wsCat.Cells(r, 3).Value = categories(i) Then       ' C列=分类 匹配
        ' ① 写静态信息：品号(B列)、名称(D列)、装箱数(G列)、盘点单位(H列)
        wsOrder.Cells(currentRow, 1).Value = wsCat.Cells(r, 2).Value
        ...
        ' ② 写公式列（见下）
```
**产品目录列 → 订货表列对应关系**：

| 产品目录 | 订货表 |
|---|---|
| B 列 品号 | A 列 |
| D 列 产品名称 | B 列 |
| G 列 装箱数 | C 列 |
| H 列 盘点单位 | D 列 |

**④ 核心公式（每个产品行都写 6 个公式）**

```vba
' 期末库存 H：按品号从盘点表取「当前库存」
fH = "=IFERROR(IF(ISBLANK(INDEX(盘点表!$E$1:$E$2000,MATCH(A{r},盘点表!$B$1:$B$2000,0))),"""",INDEX(...)),"""")"
' 本周使用量 I = 期初 + 进货 - 报废 - 期末（期末空时显示 0）
wsOrder.Cells(currentRow, 9).Formula = "=IF(H{r}="""",0,E{r}+F{r}-G{r}-H{r})"
' 业绩系数 J = 引用 F1
wsOrder.Cells(currentRow, 10).Formula = "=$F$1"
' 需求量 K = 使用量 × 系数
wsOrder.Cells(currentRow, 11).Formula = "=I{r}*J{r}"
' 订货量 L = 需求 - 期末，按装箱数向上取整（库存够时 0）
wsOrder.Cells(currentRow, 12).Formula = "=IF(H{r}="""",0,IF(K{r}<=H{r},0,ROUNDUP((K{r}-H{r})/C{r},0)))"
' 订货单位 M = 装箱数>1 显示"箱"，否则显示盘点单位
wsOrder.Cells(currentRow, 13).Formula = "=IF(C{r}>1,""箱"",D{r})"
```

> 代码里 `q = Chr(34)` 是为了在 VBA 字符串里拼出双引号，`{r}` 会被替换成实际行号。

**⑤ 数据验证**（每个产品行）

- E/F/G 列（期初/进货/报废）：只能填 ≥0 数字
- H 列（期末库存）：**自定义验证 = FALSE**，即**任何手动输入都会被拒绝**

```vba
With wsOrder.Cells(currentRow, 8)
    .Validation.Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Formula1:="FALSE"
    .Validation.ErrorMessage = "期末库存自动取自盘点表，请勿手动填写；如需修改请改盘点表"
End With
```

**⑥ 完成**

```vba
If showMsg Then MsgBox "订货表已生成！"
```

**注意**：这个宏**会清空订货表全部内容**（包括用户填的期初/进货/报废）。所以按钮入口用 RegenOrder 加确认框。

#### 2. `RegenOrder()`

**作用**：订货表「生成订货表」按钮的入口，**先弹确认框**，确认后才调用生成订货表。

```vba
Sub RegenOrder()
    If MsgBox("确定要重新生成订货表吗？" & vbCrLf & _
        "期初库存/进货量/报废会被清空，期末库存将自动从盘点表同步。", _
        vbYesNo + vbQuestion, "生成订货表") = vbNo Then
        Exit Sub          ' 点「否」直接退出，什么都不做
    End If
    Call 生成订货表       ' 点「是」执行
End Sub
```

- `MsgBox(..., vbYesNo + vbQuestion, "生成订货表")`：弹出带「是/否」按钮的对话框，标题"生成订货表"
- 返回 `vbYes`（点"是"）才继续；`vbNo`（点"否"）直接 `Exit Sub`

**为什么按钮不直接绑「生成订货表」？**
「生成订货表」带 Optional 参数（showMsg），表单按钮调用时可能出兼容问题；用无参的 `RegenOrder` 包装一层最稳，还能加确认框。

#### 3. `读取上周数据()`
**作用**：从「订货历史」里找每个品号**最新一条**期末库存，填到订货表「期初库存」（E 列）。

```vba
orderLastRow = wsOrder.Cells(wsOrder.Rows.count, 1).End(xlUp).Row   ' 订货表最后一行
hisLastRow = wsHis.Cells(wsHis.Rows.count, 2).End(xlUp).Row          ' 历史表最后一行
If hisLastRow < 2 Then
    MsgBox "订货历史里还没有数据。"
    Exit Sub
End If
```

遍历订货表的产品行（A 列有品号 + C 列是数字），对每个品号在历史表里**从后往前**找：

```vba
For h = hisLastRow To 2 Step -1
    If wsHis.Cells(h, 2).Value = pinhao Then   ' B列=品号
        latestStock = wsHis.Cells(h, 3).Value  ' C列=期末库存
        Exit For                               ' 找到最新一条就停
    End If
Next h
If latestStock <> "" Then
    wsOrder.Cells(r, 5).Value = latestStock    ' 写入 E 列（期初库存）
End If
```

- **从后往前找**：历史表按时间追加，最后一行是最新，所以从最后往前找到的第一条就是最新
- 找不到该品号（历史里没有）→ 不覆盖，保持原值

#### 4. `保存本周数据()`
**作用**：把订货表「期末库存」（H 列）按「日期+品号」存入「订货历史」，**同一天同品号只保留一条**（防止重复点按钮产生重复行）。

```vba
today = Date                       ' 今天日期
For r = 3 To orderLastRow
    If pinhao <> "" And IsNumeric(wsOrder.Cells(r, 3).Value) And _
       wsOrder.Cells(r, 8).Value <> "" Then     ' 有品号 + 装箱数是数字 + 期末库存非空
        found = False
        For h = 2 To hisLastRow
            If wsHis.Cells(h, 1).Value = today And wsHis.Cells(h, 2).Value = pinhao Then
                wsHis.Cells(h, 3).Value = wsOrder.Cells(r, 8).Value   ' 更新已有记录
                found = True
                Exit For
            End If
        Next h
        If Not found Then                     ' 没有则追加新行
            hisLastRow = hisLastRow + 1
            wsHis.Cells(hisLastRow, 1).Value = today
            wsHis.Cells(hisLastRow, 2).Value = pinhao
            wsHis.Cells(hisLastRow, 3).Value = wsOrder.Cells(r, 8).Value
        End If
        count = count + 1
    End If
Next r
MsgBox "已保存 " & count & " 条数据到订货历史！"
```

- **判断产品行**：`pinhao <> "" And IsNumeric(C列)`，靠"装箱数是数字"识别产品行（表头/标题行装箱数为空或文字）
- **期末库存为空**（盘点没填）→ 跳过不保存（`Cells(r, 8).Value <> ""`）
- 历史表列：A=日期、B=品号、C=期末库存

#### 5. `生成盘点表(Optional ByVal showMsg As Boolean = True)`（盘点表同步）
**作用**：按产品目录重建盘点表；**保留店员已填的库存**（按品号迁移），行数随产品数量自动变化。

**参数 `showMsg`**：True 弹窗 / False 静默（打开文件时用 False）。

**开头：性能优化 + 错误处理**

```vba
On Error GoTo EH                     ' 出任何错跳到底部 EH 统一处理
Application.ScreenUpdating = False   ' 关闭屏幕刷新（加快速度）
Application.EnableEvents = False     ' 暂时禁用事件
Application.Calculation = xlCalculationManual  ' 手动计算（加快速度）
```

**第 1 步：记录现有库存（品号 → 库存）**

```vba
Dim oldStock As Object
Set oldStock = CreateObject("Scripting.Dictionary")   ' 字典：品号 → 库存
For r = 4 To lastRow
    key = Trim(CStr(wsInv.Cells(r, 2).Value))         ' B 列品号（去空格）
    If Len(key) > 0 Then
        If Not oldStock.Exists(key) Then              ' 同一品号只记一次
            oldStock.Add key, wsInv.Cells(r, 5).Value  ' E 列库存
        End If
    End If
Next r
```

> 为什么先记录？因为下一步要 `Cells.Clear` 清空整张盘点表，清空前必须把店员填的库存"记下来"，重建后按品号"还回去"。

**第 2 步：清空 + 重建**

```vba
wsInv.Cells.Clear
currentRow = 2                      ' 从第 2 行开始（第 1 行留空）
```

每个分类：
- 分类标题行（浅色底，颜色数组 `colors` 与产品目录条件格式一致）
- 列标题行（品号/产品名称/盘点单位/当前库存）
- 产品行：写品号/名称/单位 + **按品号回填库存**

```vba
If oldStock.Exists(key) Then
    wsInv.Cells(currentRow, 5).Value = oldStock(key)   ' 有旧值才回填
End If
```

**第 3 步：E 列数据验证**（店员只能填 ≥0 数字）

```vba
With wsInv.Cells(currentRow, 5)
    .Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
        Operator:=xlGreaterEqual, Formula1:="0"
```

**结尾：恢复设置**

```vba
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.ScreenUpdating = True
If showMsg Then MsgBox "盘点表已生成！"
Exit Sub

EH:                                  ' 出错时到这里
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If showMsg Then MsgBox "生成盘点表时出错：" & Err.Description
End Sub
```

**注意**：`EH:` 标签负责在**任何错误时恢复屏幕刷新/事件/计算模式**，避免出错后 Excel 卡在手动计算状态。

#### 6. `SyncInventory()`（盘点表同步）

**作用**：盘点表按钮的入口，调用 `生成盘点表(True)`（弹窗版）。

```vba
Sub SyncInventory()
    ' 供自动化测试/调用使用的英文入口
    Call 生成盘点表(True)
End Sub
```

**为什么用英文名？** 按钮 OnAction 写英文宏名更稳，避免中文宏名在某些环境下的编码兼容问题。

#### 7. `Workbook_Open()`（ThisWorkbook）

**作用**：打开文件时自动执行，静默同步盘点表。

```vba
Private Sub Workbook_Open()
    ' 打开文件时自动按产品目录同步盘点表（行数自动、库存按品号保留）
    On Error Resume Next              ' 出错也不弹任何提示（静默失败）
    Call 生成盘点表(False)            ' False = 不弹窗
End Sub
```

- `Private`：事件代码，不能被其他宏直接调用
- `On Error Resume Next`：若打开时出错（比如只读打开、宏被禁用），静默跳过，不打扰用户
- 用 `False` 参数：打开时不弹「盘点表已生成！」，避免每次打开都弹窗


### 三、如果你想修改代码

#### 安全流程
1. 先**复制备份**文件
2. 打开 VBA 编辑器：`Alt + F11`
3. 改完按 `Ctrl + S` 保存（会提示"启用宏"确认）
4. **先在副本上测试**，确认没问题再用于正式文件

#### 常见修改场景

**加/减一个分类**
- 修改两处数组：`生成订货表` 和 `生成盘点表` 里的
  `categories = Array("乳制品", "茶叶咖啡", ...)`
- 盘点表还要同步 `colors` 数组（颜色数量要和分类数量一致）
- 产品目录里的分类名必须与之一致

**改分类颜色**
- 盘点表：`colors` 数组，元素顺序与 `categories` 数组对应
- 订货表分类标题颜色在 `生成订货表` 里是固定 RGB(220,230,241)

**改公式逻辑**（比如订货量算法）
- 找到 `生成订货表` 里对应的 `.Formula = "..."` 行
- 注意 VBA 字符串里的引号：公式中的 `""` 在 VBA 里要写成 `""""`（用 `Chr(34)` 拼接更简单）

**调整盘点表/订货表从第几行开始**
- `currentRow = 2`（盘点表）或 `currentRow = 3`（订货表）就是起始行

#### 调试技巧
- `MsgBox` 打印变量：`MsgBox "当前行：" & currentRow`
- 逐步执行：VBA 编辑器里按 `F8` 单步运行
- 出错时看 `Err.Description`

