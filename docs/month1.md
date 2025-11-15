## 总体说明

- 此报表通过 **`month1carbone.xlsx`** 模板生成，模板中已配置 carbone 标记。
- 此报表有两个数据源：
  1. 第一数据源：工作簿【放款明细】的 Sheet1 工作表（下称“放款明细表”）；
  2. 第二数据源：工作簿【资产明细】的 Sheet1 工作表（下称“资产明细表”）。
- 模板变量在文件中通常以 `d.xxx` 形式出现（如 `{d.monthlySummary.statYear}`）。下面的变量说明中，默认省略前缀 `d.`。

---

## 第一部分：`monthlySummary`

### 模板文本

```
{d.monthlySummary.statYear}年{d.monthlySummary.statMonth}月公司新增放款【{d.monthlySummary.newLoanAmount}】万元（基建占比【{d.monthlySummary.newInfraRatio}】%，医药占比【{d.monthlySummary.newMedicalRatio}】%，再保理占比【{d.monthlySummary.newRefactoringRatio}】%），
受让对应的应收账款【{d.monthlySummary.newArAssignedAmount}】万元，平均保理利率【】%，新增合作客户【{d.monthlySummary.newCoopCustomerCount}】家，受理业务【{d.monthlySummary.newAcceptedBusinessCount}】笔，已发生实际业务合作的核心企业新增【】家；

```

> 说明：其中“平均保理利率【】%”和“核心企业新增【】家”为文字占位，模板中不使用变量，保持为空。

### 变量说明（按出现顺序）

1. **`monthlySummary.statYear`**

   用户输入的统计年份，例如：2025。

2. **`monthlySummary.statMonth`**

   用户输入的统计月份（1–12，或“01”–“12”）。

   后续提到“P 列（实际放款日期）属于用户输入的年月”统一定义为：

   > 从放款明细表中筛选：
   >
   > P 列“实际放款日期”的年份 = `monthlySummary.statYear`，且月份 = `monthlySummary.statMonth`。

3. **`monthlySummary.newLoanAmount`**（本月新增放款金额，单位：万元）
   - 在放款明细表中，筛选出 P 列“实际放款日期”属于用户输入的年月的所有行；
   - 对这些行的 AW 列“放款金额”求和（假定 AW 单位为元）；
   - 将求和结果除以 10000，得到 `monthlySummary.newLoanAmount`（万元）。
4. **`monthlySummary.newInfraRatio`**（基建占比，单位：%）
   - 在放款明细表中，筛选出：
     - P 列“实际放款日期”属于用户输入的年月，且
     - AA 列“所属行业” = “基建工程” 的所有行；
   - 对这些行 AW 列“放款金额”求和，得到本月“基建工程”放款金额（元），记为 `infraAmount`;
   - 若 `monthlySummary.newLoanAmount > 0`，则：
     > monthlySummary.newInfraRatio = infraAmount ÷ (monthlySummary.newLoanAmount × 10000) × 100
     即：行业金额 ÷ 本月总放款金额 × 100，结果为百分数数值（如 35.6 表示 35.6%）；
   - 若 `monthlySummary.newLoanAmount = 0`，则 `monthlySummary.newInfraRatio` 记为 0。
5. **`monthlySummary.newMedicalRatio`**（医药占比，单位：%）
   - 条件同上，将 AA 列行业改为 “医药医疗”；
   - 求和得到 `medicalAmount` 后：
     > monthlySummary.newMedicalRatio = medicalAmount ÷ (monthlySummary.newLoanAmount × 10000) × 100（分母同上，0 时记为 0）。
6. **`monthlySummary.newRefactoringRatio`**（再保理占比，单位：%）
   - 条件同上，将 AA 列行业改为 “大宗商品”；
   - 求和得到 `refactoringAmount` 后：
     > monthlySummary.newRefactoringRatio = refactoringAmount ÷ (monthlySummary.newLoanAmount × 10000) × 100（分母同上，0 时记为 0）。
7. **`monthlySummary.newArAssignedAmount`**（本月受让应收账款金额，单位：万元）
   - 在放款明细表中，筛选：P 列“实际放款日期”属于用户输入的年月的所有行；
   - 对这些行 AR 列“转让总金额”求和（元）；
   - 将求和结果除以 10000，得到 `monthlySummary.newArAssignedAmount`（万元）。
8. **`monthlySummary.newCoopCustomerCount`**（本月新增合作客户数）
   - 在放款明细表中，筛选：P 列“实际放款日期”属于用户输入的年月的所有行；
   - 取这些行的 C 列“保理/再保理申请人名称”；
   - 去重后统计**不为空**的名称数量，得到 `monthlySummary.newCoopCustomerCount`。
9. **`monthlySummary.newAcceptedBusinessCount`**（本月受理业务笔数）
   - 在放款明细表中，筛选：P 列“实际放款日期”属于用户输入的年月的所有行；
   - 取这些行的 N 列“保理融资申请书合同编号”；
   - 去重后统计**不为空**的编号数量，得到 `monthlySummary.newAcceptedBusinessCount`。

> 注意：第一部分中其他以【】包裹但没有变量名的位置（平均保理利率、核心企业新增数），保持为空，不做变量定义与计算。

---

## 第二部分：`asOfDateSummary`

### 模板文本

```
截止{d.asOfDateSummary.queryYear}年{d.asOfDateSummary.queryMonth}月{d.asOfDateSummary.queryDay}日，公司应收账款余额【】万元，放款余额【】万元，业务不良率为【】%,
累计受让应收账款金额【】万元，累计放款【{d.asOfDateSummary.cumLoanAmount}】万元（基建占比【{d.asOfDateSummary.infraRatio}】% ,医药占比【{d.asOfDateSummary.medicalRatio}】%,再保理占比【{d.asOfDateSummary.refactoringRatio}】%）,累计还款【】万元；

```

> 说明：除文中 {d.asOfDateSummary.cumLoanAmount}、{d.asOfDateSummary.infraRatio}、{d.asOfDateSummary.medicalRatio}、{d.asOfDateSummary.refactoringRatio} 四个变量外，其余【】为文字占位，不引入变量。

### 变量说明（按出现顺序）

1. **`asOfDateSummary.queryYear`**

   用户输入的统计年份（例如 2025）。

2. **`asOfDateSummary.queryMonth`**

   用户输入的统计月份（1–12）。

3. **`asOfDateSummary.queryDay`**

   用户输入的统计日期（1–31）。

   三者共同构成一个日期 `queryDate`，后续“早于等于 用户输入的日期”统一表示为：

   > 放款明细表中 P 列“实际放款日期” ≤ queryDate（按完整日期比较，含当日）。

4. **`asOfDateSummary.cumLoanAmount`**（截至 queryDate 的累计放款金额，单位：万元）
   - 在放款明细表中，筛选：P 列“实际放款日期” ≤ `queryDate` 的所有行；
   - 对这些行的 AW 列“放款金额”求和（元）；
   - 将结果除以 10000，得到 `asOfDateSummary.cumLoanAmount`（万元）。
5. **`asOfDateSummary.infraRatio`**（累计基建占比，单位：%）
   - 在放款明细表中，筛选：
     - P 列“实际放款日期” ≤ `queryDate`；
     - 且 AA 列“所属行业” = “基建工程”；
   - 对这些行 AW 列“放款金额”求和（元），记为 `cumInfraAmount`；
   - 若 `asOfDateSummary.cumLoanAmount > 0`，则：
     > asOfDateSummary.infraRatio = cumInfraAmount ÷ (asOfDateSummary.cumLoanAmount × 10000) × 100
   - 若 `asOfDateSummary.cumLoanAmount = 0`，则 `asOfDateSummary.infraRatio` 记为 0。
6. **`asOfDateSummary.medicalRatio`**（累计医药占比，单位：%）
   - 条件同上，将 AA 列行业改为 “医药医疗”；
   - 求得 `cumMedicalAmount` 后，公式同 5（分母为 `asOfDateSummary.cumLoanAmount × 10000`，为 0 时结果记为 0）。
7. **`asOfDateSummary.refactoringRatio`**（累计再保理占比，单位：%）
   - 条件同上，将 AA 列行业改为 “大宗商品”；
   - 求得 `cumRefactoringAmount` 后，公式同 5（分母为 `asOfDateSummary.cumLoanAmount × 10000`，为 0 时结果记为 0）。

> 第二部分中其他【】（应收账款余额、放款余额、业务不良率、累计受让应收账款金额、累计还款）保持为空，占位文本，不定义变量与计算逻辑。

---

## 第三部分：`sinceInceptionSummary`

### 模板文本

```
展业至今，公司合作客户【】家，累计受理业务【{d.sinceInceptionSummary.cumAcceptedBusinessCount}】笔，已发生实际业务合作的核心企业【】家，
累计受让应收账款【{d.sinceInceptionSummary.cumArAssignedAmount}】亿元，累计放款【{d.sinceInceptionSummary.cumLoanAmount}】亿元（基建占比【{d.sinceInceptionSummary.infraRatio}】%，医药占比【{d.sinceInceptionSummary.medicalRatio}】%，再保理占比【{d.sinceInceptionSummary.refactoringRatio}】%）；放款余额【】亿元；业务不良率为【】%。

```

> 说明：本段中只对已有的 6 个变量进行计算说明，其余【】为文字占位，不补充变量。

### 变量说明（按出现顺序）

> “展业至今”统一理解为：基于数据源中全部历史数据（不按日期过滤）。

1. **`sinceInceptionSummary.cumAcceptedBusinessCount`**（累计受理业务笔数）
   - 在放款明细表中，取 N 列“保理融资申请书合同编号”；
   - 对 N 列值去重，统计**不为空**的编号个数，得到 `sinceInceptionSummary.cumAcceptedBusinessCount`。
2. **`sinceInceptionSummary.cumArAssignedAmount`**（累计受让应收账款金额，单位：亿元）
   - 在资产明细表中，取 AD 列“转让总金额”；
   - 对全表 AD 列求和（元）；
   - 将结果除以 100000000（即一亿），得到 `sinceInceptionSummary.cumArAssignedAmount`（亿元）。
3. **`sinceInceptionSummary.cumLoanAmount`**（累计放款金额，单位：亿元）
   - 在放款明细表中，取 AW 列“放款金额”；
   - 对全表 AW 列求和（元）；
   - 将结果除以 100000000（即一亿），得到 `sinceInceptionSummary.cumLoanAmount`（亿元）。
4. **`sinceInceptionSummary.infraRatio`**（累计基建占比，单位：%）
   - 在放款明细表中，筛选：AA 列“所属行业” = “基建工程”的所有行；
   - 对这些行 AW 列“放款金额”求和（元），记为 `totalInfraAmount`；
   - 设 `totalLoanAmount` 为放款明细表中 AW 列全部行的和（元）；
   - 若 `totalLoanAmount > 0`，则：
     > sinceInceptionSummary.infraRatio = totalInfraAmount ÷ totalLoanAmount × 100
   - 若 `totalLoanAmount = 0`，则 `sinceInceptionSummary.infraRatio` 记为 0。
5. **`sinceInceptionSummary.medicalRatio`**（累计医药占比，单位：%）
   - 同上，将 AA 列行业改为 “医药医疗”；
   - 求得 `totalMedicalAmount` 后：
     > sinceInceptionSummary.medicalRatio = totalMedicalAmount ÷ totalLoanAmount × 100（分母同上，0 时结果记为 0）。
6. **`sinceInceptionSummary.refactoringRatio`**（累计再保理占比，单位：%）
   - 同上，将 AA 列行业改为 “大宗商品”；
   - 求得 `totalRefactoringAmount` 后：
     > sinceInceptionSummary.refactoringRatio = totalRefactoringAmount ÷ totalLoanAmount × 100（分母同上，0 时结果记为 0）。

> 第三部分中其他【】（合作客户、核心企业、放款余额、业务不良率）保持为空，占位文本，不新增变量和口径。
