# Intel (INTC) 投行级三表DCF财务模型

**Intel 公司投资银行级别财务模型** | 细分业务驱动架构 | 情景切换（Base / Bull / Bear） | 季度+年度混合颗粒度

基于 Intel FY2022–FY2025 10-K 年报、Q1 2026 10-Q 季报及 Q1 2026 Earnings Call 构建。模型以**自下而上（Bottom-Up）**的增长驱动因子（TAM × 市场份额 × ASP）替代简单的增长率外推，并通过 `CHOOSE/MATCH` 公式实现一键情景切换。

---

## 项目说明

本项目是一个完整的 Intel (NASDAQ: INTC) 投行级财务模型，使用 Python `openpyxl` 库程序化生成 Excel 工作簿。模型涵盖：

- **11 个 Tab**：Cover → Key_Summary → Assumptions → Growth_Drivers → Segment_PL → Consolidated_PL → BS → Cash_Flow → DCF → Sensitivity → Ratio_Analysis
- **季度+年度混合颗粒度**：Growth_Drivers、Assumptions、Segment_PL 采用 35 列季度布局（Q1–Q4 + FY 年度汇总），其余 Tab 保持年度布局
- **三表联动**：利润表（P&L）→ 资产负债表（BS）→ 现金流量表（CF），BS 以 "现金作为平衡项" 确保 A = L+E
- **双重 P&L 交叉验证**：Segment_PL（按业务分部）与 Consolidated_PL（按费用类型）的 EBIT 必须一致
- **DCF 估值**：从 UFCF → TV → EV → 股权价值 → 每股隐含价格

预测期：**FY2026E – FY2028E**（3 年），基于 FY2025A 历史基准。

---

## 主交付件

### `Intel_IB_Model_v2.xlsx` — 11-Tab Excel 模型
初版模型预测十分粗糙，财务预测颗粒度只到年度，并且是分部收入的简单线性外推，是无法直接用于估值的。V2版本将盈利预测颗粒度拆到季度，并且各个分部都找到了重要的driver，基本微调即可用于估值。

| # | Tab | 内容说明 | 是否可编辑 |
|---|-----|---------|-----------|
| 1 | **Cover** | 封面、公司信息、使用说明 | 只读 |
| 2 | **Key_Summary** | 高管仪表盘：核心 KPI + 分业务收入及占比 + BS 要点 + ROE/ROA + 投资要点 |
| 3 | **Assumptions** | **情景控制中心**：Base/Bull/Bear 三情景 × 季度颗粒度，所有假设参数（TAM、份额、ASP、OPM%、费用率、CapEx、WC 等），每个参数一行 Label + 三行情景值 + CHOOSE/MATCH 选择行 |
| 4 | **Growth_Drivers** | **自下而上收入拆解**：每个业务分部的收入由 TAM × 份额 × ASP 等底层驱动因子计算得出，含历史调节验证（Implied vs Reported）和驱动因子弹性分析 | 
| 5 | **Segment_PL** | **分部 P&L（收入 + 营业利润）**：CCG / DCAI / Intel Foundry / All Other 四个分部的季度收入与 OI，含 Corporate Unallocated、合并收入、交叉验证行 | 
| 6 | **Consolidated_PL** | 合并利润表（按费用类型）：COGS / R&D / MG&A → EBIT（来自 Segment_PL Sum OI）→ NI → EPS，含毛利率/EBIT 率/净利率 |
| 7 | **BS** | 资产负债表：流动资产（现金为平衡项）+ 非流动资产 + 流动负债 + 非流动负债 + 权益，含 Balance Check 行 |
| 8 | **Cash_Flow** | 现金流量表：CFO（含 22 项 WC 变动）+ CFI（CapEx）→ FCF |
| 9 | **DCF** | 无杠杆 DCF 估值：WACC 参数 + NOPAT + UFCF + 折现因子 + 终值 + EV → 隐含股价 |
| 10 | **Sensitivity** | 二维数据表：WACC × Terminal Growth → 隐含股价矩阵 | Excel 数据表 |
| 11 | **Ratio_Analysis** | 财务比率分析：盈利能力 / 效率 / 杠杆 / 现金流质量 / 增长率 / **估值倍数**（TTM & NTM 的 P/E, P/B, P/S, EV/EBITDA, P/CF, PEG） |

> \*Ratio_Analysis 中 **F 列（NTM/FY2026E）Stock Price** 为蓝色填充单元格，可编辑当前股价，所有倍数自动更新。

### 模型核心特色

- **季度颗粒度**：Growth_Drivers、Assumptions、Segment_PL 三个 Tab 采用 Q1–Q4 + FY 年度汇总的 35 列布局，用户可在 Assumptions 中独立调整每个季度的假设（如 Q1 和 Q4 的 ASP 不同），年度值由 `=SUM(Q1:Q4)` 或 `=AVERAGE(Q1:Q4)` 汇总得出
- **自下而上收入驱动**：不依赖 `Prior × (1+g)` 的简单外推。CCG 收入 = PC TAM × Intel 份额 × ASP；DCAI 收入 = Server TAM × DC 份额 × ASP + AI/ASIC 收入。每个驱动因子均可独立调整
- **三情景切换**：Assumptions!B2 下拉选择 Base / Bull / Bear，288 个 CHOOSE/MATCH 公式自动更新全模型
- **双重 P&L 交叉验证**：Segment_PL Sum of Segment OI ≡ Consolidated_PL EBIT（所有年度列 Difference = 0）
- **Intel Foundry 内部销售处理**：Foundry 分部收入 ~99% 为内部晶圆销售给 Intel Products 部门。模型单独追踪 External Foundry Revenue（第三方客户），合并收入 = CCG + DCAI + All Other + External Foundry + Eliminations
- **投行规范格式**：蓝色字体 = 历史硬编码数据，绿色字体 = 跨表链接，蓝色底色 = 可编辑假设，黄色底色 = 关键输出行，灰色斜体 = 百分比/备注行
- **现金流：BS 差额驱动**：22 项营运资金变动全部链接到 BS 的期初期末差额
- **DCF 估值完整链条**：NOPAT + D&A − CapEx − ΔNWC = UFCF → Gordon Growth TV → EV → − 净债务 → 每股隐含价格
- **WACC 参数内置 DCF Tab**：Rf、Beta、ERP、Kd、税率、D/TC、Terminal g 等估值参数位于 DCF Tab 顶部，与 Assumptions 中的经营假设分离
- **驱动因子弹性分析**：Growth_Drivers 底部展示每个关键驱动因子 ±1% 或 ±$1 变动对收入的影响金额和百分比
- **历史数据锚定**：所有假设行均展示 2023A–2025A 三个完整财年的历史数据作为审阅基准

---

## 年报文件（数据来源）

| 文件 | 年份 | 说明 |
|------|------|------|
| Intel 10-K FY2022 | FY2022 (ended Dec 31, 2022) | 历史财务数据 |
| Intel 10-K FY2023 | FY2023 (ended Dec 30, 2023) | 历史财务数据 |
| Intel 10-K FY2024 | FY2024 (ended Dec 28, 2024) | 历史财务数据 |
| Intel 10-K FY2025 | FY2025 (ended Dec 27, 2025) | 最新完整财年，TTM 基准 |
| Intel 10-Q Q1 2026 | Q1 2026 (ended Mar 29, 2026) | 最新季度数据，含分部重述 |
| Intel Q1 2026 Earnings Call | April 23, 2026 | 管理层指引（CapEx、Foundry 进展、18A 节点） |

> 历史分部收入已按 2025 Q1 8-K 的重述口径调整（NEX 整合至 DCAI、Altera 出表）。

---

## 数据图例（IB 格式标准）

| 格式 | 含义 | 说明 |
|------|------|------|
| 🔵 蓝色字体 (#305496) | 财报原始数据（Hardcoded） | 来自 10-K/10-Q 的历史实际数 |
| 🟢 绿色字体 (#548235) | 跨表公式链接 | 引用其他 Tab 的单元格（如 `=Segment_PL!F36`） |
| 🔵 蓝色底色 (#DDEBF7) | 可编辑假设 | **用户可修改的输入值** |
| 🟡 黄色底色 (#FFF2CC) | 关键输出行 | 模型最重要的计算结果 |
| ⚫ 黑色字体 | 计算公式 | Tab 内部的 openpyxl 公式 |
| ⬜ 灰色斜体 (#808080) | 百分比/备注行 | 利润率、YoY%、Mix% 等 |

---

## 使用方式

### 环境准备

```bash
pip install openpyxl
```

### 生成模型

```bash
python build_intc_model.py
```

输出文件：`Intel_IB_Model_v2.xlsx`（11 Tabs，约 934 行）

### 操作步骤

1. 用 Excel 打开 `Intel_IB_Model_v2.xlsx`
2. 进入 **Assumptions** Tab
3. 在 **B2 单元格**下拉菜单中选择情景：`Base`（基准）/ `Bull`（乐观）/ `Bear`（悲观）
4. 所有 Tab 自动重算（288 个 CHOOSE/MATCH 公式联动更新）
5. 查看 **DCF Tab** → 隐含每股价格（Implied Share Price）
6. 查看 **Sensitivity Tab** → WACC × Terminal Growth 价格矩阵
7. 如需自定义假设：在 Assumptions Tab 中修改蓝色底色单元格（季度级别可独立调整）
8. 查看 **Key_Summary Tab** → 高管仪表盘一览
9. 查看 **Ratio_Analysis Tab** → F 列输入当前股价，所有倍数（P/E, EV/EBITDA 等）自动更新

### 情景说明

| 情景 | 核心假设 | 适用场景 |
|------|---------|---------|
| **Base** | PC 温和复苏、DCAI CPU 复兴（+18% FY26）、Foundry 亏损收窄（OPM −42%→−22%）、CapEx $17.7–21.0B | 基准估值 |
| **Bull** | DCAI 强劲增长（+25% FY26）、Foundry 18A 外部客户加速导入、良率超预期、CapEx 扩至 $25B | 乐观估值 |
| **Bear** | PC 衰退（−3% FY26 CCG）、AMD 份额持续增长、供应链受限、Foundry 亏损维持深度（−50%→−40%） | 压力测试 |

### 关键验证行（打开模型后必查）

- **Segment_PL**：CROSS-CHECK 区域 → Difference 行 = 0（所有年度列）
- **BS**：Balance Check 行 = 0（A = L+E）
- **Growth_Drivers**：各分部 2023A–2025A 历史 Reconciliation < 1%

---

## 模型截图

详细截图请参见 `screenshots/` 文件夹（建议至少截取以下 Tab）：

| Tab | 截图内容 |
|-----|---------|
| Key_Summary | 完整仪表盘（KPIs + 分部收入 + BS 要点 + 投资要点） |
| Assumptions | 季度假设（展示 Base/Bull/Bear 三行 + Selected 行 + 情景下拉框） |
| Growth_Drivers | CCG 分部自下而上拆解（展示 TAM → Share → Units → ASP → Implied Rev 完整链条） |
| Segment_PL | 四个分部季度收入 + OI（展示 CCG 和 DCAI 的完整季度数据） |
| DCF | UFCF 链条 → EV → 隐含股价 |
| Sensitivity | WACC × g 二维价格矩阵 |

---

## 辅助文件与文档

### Python 构建脚本

| 文件 | 说明 |
|------|------|
| `build_intc_model.py` | **主构建脚本**（~3,500 行，32 个函数）—— 当前版本，支持季度布局 + 11 Tab |
| `build_intc_model_final.py` | 早期版本（~2,545 行，年度布局，12 Tab）—— 保留以供参考 |

### 模型输出

| 文件 | 说明 |
|------|------|
| `Intel_IB_Model_v2.xlsx` | 最新生成的 Excel 模型（11 Tab，934 行，288 CHOOSE/MATCH） |
| `Intel_IB_Model.xlsx` | 早期 12-Tab 年度版本 —— 保留以供参考 |

### 设计文档

| 文件 | 说明 |
|------|------|
| `intc_financial_data.md` | Intel 财务数据提取（10-K 合并 P&L、BS、CF、分部数据 + 10-Q 季度分部收入） |
| `intc_assumptions.md` | 预测假设设计文档（Base/Bull/Bear 三情景的增长率、利润率、驱动因子） |
| `example-structure.md` | 年度布局参考架构（11-Tab 逐行映射 + 跨表引用追踪模板） |
| `model_summary.md` | 模型详细摘要（假设表、输出表、架构要点、验证状态） |
| `progress.md` | 构建进度日志（所有阶段记录） |
| `SKILL-ib-financial-model.md` | IB 财务模型 Skill 定义（可用于快速复刻到其他公司如 AMD、NVDA） |

### 验证工具

| 文件 | 说明 |
|------|------|
| `verfity tools/` | 验证脚本文件夹（BS 平衡检查、收入桥验证、现金公式检查等） |

### 参考资料

| 文件/目录 | 说明 |
|-----------|------|
| `company filings/` | Intel 年报/季报 PDF 原文 |

---

## 模型架构深度说明

### 季度布局设计（35 列）

```
B-Q1'23 C-Q2'23 D-Q3'23 E-Q4'23 F-FY23  G-Q1'24 ... P-FY25  Q=spacer
R-Q1'26 S-Q2'26 T-Q3'26 U-Q4'26 V-FY26  W-Q1'27 ... AF-FY28  AG=Sel AH=sp AI=Notes
```

Growth_Drivers、Assumptions、Segment_PL 使用此季度布局。年度汇总列（F/K/P/V/AA/AF）链接到其他年度 Tab。

### 跨布局列映射

由于季度 Tab 和年度 Tab 的列索引不同，必须使用映射字典：

```python
# 年度 Tab 列 → Segment_PL 季度年度汇总列
SEGPL_COL_MAP = {2: 6, 3: 11, 4: 16, 6: 22, 7: 27, 8: 32}

# 年度 Tab 预测列 → Assumptions 季度年度汇总列
ASP_ANN_COLS = {6: 22, 7: 27, 8: 32}
```

### 多轮构建流程

```
Pass 1: Assumptions (情景数据)
Pass 2: Segment_PL 第1轮 (结构 + 历史值, 分配 R-keys)
Pass 3: Growth_Drivers (引用 Segment_PL R-keys 获取历史数据)
Pass 4: Segment_PL 第2轮 (预测季度收入链接到 Growth_Drivers)
Pass 5: Consolidated_PL (通过 SEGPL_COL_MAP 引用 Segment_PL 年度列)
Pass 6: Segment_PL 第3轮 (填入交叉验证行的 Consolidated_PL EBIT 链接)
Pass 7: BS → Cash_Flow → DCF → Sensitivity → Ratio_Analysis
Pass 8: Key_Summary (最后构建，链接所有 Tab)
```

### Intel Foundry 收入桥（关键架构决策）

Intel Foundry 的分部报告收入（FY2025 ~$17.8B）**约 99% 为内部**晶圆销售（销售给 Intel Products 部门 CCG/DCAI）。模型处理方式：

- **Segment_PL** 追踪 **Total Foundry Revenue**（含内部销售）用于计算分部 OI
- **External Foundry Revenue** 单独一行（FY2025 ~$150M，FY2028E 增长至 ~$2.5B）
- **合并收入** = CCG + DCAI + All Other + External Foundry + Eliminations
- 历史合并收入从 10-K 直接硬编码（避免内部销售重复计算）

---

## 已知限制

1. **预测期较短**：3 年（FY2026E–FY2028E）。Intel 的代工转型属于资本密集型，可能需要 5 年以上才能充分反映 CapEx 回报
2. **DCF 估值敏感**：高债务负担（FY2025A ~$46.6B）+ 高 CapEx 意味着股权价值对 WACC 和终端增长率高度敏感
3. **External Foundry 收入**：目前几乎可以忽略（~$150M FY2025），增长至 FY2028E $2.5B 的预测有较大不确定性
4. **税率假设**：当前 ETR ~12%（因估值备抵），盈利恢复后可能正常化至 ~21%
5. **非经营性项目**：政府补贴（USG Agreement）、重组费用、NCI 等做了简化处理
6. **季度历史数据**：Q2–Q3 2025 的季度分部收入为插值估算（年度总额 − Q1 实际 − Q4 估计后均分）

---

## 模型验证状态

| 检查项 | 状态 |
|--------|------|
| 11 个 Tab 完整 | ✅ PASS |
| Segment_PL 交叉验证（Sum Seg OI − Consol EBIT = 0） | ✅ PASS（6 个年度列） |
| BS Balance Check（A = L+E） | ✅ PASS |
| EBIT 链接正确（Consolidated_PL → Segment_PL Sum OI） | ✅ PASS |
| 现金流量表 WC 链接到 BS（22 项引用） | ✅ PASS |
| DCF UFCF = NOPAT + D&A − CapEx − ΔNWC | ✅ PASS |
| Notes 列无公式（季度 Tab: AI=35, 年度 Tab: K=11） | ✅ PASS |
| CHOOSE/MATCH 公式数量 | ✅ PASS（288 个） |
| 情景切换器（Data Validation 下拉菜单） | ✅ PASS |
| Segment_PL 预测公式引用 Assumptions | ✅ PASS（57 个引用） |

---

## 扩展到其他公司

如需为类似业务结构的半导体公司（如 AMD、NVDA）建模，可基于本项目快速复刻：

1. 复制 `build_intc_model.py` 作为模板
2. 更新 `GD_QREV` 为目标公司的季度分部数据
3. 修改 `build_assumptions()` 中的驱动因子框架（如 AMD: EPYC 份额、Radeon ASP 等）
4. 修改 `build_segment_pl()` 中的分部列表和历史值
5. 修改 `build_growth_drivers()` 中的驱动公式
6. 更新所有历史硬编码值
7. 运行 `build_model()` → 验证 Self-Check 全部通过

详细指南参见 `SKILL-ib-financial-model.md`。

---

*Generated via Claude Code / ib-financial-model skill | Intel 10-K/Q data | 构建日期: May 2026*
