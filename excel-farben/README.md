# Excel 项目页着色宏（Farben_Makros）

给 Controlling 工作簿里的项目页（Projektbericht）统一上色。半自动：选中一个区块，按快捷键。
不需要安装任何软件，只用 Excel 自带的 VBA。

## 颜色规则

| 快捷键 | 区块 | 颜色 |
|---|---|---|
| Strg+Umschalt+K | 蓝色标题行（Gesamtbudget、IST Ausgaben、PLAN …） | Blau, Akzent 5, heller 60 %，Frutiger，加粗 |
| Strg+Umschalt+B | Gesamtbudget 下三行、gesamt 行、O 列绿色部分 | Grün, Akzent 6, heller 60 % |
| Strg+Umschalt+O | Obligo 下面的色块 | Blau, Akzent 1, heller 60 % |
| Strg+Umschalt+I | IST 下面的色块 | Orange, Akzent 2, heller 60 % |
| Strg+Umschalt+P | PLAN 下面的色块 | Gold, Akzent 4, heller 60 % |
| Strg+Umschalt+R | IST Vorjahre … Projektende（从 A 列选到金额列） | 左侧 Weiß, Hintergrund 1, dunkler 15 %；金额列 Grün, Akzent 6, heller 60 % |

第一大块（Projektbericht 到灰色表头结束）不变。

## 安装（只做一次）

1. 新建空白工作簿，按 **Alt+F11** 打开 VBA 编辑器（打不开说明宏被禁用）。
2. **Datei → Datei importieren…** 选 `Farben_Makros.bas`；或 **Einfügen → Modul**，把文件第一行之后的内容粘贴进去。
3. 另存为 `Farben_Makros.xlsm`（Excel-Arbeitsmappe mit Makros）。

## 使用

1. **先给要改的工作簿另存一个副本。宏的修改不能用 Strg+Z 撤销。**
2. 打开 `Farben_Makros.xlsm`（出现提示条时点 “Inhalt aktivieren”），再打开要改的工作簿。
3. **Alt+F8** 运行 `Tastenkuerzel_An`。
4. 在每张项目表上选中区块（按住 Strg 可多选），按对应快捷键。
5. 右下角的图例（Legende）：
   - **已有图例、只改颜色**：选中图例的各行（不含 “Legende” 标题），运行 `F_LegendeNeuFaerben`，按文字自动配色。
   - **新建或重写图例**：点图例起始单元格（如 O18），运行 `F_Legende`，按提示选择是否包含 Obligo 行。
     顺序固定为 Budget / Restbudget →（Obligo (gebunden)）→ IST Ausgaben (gebucht) → PLAN Ausgaben (Prognose)。
6. 做完后运行 `Tastenkuerzel_Aus`，恢复 Excel 原有快捷键。

## 全自动模式（推荐先在副本上试）

- `Auto_AktivesBlatt`：只处理当前显示的工作表，适合先测试。
- `Auto_AlleProjektseiten`：处理当前工作簿**从第 3 张起**的所有工作表。

识别规则：
1. **标题行**：A 列以 “Gesamtbudget” 开头的那一行 → 整行 Blau Akzent 5，Frutiger 字体，加粗。
2. **色块**：标题行里每个标题决定它下方**原本有颜色的格子**的新颜色，一直到 A 列写 “gesamt” 的那一行
   （中间的空白行若原本无色则不动）：
   …Budget… / …Mittel… / Gesamt… → 绿（Budget/Mittel 区块里**有数字的格子即使原来没颜色也涂绿**）；
   …Obligo… → Blau Akzent 1；…PLAN… / …PLN… → Gold；…IST… / …Ausg… → 橙；
   其他标题不动，并在报告中列出。标题行本身始终是蓝色。
3. **gesamt 行**：整行加粗。
4. **余额区**：gesamt 下方、A 列以 IST / PLAN / PLN / Restbudget / Projektmittel 开头的几行 →
   左侧 Weiß dunkler 15 %，第一列金额 Grün Akzent 6，全部加粗。
5. **图例**：找到 “Legende”，按下面各行的文字自动配色。

运行结束后会打开一个**新的工作簿显示报告**：每张表改了哪些区域、哪些没识别出来需要手动处理。

**安全措施**
- **自动备份**：`Auto_AktivesBlatt` / `Auto_AlleProjektseiten` 在改任何东西之前，先在同一文件夹存一份副本
  `<文件名>_vor_Farben_<日期_时间>.xlsx`（当前状态，含未保存的修改）。备份失败则不做任何修改。
- **只处理项目页**：必须同时满足 ① A1（或 A2）以 “Projektbericht” 开头 ② A 列有 “Gesamtbudget” 行
  ③ 其下有 “gesamt” 行 ④ 工作表未保护。`Auto_AlleProjektseiten` 运行前会列出“将处理 / 将跳过（及原因）”供确认。

## 辅助：查 Personal-Ressourcenplanung 的 Cost 列

改 `SUMME('Personal-Ressourcenplanung'!…)` 公式时用来核对列号。

- `Hilfe_PRP_AlleBlaetter`：在第 3 张起的每张项目表**已用区域右侧的空白处**写一个小块：
  Okt 26 / Dez 26 / Jan 27 / Dez 27 / Jan 28 的 Cost 列、Laufzeit von/bis、**项目结束月的 Cost 列**，
  以及常用区间（如 Okt 26 – Dez 26、Okt 26 – Projektende）。
- `Hilfe_PRP_AktivesBlatt`：只处理当前表（在 Personal-Ressourcenplanung 上运行时只弹出列号）。
- `Hilfe_Loeschen`：一键删除所有这些小块（只清空带 `### Hilfe PRP-Spalten ###` 标记的块的内容，不动格式）。

月份列表可在代码开头 `HILFE_MONATE` 修改（格式 `月/年`，逗号分隔）。

## 撤销（多级）

- `Rueckgaengig_LetzterLauf`（快捷键开启后为 **Strg+Umschalt+Z**）：撤销最近一次运行；**再运行一次就撤销倒数第二次**，以此类推（最多记 30 次）。
- `Rueckgaengig_Liste`：显示目前记住了哪些运行。
- 适用于所有着色宏（快捷键宏、`F_Legende`、`F_LegendeNeuFaerben`、`Auto_…`）。恢复的是填充色、字体、加粗；图例还会恢复原来的文字。
- **限制**：记录只在 Excel 打开期间有效。关闭 Excel、修改 VBA 代码，或在错误提示里点了 “Beenden”，记录都会清空——这时用自动备份恢复。
- 如果运行宏之后又手动改了同一些格子，撤销会把它们也恢复成运行宏之前的样子。

## 首次使用前检查字体名

宏里标题行的字体名是猜的（`Frutiger LT 45 Light`）。在创建者改好的表里点一下蓝色标题行，
运行 `FarbeAnzeigen`，把弹窗里显示的字体名填到代码开头的 `SCHRIFT_KOPF` 里。

## 注意

- 宏使用工作簿的**主题颜色**；若某个工作簿主题不是 Office 默认主题，颜色看起来会不同。
- `F_Restbudget` 把选区**最右边一列**当作金额列。
- 只改填充色（标题行另改字体、加粗），不改数值、公式和数字格式。
