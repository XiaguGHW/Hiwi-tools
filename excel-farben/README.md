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
   Gesamtbudget / Gesamt… → 绿；…Obligo… → Blau Akzent 1；…PLAN… / …PLN… → Gold；…IST… / …Ausg… → 橙；
   其他标题（如 “Verfügbares Budget”）不动，并在报告中列出。
3. **gesamt 行**：整行加粗。
4. **余额区**：gesamt 下方、A 列以 IST / PLAN / PLN / Restbudget / Projektmittel 开头的几行 →
   左侧 Weiß dunkler 15 %，第一列金额 Grün Akzent 6，全部加粗。
5. **图例**：找到 “Legende”，按下面各行的文字自动配色。

运行结束后会打开一个**新的工作簿显示报告**：每张表改了哪些区域、哪些没识别出来需要手动处理。

## 首次使用前检查字体名

宏里标题行的字体名是猜的（`Frutiger LT 45 Light`）。在创建者改好的表里点一下蓝色标题行，
运行 `FarbeAnzeigen`，把弹窗里显示的字体名填到代码开头的 `SCHRIFT_KOPF` 里。

## 注意

- 宏使用工作簿的**主题颜色**；若某个工作簿主题不是 Office 默认主题，颜色看起来会不同。
- `F_Restbudget` 把选区**最右边一列**当作金额列。
- 只改填充色（标题行另改字体、加粗），不改数值、公式和数字格式。
