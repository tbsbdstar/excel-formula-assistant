/* 工作流配置器：原子操作目录。
 * 等号后面必须是严格 JSON（键名加双引号、无注释、无尾逗号）：executor/tests/test_catalog.py 会读取它，
 * 与执行器实际注册的操作和参数逐一核对，防止两边不一致。
 * keys：该字段写入配置的参数名（复合字段对应多个参数）；hidden：执行器支持、但界面不提供编辑的参数（导入时原样保留）。 */
(window.EFA = window.EFA || {}).WF_CATALOG = {
  "groups": [
    {"id": "col", "label": "列"},
    {"id": "row", "label": "行"},
    {"id": "fill", "label": "填充与格式"},
    {"id": "match", "label": "跨表匹配与比对"},
    {"id": "merge", "label": "多文件合并与分组汇总"},
    {"id": "copy", "label": "复制区块 / 工作表"},
    {"id": "check", "label": "校验"}
  ],
  "ops": [
    {"op": "delete_columns", "label": "删除列", "group": "col",
     "desc": "删除一列或多列。其他单元格里引用到它们的公式会像 Excel 一样自动调整。",
     "fields": [
       {"key": "columns", "type": "columns", "label": "要删除的列", "required": true}
     ]},
    {"op": "insert_column", "label": "插入列", "group": "col",
     "desc": "在指定位置插入一列，可同时写入固定值、公式或另一列的值。",
     "fields": [
       {"key": "header", "type": "text", "label": "新列的表头", "required": true, "placeholder": "如：备注"},
       {"key": "position", "type": "position", "keys": ["before", "after", "at", "end"], "label": "插入位置", "required": true},
       {"key": "content", "type": "content", "keys": ["value", "formula", "formula_from", "from_col"], "label": "新列写入", "allowNone": true},
       {"key": "where", "type": "condition", "label": "只在符合条件的行写入", "help": "不设置=所有数据行都写入"},
       {"key": "style_from", "type": "enum", "label": "新列的格式", "default": "left", "allowColumn": true,
        "options": [{"value": "left", "label": "沿用左边一列的格式"}, {"value": "right", "label": "沿用右边一列的格式"}, {"value": "none", "label": "不带格式"}]},
       {"key": "width", "type": "number", "label": "列宽", "placeholder": "不填=沿用", "min": 0}
     ]},
    {"op": "rename_header", "label": "修改表头", "group": "col",
     "desc": "把某一列的表头改成新名字。后面的步骤要用新名字指代这一列。",
     "fields": [
       {"key": "column", "type": "column", "label": "要改的列", "required": true},
       {"key": "to", "type": "text", "label": "新表头", "required": true}
     ]},
    {"op": "set_cells", "label": "写入指定单元格", "group": "col",
     "desc": "向固定位置的单元格写入文字、数字或公式（如在 AC1 写表头、在 B2 写合计公式）。",
     "fields": [
       {"key": "cells", "type": "cells", "label": "单元格与内容", "required": true}
     ]},
    {"op": "number_format", "label": "设置数字格式", "group": "col",
     "desc": "设置列的显示格式，如千分位、两位小数、百分比、日期。",
     "fields": [
       {"key": "columns", "type": "columns", "label": "哪些列", "required": true},
       {"key": "format", "type": "numfmt", "label": "格式", "required": true},
       {"key": "where", "type": "condition", "label": "只处理符合条件的行"},
       {"key": "include_header", "type": "bool", "label": "表头单元格也设置", "default": false}
     ]},
    {"op": "set_width", "label": "设置列宽", "group": "col",
     "desc": "统一设置一列或多列的宽度。",
     "fields": [
       {"key": "columns", "type": "columns", "label": "哪些列", "required": true},
       {"key": "width", "type": "number", "label": "列宽", "required": true, "min": 0, "placeholder": "如：12"}
     ]},

    {"op": "delete_rows", "label": "删除行", "group": "row",
     "desc": "删除符合条件的行、指定行号范围，或整行为空的行。",
     "fields": [
       {"key": "mode", "type": "rowMode", "keys": ["where", "rows", "blank"], "label": "删除哪些行", "required": true}
     ]},
    {"op": "keep_rows", "label": "只保留符合条件的行", "group": "row",
     "desc": "不符合条件的数据行全部删除。",
     "fields": [
       {"key": "where", "type": "condition", "label": "保留条件", "required": true}
     ]},
    {"op": "insert_rows", "label": "插入空行", "group": "row",
     "desc": "在指定行之前或表格最后插入空行。",
     "fields": [
       {"key": "at", "type": "rowsAt", "keys": ["before_row", "at_end"], "label": "插入位置", "required": true},
       {"key": "count", "type": "number", "label": "插入几行", "default": 1, "min": 1}
     ]},
    {"op": "dedupe_rows", "label": "删除重复行", "group": "row",
     "desc": "按指定列判断重复，只保留第一次（或最后一次）出现的行。",
     "fields": [
       {"key": "keys", "type": "columns", "label": "按哪些列判断重复", "help": "不选=整行所有列都相同才算重复"},
       {"key": "keep", "type": "enum", "label": "保留", "default": "first",
        "options": [{"value": "first", "label": "第一次出现的行"}, {"value": "last", "label": "最后一次出现的行"}]}
     ]},
    {"op": "sort_rows", "label": "排序", "group": "row",
     "desc": "按一列或多列排序数据行（表头不动）。",
     "fields": [
       {"key": "by", "type": "sortBy", "label": "排序依据", "required": true}
     ]},

    {"op": "fill", "label": "填充 / 按条件写入", "group": "fill",
     "desc": "在某列写入固定值、公式或另一列的值，可只写符合条件的行（如“未分配数≥10 的写逾期”）。",
     "aliases": ["write_if"],
     "fields": [
       {"key": "column", "type": "column", "label": "写入哪一列", "required": true},
       {"key": "content", "type": "content", "keys": ["value", "formula", "formula_from", "from_col"], "label": "写入内容", "required": true},
       {"key": "where", "type": "condition", "label": "只写符合条件的行", "help": "不设置=所有数据行"},
       {"key": "rows", "type": "rowRange", "label": "只写这些行号", "placeholder": "如 5:20，不填=所有数据行"},
       {"key": "only_empty", "type": "bool", "label": "只写原本为空的单元格", "default": false}
     ]},
    {"op": "fill_down", "label": "向下填充", "group": "fill",
     "desc": "像 Excel 下拉一样，把某一行的公式/值填充到下面各行（相对引用自动平移）。",
     "fields": [
       {"key": "columns", "type": "columns", "label": "哪些列", "required": true},
       {"key": "from_row", "type": "number", "label": "从第几行复制", "placeholder": "不填=第一行数据", "min": 1},
       {"key": "to_row", "type": "number", "label": "填充到第几行", "placeholder": "不填=最后一行数据", "min": 1},
       {"key": "with_style", "type": "bool", "label": "连同格式一起填充", "default": true}
     ]},
    {"op": "replace_text", "label": "查找替换", "group": "fill",
     "desc": "在指定列里查找文字并替换。",
     "fields": [
       {"key": "columns", "type": "columns", "label": "在哪些列", "required": true},
       {"key": "find", "type": "text", "label": "查找", "required": true},
       {"key": "replace", "type": "text", "label": "替换为", "placeholder": "不填=删除找到的文字", "keepEmpty": true},
       {"key": "whole_cell", "type": "bool", "label": "整个单元格完全相同才替换", "default": false},
       {"key": "match_case", "type": "bool", "label": "区分大小写", "default": false},
       {"key": "regex", "type": "bool", "label": "按正则表达式查找", "default": false},
       {"key": "where", "type": "condition", "label": "只处理符合条件的行"}
     ]},
    {"op": "format", "label": "按条件设置格式（着色）", "group": "fill",
     "desc": "给符合条件的单元格加底色、改字体颜色、加粗或改数字格式。",
     "aliases": ["color"],
     "fields": [
       {"key": "columns", "type": "columns", "label": "给哪些列设置", "required": true},
       {"key": "where", "type": "condition", "label": "只处理符合条件的行", "help": "不设置=所有数据行"},
       {"key": "fill", "type": "color", "label": "底色"},
       {"key": "font_color", "type": "color", "label": "字体颜色"},
       {"key": "bold", "type": "triBool", "label": "加粗", "options": [{"value": "", "label": "不改"}, {"value": true, "label": "加粗"}, {"value": false, "label": "取消加粗"}]},
       {"key": "clear_fill", "type": "bool", "label": "清除原有底色（未设置底色时）", "default": false},
       {"key": "number_format", "type": "numfmt", "label": "数字格式"},
       {"key": "include_header", "type": "bool", "label": "表头也设置", "default": false}
     ]},

    {"op": "match", "label": "建立匹配（单键/多键）", "group": "match", "writes": false,
     "desc": "按一列或多列组合（如 订单号+商品编码）把主表每行和参照表对应起来，供后面的“按匹配复制”“比对”使用。",
     "fields": [
       {"key": "id", "type": "matchId", "label": "匹配名称", "required": true, "help": "后面的步骤用这个名字引用本次匹配"},
       {"key": "ref", "type": "fileRef", "label": "参照表", "required": true},
       {"key": "main_keys", "type": "columns", "label": "主表的键列（可多列组合）", "required": true},
       {"key": "ref_keys", "type": "columns", "source": "ref", "label": "参照表对应的键列", "help": "不选=与主表同名的列；多列时顺序要与主表一致"},
       {"key": "ref_sheet", "type": "sheet", "source": "ref", "label": "参照表的工作表", "placeholder": "不填=该文件设置的工作表"},
       {"key": "ref_header_row", "type": "number", "label": "参照表表头行", "placeholder": "不填=该文件设置的表头行", "min": 1},
       {"key": "trim", "type": "bool", "label": "忽略首尾空格", "default": true},
       {"key": "ignore_case", "type": "bool", "label": "忽略大小写", "default": false},
       {"key": "duplicates", "type": "enum", "label": "参照表键重复时", "default": "first",
        "options": [{"value": "first", "label": "取第一次出现的行"}, {"value": "last", "label": "取最后一次出现的行"}, {"value": "error", "label": "报错停止"}]},
       {"key": "ref_where", "type": "condition", "source": "ref", "label": "只用参照表中符合条件的行"}
     ]},
    {"op": "copy_matched", "label": "按匹配复制（值或公式）", "group": "match",
     "desc": "对匹配成功的行，把参照表指定列的值（或公式）复制到主表，可附加条件、复制后着色、标记未匹配。",
     "hidden": ["from", "to"],
     "fields": [
       {"key": "match", "type": "matchRef", "label": "使用哪个匹配", "required": true},
       {"key": "columns", "type": "copyColumns", "label": "复制哪些列", "required": true, "requiredKeys": ["columns", "from"]},
       {"key": "mode", "type": "enum", "label": "复制方式", "default": "value",
        "options": [{"value": "value", "label": "复制值"}, {"value": "formula", "label": "复制公式（相对引用自动平移）"}]},
       {"key": "where", "type": "condition", "label": "只复制符合条件的行", "allowRef": true},
       {"key": "skip_empty", "type": "bool", "label": "参照表为空时不覆盖", "default": true},
       {"key": "only_empty", "type": "bool", "label": "只写主表原本为空的单元格", "default": false},
       {"key": "create_missing", "type": "bool", "label": "主表没有目标列时自动加在最后", "default": false},
       {"key": "copy_format", "type": "bool", "label": "同时复制数字格式", "default": false},
       {"key": "fill", "type": "color", "label": "复制后给这些行着色"},
       {"key": "color_columns", "type": "columns", "label": "着色哪些列", "help": "不选=复制的目标列"},
       {"key": "not_found", "type": "notFound", "label": "未匹配到的行"}
     ]},
    {"op": "compare", "label": "比对两表差异", "group": "match",
     "desc": "逐行逐列比较主表与参照表：标出不一致的单元格，写比对结果列，并生成比对报告。",
     "fields": [
       {"key": "match", "type": "matchRef", "label": "使用哪个匹配", "required": true},
       {"key": "compare_columns", "type": "comparePairs", "label": "比较哪些列"},
       {"key": "tolerance", "type": "number", "label": "数值容差", "placeholder": "如 0.01，差值不超过它视为一致", "min": 0},
       {"key": "status_column", "type": "text", "label": "比对结果写到列", "placeholder": "比对结果"},
       {"key": "diff_column", "type": "optText", "label": "写出差异列名", "defaultText": "差异列"},
       {"key": "labels", "type": "labels", "label": "结果文字"},
       {"key": "highlight", "type": "color", "label": "差异单元格底色", "default": "FFEB9C", "nullable": true},
       {"key": "report_sheet", "type": "optText", "label": "在结果里新建比对报告表", "defaultText": "比对报告"},
       {"key": "export_csv", "type": "bool", "label": "另外导出差异明细 CSV", "default": false}
     ]},

    {"op": "append_rows", "label": "追加行（多文件合并）", "group": "merge",
     "desc": "把参照表的数据行接到主表末尾，按表头名称对应列（列顺序不同也没关系）。参照表设为“多个文件”时，会把所有符合规则的文件都接上（如各门店的表）。主表最后一行的公式和汇总范围会自动延续到新行。",
     "fields": [
       {"key": "from", "type": "sheetSource", "label": "从哪个文件追加", "required": true},
       {"key": "source_column", "type": "text", "label": "写入来源文件名的列", "placeholder": "如：来源文件（不填=不写）"},
       {"key": "add_missing_columns", "type": "bool", "label": "来源有、主表没有的列也加进来（加在最后）", "default": false},
       {"key": "extend_formulas", "type": "bool", "label": "来源里没有的列，按主表最后一行的公式补上（来源有值的列照抄来源的值）", "default": true},
       {"key": "skip_blank", "type": "bool", "label": "跳过整行为空的行", "default": true},
       {"key": "where", "type": "condition", "source": "from", "label": "只追加符合条件的行", "help": "条件里的列是来源文件的列"}
     ]},
    {"op": "group_summary", "label": "分组汇总", "group": "merge",
     "desc": "按一列或多列分组，计算合计、记录数、平均、最大、最小，生成一张新的汇总工作表。默认写成公式，数据改了打开文件就自动更新。",
     "fields": [
       {"key": "group_by", "type": "columns", "label": "按哪些列分组", "required": true},
       {"key": "metrics", "type": "metrics", "label": "汇总指标", "required": true},
       {"key": "sheet_name", "type": "text", "label": "汇总表名称", "placeholder": "分组汇总"},
       {"key": "mode", "type": "enum", "label": "写入方式", "default": "formula",
        "options": [{"value": "formula", "label": "公式（数据变了打开即更新）"}, {"value": "value", "label": "结果值（可用“不重复个数”和筛选条件）"}]},
       {"key": "sort", "type": "enum", "label": "排列顺序", "default": "key",
        "options": [{"value": "key", "label": "按分组列升序"}, {"value": "metric_desc", "label": "按第一个指标从大到小"}]},
       {"key": "total_row", "type": "bool", "label": "最后加一行合计", "default": true},
       {"key": "replace", "type": "bool", "label": "已有同名工作表时替换", "default": false},
       {"key": "where", "type": "condition", "label": "只汇总符合条件的行", "help": "只能配合“结果值”方式；用公式方式时，可先用“删除行”去掉不要的行"}
     ]},

    {"op": "copy_range", "label": "区块复制", "group": "copy",
     "desc": "把某个文件指定区域（含公式、底色、字体、边框、列宽、合并单元格）复制到主表指定位置。",
     "fields": [
       {"key": "from", "type": "rangeSource", "label": "复制来源", "required": true},
       {"key": "to", "type": "cellTarget", "label": "粘贴到主表", "required": true},
       {"key": "content", "type": "contentParts", "label": "复制哪些内容"},
       {"key": "formula_fallback", "type": "enum", "label": "公式引用了主表没有的工作表时", "default": "value",
        "options": [{"value": "value", "label": "改为复制计算结果"}, {"value": "error", "label": "报错停止"}]}
     ]},
    {"op": "copy_sheet", "label": "复制整个工作表", "group": "copy",
     "desc": "把参照表中的整张工作表复制进主表（如单价表），之后公式可以直接引用它。",
     "fields": [
       {"key": "from", "type": "sheetSource", "label": "复制哪张表", "required": true},
       {"key": "name", "type": "text", "label": "在主表里的名字", "placeholder": "不填=与原表同名"},
       {"key": "replace", "type": "bool", "label": "主表已有同名工作表时替换", "default": false},
       {"key": "content", "type": "contentParts", "label": "复制哪些内容"},
       {"key": "formula_fallback", "type": "enum", "label": "公式引用了主表没有的工作表时", "default": "value",
        "options": [{"value": "value", "label": "改为复制计算结果"}, {"value": "error", "label": "报错停止"}]}
     ]},

    {"op": "validate", "label": "校验", "group": "check", "writes": false,
     "desc": "检查数据（必需列、不能为空、不能重复、行数、表头一致、自定义条件），结果写进校验报告；可设置不通过就中止。",
     "fields": [
       {"key": "checks", "type": "checks", "label": "检查项", "required": true}
     ]}
  ]
};
