/* 由 build.py 根据 executor/templates/*.json 自动生成，请勿手改（改示例请改 executor/templates）。 */
(window.EFA = window.EFA || {}).WF_EXAMPLES = [
 {
  "schema": "excel-workflow/1",
  "name": "示例：两表比对",
  "description": "按“订单号+商品编码”比对本期与上期，标出金额、状态不一致的单元格，并生成比对报告。",
  "files": [
   {
    "id": "main",
    "role": "main",
    "label": "本期订单",
    "match": "*本期*.xlsx"
   },
   {
    "id": "ref",
    "role": "ref",
    "label": "上期订单",
    "match": "*上期*.xlsx"
   }
  ],
  "output": {
   "file_name": "{main}_比对结果_{time}.xlsx"
  },
  "options": {
   "recalc": "auto"
  },
  "steps": [
   {
    "op": "match",
    "title": "按订单号+商品编码匹配上期",
    "id": "上期",
    "ref": "ref",
    "main_keys": [
     "订单号",
     "商品编码"
    ]
   },
   {
    "op": "compare",
    "title": "比对金额和状态",
    "match": "上期",
    "compare_columns": [
     "金额",
     "状态"
    ],
    "tolerance": 0.01,
    "highlight": "FFEB9C",
    "report_sheet": "比对报告",
    "export_csv": true
   },
   {
    "op": "format",
    "title": "仅本期有的订单标绿",
    "columns": "A:B",
    "where": {
     "col": "比对结果",
     "op": "=",
     "value": "仅主表有"
    },
    "fill": "C6EFCE"
   }
  ]
 },
 {
  "schema": "excel-workflow/1",
  "name": "示例：按多列匹配回填并标色",
  "description": "用“订单号+商品编码”两列组合匹配上期表；未分配数≥10 的行回填上期的下单/入库时间并标色，未分配数=数量的写“逾期”。",
  "files": [
   {
    "id": "main",
    "role": "main",
    "label": "本期订单",
    "match": "*本期*.xlsx"
   },
   {
    "id": "ref",
    "role": "ref",
    "label": "上期订单",
    "match": "*上期*.xlsx"
   }
  ],
  "options": {
   "recalc": "auto"
  },
  "steps": [
   {
    "op": "insert_column",
    "title": "新增备注列",
    "header": "备注",
    "end": true
   },
   {
    "op": "format",
    "title": "未分配数≥10 标色",
    "columns": [
     "未分配数"
    ],
    "where": {
     "col": "未分配数",
     "op": ">=",
     "value": 10
    },
    "fill": "FCE4D6"
   },
   {
    "op": "match",
    "title": "订单号+商品编码匹配上期",
    "id": "m",
    "ref": "ref",
    "main_keys": [
     "订单号",
     "商品编码"
    ]
   },
   {
    "op": "copy_matched",
    "title": "回填下单/入库时间",
    "match": "m",
    "columns": [
     "下单时间",
     "入库时间"
    ],
    "create_missing": true,
    "where": {
     "col": "未分配数",
     "op": ">=",
     "value": 10
    },
    "fill": "FCE4D6"
   },
   {
    "op": "fill",
    "title": "写逾期",
    "column": "备注",
    "value": "逾期",
    "where": {
     "all": [
      {
       "col": "未分配数",
       "op": ">=",
       "value": 10
      },
      {
       "col": "未分配数",
       "op": "=",
       "col2": "数量"
      }
     ]
    }
   },
   {
    "op": "validate",
    "title": "检查",
    "checks": [
     {
      "type": "unique",
      "columns": [
       "订单号",
       "商品编码"
      ]
     },
     {
      "type": "not_empty",
      "columns": [
       "订单号"
      ],
      "on_fail": "stop"
     }
    ]
   }
  ]
 },
 {
  "schema": "excel-workflow/1",
  "name": "示例：整理明细表",
  "description": "删掉无用列、去掉测试数据和重复行、加一列含税金额公式、按金额从高到低排序。",
  "files": [
   {
    "id": "main",
    "role": "main",
    "label": "本期订单",
    "match": "*本期*.xlsx"
   }
  ],
  "options": {
   "recalc": "auto"
  },
  "steps": [
   {
    "op": "delete_columns",
    "title": "删除内部字段",
    "columns": [
     "内部备注"
    ]
   },
   {
    "op": "delete_rows",
    "title": "删除测试数据",
    "where": {
     "col": "客户",
     "op": "contains",
     "value": "测试"
    }
   },
   {
    "op": "dedupe_rows",
    "title": "按订单号+商品编码去重",
    "keys": [
     "订单号",
     "商品编码"
    ]
   },
   {
    "op": "insert_column",
    "title": "含税金额",
    "header": "含税金额",
    "after": "金额",
    "formula": "=ROUND([@金额]*1.13,2)"
   },
   {
    "op": "number_format",
    "title": "金额两位小数",
    "columns": [
     "金额",
     "含税金额"
    ],
    "format": "#,##0.00"
   },
   {
    "op": "sort_rows",
    "title": "按金额降序",
    "by": [
     {
      "column": "金额",
      "order": "desc"
     }
    ]
   },
   {
    "op": "rename_header",
    "title": "改表头",
    "column": "状态",
    "to": "订单状态"
   }
  ]
 }
];
