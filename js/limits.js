/* 处理上限与 AI 用量上限（集中放在这里，按需调整数字即可）
 *
 * 文件上限：保护浏览器不卡死，同时间接限制发给 AI 的表结构描述大小。
 * AI 上限：每次请求只发送“表结构描述 + 需求”，不发送整份数据；
 *          文件行数不影响费用，列数和工作表数才会影响，所以按“估算 token”来限。
 * 注意：这些限制在浏览器里执行，只能防止误用；不能当作安全边界（会看源码的人可以绕过）。
 */
(function (EFA) {
  'use strict';
  EFA.LIMITS = {
    file: {
      maxFileMB: 30,            // 单个文件大小上限
      maxSheets: 20,            // 一个文件最多载入的工作表数，超出的不载入
      maxRows: 100000,          // 每个工作表最多数据行数，超出的工作表不载入（不截断，避免算错）
      maxCols: 200,             // 每个工作表最多列数，超出的工作表不载入
      maxTotalCells: 3000000,   // 所有已载入工作表的单元格总数（行×列）
    },
    ai: {
      maxInputTokens: 6000,     // 单次请求发送内容的估算上限（含说明和表结构），超出时自动精简
      maxColsPerSheet: 80,      // 每个工作表最多向 AI 描述的列数
      maxOtherSheets: 5,        // 除当前表外，最多向 AI 描述的其他工作表数
      maxCallsPerRequest: 3,    // 一次“生成”最多调用 AI 的次数（含自动修正）
      maxOutputTokens: 2048,    // 普通模型单次最多输出
      maxOutputTokensThinking: 8192, // 思考型模型单次最多输出（含思考过程）
      defaultDailyCalls: 50,    // 默认每天最多调用次数（使用者可在“AI 设置”里修改，0 表示不限）
    },
  };
})(window.EFA = window.EFA || {});
