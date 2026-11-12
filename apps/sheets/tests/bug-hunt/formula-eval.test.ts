/**
 * formula-eval 猎捕（F3 牙试验，~20 题）：
 * - R1 真实公式画像 1:1：16× SUMPRODUCT((col="x")*col)、SUMIF 单位条件
 *   （"m³"）、COUNTIF 通配（"*DN300*"）、COUNTIFS "<>" / ""、SUM、
 *   COUNTA 整列、裸跨表 =汇总!C74——全部断言精确数值
 * - 错误分支：不支持函数（须含 aggregate_range 引导）、未知 sheet、
 *   整行引用、COUNTIFS 奇数参数（须含 COUNTIF 引导）、非函数表达式、空公式、
 *   SUMIF 两范围尺寸不一
 * - matchCriteria 表驱动："" / "<>" / 比较运算符 / 通配（大小写不敏感）/ 普通相等
 * - F13（R2 后）：R2 真实 46 错的 idiom 1:1——ISNUMBER(SEARCH()) 掩码（含 -- 前缀、
 *   逗号参两种形态）、数字字面量因子、ROW() 行号因子、MAX/MIN/COUNTBLANK/MATCH/LARGE、
 *   嵌套因子错误文案须点名嵌套函数并给 filterColumn 回退
 * 断言对象是求值器本身（确定性行为），删/坏实现必红。
 */
import { describe, expect, it } from 'vitest'

import {
  evaluateDemoFormula,
  matchCriteria,
  type DemoCellMap,
  type DemoCellValue,
  type DemoFormulaResult,
} from './formula-eval'

const sheet = (m: Record<string, DemoCellValue>): DemoCellMap => {
  const out: Record<string, { value: DemoCellValue }> = {}
  for (const [a, v] of Object.entries(m)) out[a] = { value: v }
  return out
}

const main = sheet({
  A2: 'x',
  A3: '',
  A5: 'b',
  B2: 'm³',
  B3: 'm²',
  C2: 'DN300-A',
  C3: 'DN300-B',
  C4: 'DN200',
  D2: 'DN300',
  D3: 'DN200',
  F2: 10,
  F3: 20,
})
const summary = sheet({ C74: 42 })
const resolve = (key?: string) =>
  key === undefined || key === '主表' ? main : key === '汇总' ? summary : undefined

const eval1 = (formula: string): DemoFormulaResult => evaluateDemoFormula(formula, resolve)

const err = (r: DemoFormulaResult): string => {
  if (typeof r === 'object' && r !== null && 'error' in r) return r.error
  throw new Error(`expected an error result, got ${JSON.stringify(r)}`)
}

describe('formula-eval：R1 真实公式 1:1', () => {
  it('SUMPRODUCT 条件乘积（R1 最高频形态）', () => {
    expect(eval1('=SUMPRODUCT((D2:D3="DN300")*F2:F3)')).toBe(10)
  })
  it('SUMPRODUCT <> 因子 = NOT 语义', () => {
    expect(eval1('=SUMPRODUCT((D2:D3<>"DN300")*F2:F3)')).toBe(20)
  })
  it('SUMPRODUCT 数值比较因子', () => {
    expect(eval1('=SUMPRODUCT((F2:F3>15)*F2:F3)')).toBe(20)
  })
  it('SUMIF 单位条件 + sum_range', () => {
    // 仅 B2='m³' 命中 → F2=10
    expect(eval1('=SUMIF(B2:B3,"m³",F2:F3)')).toBe(10)
  })
  it('COUNTIF 通配 *DN300*', () => {
    expect(eval1('=COUNTIF(C2:C4,"*DN300*")')).toBe(2)
  })
  it('COUNTIFS 双 "<>"（非空交集）', () => {
    // 行2 A=x B=m³ 均非空；行3 A=''；行4 双缺；行5 B 缺 → 仅行 2
    expect(eval1('=COUNTIFS(A2:A5,"<>",B2:B5,"<>")')).toBe(1)
  })
  it('COUNTIFS "" 条件（空值或空串）', () => {
    // A3='' 与 A4 缺失都算空
    expect(eval1('=COUNTIFS(A2:A5,"")')).toBe(2)
  })
  it('SUM 小范围', () => {
    expect(eval1('=SUM(F2:F3)')).toBe(30)
  })
  it('COUNTA 整列（sparse 枚举）', () => {
    // A3 是空串，COUNTA 不计
    expect(eval1('=COUNTA(A:A)')).toBe(2)
  })
  it('COUNT / AVERAGE', () => {
    expect(eval1('=COUNT(F2:F3)')).toBe(2)
    expect(eval1('=AVERAGE(F2:F3)')).toBe(15)
  })
  it('裸跨表引用 =汇总!C74', () => {
    expect(eval1('=汇总!C74')).toBe(42)
  })
  it('裸单元格引用（空格返回 null）', () => {
    expect(eval1('=A1')).toBe(null)
  })
})

describe('formula-eval：错误分支（消息须带回退引导，不许静默错值）', () => {
  it('不支持的函数 → 点名函数 + aggregate_range 引导', () => {
    const e = err(eval1('=VLOOKUP(A2,B2:C3,2)'))
    expect(e).toContain('VLOOKUP')
    expect(e).toContain('aggregate_range')
  })
  it('未知 sheet', () => {
    expect(err(eval1('=不存在!A1'))).toContain('unknown sheet "不存在"')
  })
  it('整行引用不支持', () => {
    expect(err(eval1('=SUM(2:5)'))).toContain('whole rows')
  })
  it('COUNTIFS 奇数参数 → 提示改用 COUNTIF', () => {
    // 3 个参数（range, crit, range）才是奇数；2 参数合法（等价 COUNTIF）
    const e = err(eval1('=COUNTIFS(A2:A4,"x",B2:B4)'))
    expect(e).toContain('odd count')
    expect(e).toContain('COUNTIF')
  })
  it('非函数表达式', () => {
    // R12 ①：二元算式给诚实文案（原 "cannot parse" 误导模型换形态重试，
    // R10 T15 ×2 / R12 T3 ×2 无效调用）
    expect(err(eval1('=A1+B1'))).toContain('binary arithmetic')
  })
  it('空公式', () => {
    expect(err(eval1('='))).toContain('empty formula')
  })
  it('SUMIF 条件范围与求和范围尺寸不一', () => {
    expect(err(eval1('=SUMIF(B2:B3,"m³",F2:F4)'))).toContain('different sizes')
  })
})

describe('formula-eval：F13 R2 真实 idiom（stand-in 对齐 Univer 能力）', () => {
  it('ISNUMBER(SEARCH("text",range)) 作 contains 掩码（R2 T14：B 含 m³ 的 F 求和）', () => {
    // B2='m³' 命中 → F2=10
    expect(eval1('=SUMPRODUCT(ISNUMBER(SEARCH("m³",B2:B3))*F2:F3)')).toBe(10)
  })
  it('-- 前缀 + 逗号参形态（R2 T14 的两种真实写法都要过）', () => {
    expect(eval1('=SUMPRODUCT(--ISNUMBER(SEARCH("m³",B2:B3)),F2:F3)')).toBe(10)
  })
  it('(range="")*1 数字字面量因子广播（R2 T7）', () => {
    // A3='' 与 A4 缺失 → 2
    expect(eval1('=SUMPRODUCT((A2:A5="")*1)')).toBe(2)
  })
  it('(range="")*ROW(range) 空值行号求和（R2 T7）', () => {
    // 行 3、4 空 → 3+4=7
    expect(eval1('=SUMPRODUCT((A2:A5="")*ROW(A2:A5))')).toBe(7)
  })
  it('MAX / MIN / COUNTBLANK', () => {
    expect(eval1('=MAX(F2:F3)')).toBe(20)
    expect(eval1('=MIN(F2:F3)')).toBe(10)
    expect(eval1('=COUNTBLANK(A2:A5)')).toBe(2)
  })
  it('MATCH 精确匹配返回 1-based 位置；未命中 → error', () => {
    expect(eval1('=MATCH("DN300",D2:D3,0)')).toBe(1)
    expect(err(eval1('=MATCH("nope",D2:D3,0)'))).toContain('not found')
  })
  it('LARGE(range,k)（k=1 同 MAX，k 越界 → error）', () => {
    expect(eval1('=LARGE(F2:F3,1)')).toBe(20)
    expect(err(eval1('=LARGE(F2:F3,3)'))).toContain('exceeds')
  })
  it('不支持的嵌套因子 → 点名嵌套函数 + filterColumn 引导（不再是 "cannot parse cell reference"）', () => {
    const e = err(eval1('=SUMPRODUCT(VALUE(F2:F3))'))
    expect(e).toContain('nested function')
    expect(e).toContain('filterColumn')
    expect(e).not.toContain('cannot parse cell reference')
  })
})

describe('matchCriteria：Excel 风格条件匹配', () => {
  const cases: Array<[DemoCellValue, string, boolean]> = [
    [null, '""', true],
    ['', '""', true],
    [5, '""', false],
    [null, '<>', false],
    [7, '<>', true],
    [10, '<=10', true],
    [11, '<=10', false],
    [15, '>=15', true],
    ['x', '<=10', false], // 非数字 → 比较恒 false
    ['DN300-A', '*dn300*', true], // 通配 + 大小写不敏感
    ['abc', 'a?c', true],
    ['ab', 'a?c', false],
    ['ABC', 'abc', true], // 普通相等大小写不敏感
    [10, '10', true], // 数字与数字串相等
  ]
  for (const [value, crit, want] of cases) {
    it(`matchCriteria(${JSON.stringify(value)}, ${JSON.stringify(crit)}) === ${want}`, () => {
      expect(matchCriteria(value, crit)).toBe(want)
    })
  }
})

describe('formula-eval: F21 大范围不超 V8 变参上限（R5 T46）', () => {
  // R5 T46：=COUNTA(A1:Z10000) 返回 "Maximum call stack size exceeded"——
  // 根因是 push(...valuesFor(...)) 变参展开超 V8 65,535 上限（V8 把它报成
  // 栈溢出）。牙试验用 A1:Z5000：展开 130,000 格 > 65,535，旧实现必返回
  // error，修复后返回真值。数据：A1:G2313 共 16,191 个数字格，值=行号。
  const bigCells: Record<string, DemoCellValue> = {}
  for (let r = 1; r <= 2313; r++) {
    for (let c = 0; c < 7; c++) bigCells[String.fromCharCode(65 + c) + r] = r
  }
  const big = sheet(bigCells)
  const evalBig = (formula: string): DemoFormulaResult =>
    evaluateDemoFormula(formula, () => big)

  it('=COUNTA(A1:Z5000) 130K 展开格返回真计数', () => {
    expect(evalBig('=COUNTA(A1:Z5000)')).toBe(2313 * 7)
  })
  it('=SUM(A1:Z5000) 同分支（SUM 与 COUNTA 共享收集循环）', () => {
    // 每列 1..2313 之和 × 7 列
    expect(evalBig('=SUM(A1:Z5000)')).toBe(7 * (2313 * 2314) / 2)
  })
  it('=MAX / =MIN(A1:Z5000) 第二处 spread', () => {
    expect(evalBig('=MAX(A1:Z5000)')).toBe(2313)
    expect(evalBig('=MIN(A1:Z5000)')).toBe(1)
  })
  it('未超限范围（A1:G5000，35K 格）行为不变', () => {
    expect(evalBig('=COUNTA(A1:G5000)')).toBe(2313 * 7)
  })
})

describe('formula-eval: R12 ① 顶层二元/拼接/MATCH 函数 lookup 诚实文案（R9–R12 误导性文案实锤）', () => {
  const m = sheet({ F2: 1, F3: 2, F4: 3 })
  const ev = (formula: string): DemoFormulaResult => evaluateDemoFormula(formula, () => m)
  const err = (r: DemoFormulaResult): string => (r as { error: string }).error

  it('顶层二元算式 → binary arithmetic + 分开调用引导（原 "cannot parse range" 误导，R10 T15 ×2 / R12 T3 ×2）', () => {
    const e = err(ev('=SUM(F2:F4)-SUM(F2:F3)'))
    expect(e).toContain('binary arithmetic')
    expect(e).toContain('evaluate_formula separately')
  })
  it('顶层 & 拼接 → string concatenation（含 LARGE 拼接实锤形态；原 "k must be a positive integer" 误导，R11 T1 ×5 / R12 T4 ×2）', () => {
    expect(err(ev('=F2&","&F3'))).toContain('string concatenation')
    expect(err(ev('=LARGE(F2:F4,1)&","&LARGE(F2:F4,2)'))).toContain('string concatenation')
  })
  it('MATCH 函数 lookup → lookup is a formula + find_cells/LARGE 引导（原 "not found (exact match only)" 数据级误导，R9/R10 T12 ×10）', () => {
    expect(err(ev('=MATCH(LARGE(F2:F4,1),F:F,0)'))).toContain('lookup value is a formula')
    expect(err(ev('=MATCH(LARGE(F2:F4,1),F:F,0)'))).toContain('find_cells')
  })
})
