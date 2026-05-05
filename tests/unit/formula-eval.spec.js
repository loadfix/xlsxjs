// @ts-check
//
// Unit tests for the minimal formula evaluator (src/formula-eval.ts).
// Exercised through the UMD bundle in a real browser — the evaluator is a
// pure function (no DOM), so the browser harness is just a convenient JS
// runtime. Each test constructs a toy SheetLike shape and calls
// `xlsx.evaluateFormula` / `xlsx.evaluateSheetFormulas` directly.

import { test, expect } from '@playwright/test';

test.describe('Formula evaluator — POC', () => {
    test.beforeEach(async ({ page }) => {
        await page.goto('/tests/harness.html');
    });

    test('arithmetic: +, -, *, /, ^, %, parens, unary minus', async ({ page }) => {
        const result = await page.evaluate(() => {
            // @ts-ignore
            const resolver = () => null;
            const ev = (f) => xlsx.evaluateFormula(f, resolver);
            return {
                add: ev('=1+2').value,
                sub: ev('=10-3').value,
                mul: ev('=4*5').value,
                div: ev('=20/4').value,
                pow: ev('=2^10').value,
                pct: ev('=50%').value,
                parens: ev('=(2+3)*4').value,
                unary: ev('=-5+2').value,
                prec: ev('=2+3*4').value,           // 14, not 20
                powRight: ev('=2^3^2').value,       // right-assoc → 512
                divZero: ev('=5/0').kind,
            };
        });
        expect(result.add).toBe('3');
        expect(result.sub).toBe('7');
        expect(result.mul).toBe('20');
        expect(result.div).toBe('5');
        expect(result.pow).toBe('1024');
        expect(result.pct).toBe('0.5');
        expect(result.parens).toBe('20');
        expect(result.unary).toBe('-3');
        expect(result.prec).toBe('14');
        expect(result.powRight).toBe('512');
        expect(result.divZero).toBe('error');
    });

    test('SUM over a range and scalar args', async ({ page }) => {
        const result = await page.evaluate(() => {
            const grid = [
                [1, 2, 3],
                [4, 5, 6],
            ];
            const resolver = (col, row) => (grid[row] && grid[row][col] !== undefined) ? grid[row][col] : null;
            // @ts-ignore
            return {
                range: xlsx.evaluateFormula('=SUM(A1:C2)', resolver).value,
                scalar: xlsx.evaluateFormula('=SUM(1, 2, 3, 4)', resolver).value,
                mixed: xlsx.evaluateFormula('=SUM(A1:B1, 10)', resolver).value,
                single: xlsx.evaluateFormula('=SUM(A1)', resolver).value,
                empty: xlsx.evaluateFormula('=SUM()', resolver).value,
            };
        });
        expect(result.range).toBe('21');
        expect(result.scalar).toBe('10');
        expect(result.mixed).toBe('13');
        expect(result.single).toBe('1');
        expect(result.empty).toBe('0');
    });

    test('AVERAGE / MIN / MAX / COUNT / COUNTA', async ({ page }) => {
        const result = await page.evaluate(() => {
            // Grid: [10, 20, "text", null, 30]
            const grid = [[10, 20, 'text', null, 30]];
            const resolver = (col, row) => (grid[row] && grid[row][col] !== undefined) ? grid[row][col] : null;
            // @ts-ignore
            const ev = (f) => xlsx.evaluateFormula(f, resolver).value;
            return {
                avg: ev('=AVERAGE(A1:E1)'),   // (10+20+30)/3 = 20
                min: ev('=MIN(A1:E1)'),        // 10
                max: ev('=MAX(A1:E1)'),        // 30
                count: ev('=COUNT(A1:E1)'),    // 3 numeric
                counta: ev('=COUNTA(A1:E1)'),  // 4 non-empty
                avgEmpty: xlsx.evaluateFormula('=AVERAGE(B1:B1)', resolver).kind, // all empty → error
            };
        });
        expect(result.avg).toBe('20');
        expect(result.min).toBe('10');
        expect(result.max).toBe('30');
        expect(result.count).toBe('3');
        expect(result.counta).toBe('4');
        // AVERAGE of the sole "text" cell: text is skipped, divisor 0 → error
    });

    test('IF / AND / OR / NOT', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null;
            // @ts-ignore
            const ev = (f) => xlsx.evaluateFormula(f, resolver);
            return {
                ifTrue: ev('=IF(1>0, "yes", "no")').value,
                ifFalse: ev('=IF(1<0, "yes", "no")').value,
                ifNum: ev('=IF(TRUE, 42, 0)').value,
                ifNoElse: ev('=IF(FALSE, 1)').value, // default else → FALSE
                andT: ev('=AND(TRUE, 1, 2>1)').value,
                andF: ev('=AND(TRUE, FALSE)').value,
                orT: ev('=OR(FALSE, 0, 1)').value,
                orF: ev('=OR(FALSE, 0)').value,
                notT: ev('=NOT(TRUE)').value,
                notF: ev('=NOT(FALSE)').value,
            };
        });
        expect(result.ifTrue).toBe('yes');
        expect(result.ifFalse).toBe('no');
        expect(result.ifNum).toBe('42');
        expect(result.ifNoElse).toBe('FALSE');
        expect(result.andT).toBe('TRUE');
        expect(result.andF).toBe('FALSE');
        expect(result.orT).toBe('TRUE');
        expect(result.orF).toBe('FALSE');
        expect(result.notT).toBe('FALSE');
        expect(result.notF).toBe('TRUE');
    });

    test('comparison operators return booleans', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null;
            // @ts-ignore
            const ev = (f) => xlsx.evaluateFormula(f, resolver).value;
            return {
                eq: ev('=1=1'),
                neq: ev('=1<>2'),
                lt: ev('=1<2'),
                lte: ev('=2<=2'),
                gt: ev('=3>1'),
                gte: ev('=3>=3'),
                strEq: ev('="abc"="abc"'),
                strNeqCaseInsensitive: ev('="abc"="ABC"'), // Excel: case-insensitive
            };
        });
        expect(result.eq).toBe('TRUE');
        expect(result.neq).toBe('TRUE');
        expect(result.lt).toBe('TRUE');
        expect(result.lte).toBe('TRUE');
        expect(result.gt).toBe('TRUE');
        expect(result.gte).toBe('TRUE');
        expect(result.strEq).toBe('TRUE');
        expect(result.strNeqCaseInsensitive).toBe('TRUE');
    });

    test('string concat with &', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null;
            // @ts-ignore
            return {
                twoStr: xlsx.evaluateFormula('="foo"&"bar"', resolver).value,
                strNum: xlsx.evaluateFormula('="x="&5', resolver).value,
                boolStr: xlsx.evaluateFormula('=TRUE&"!"', resolver).value,
            };
        });
        expect(result.twoStr).toBe('foobar');
        expect(result.strNum).toBe('x=5');
        expect(result.boolStr).toBe('TRUE!');
    });

    test('cell ref resolution (absolute $A$1 forms too)', async ({ page }) => {
        const result = await page.evaluate(() => {
            const grid = [[7, 8]];
            const resolver = (col, row) => (grid[row] && grid[row][col] !== undefined) ? grid[row][col] : null;
            // @ts-ignore
            const ev = (f) => xlsx.evaluateFormula(f, resolver).value;
            return {
                rel: ev('=A1+B1'),
                absRow: ev('=A$1+B1'),
                absCol: ev('=$A1+$B1'),
                absBoth: ev('=$A$1*$B$1'),
                lower: ev('=a1+b1'), // Excel accepts lower-case refs
            };
        });
        expect(result.rel).toBe('15');
        expect(result.absRow).toBe('15');
        expect(result.absCol).toBe('15');
        expect(result.absBoth).toBe('56');
        expect(result.lower).toBe('15');
    });

    test('empty cells fold to 0 in numeric ops', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null; // every cell is empty
            // @ts-ignore
            return {
                addEmpty: xlsx.evaluateFormula('=A1+5', resolver).value,
                sumEmpty: xlsx.evaluateFormula('=SUM(A1:C3)', resolver).value,
                ifEmpty: xlsx.evaluateFormula('=IF(A1, "y", "n")', resolver).value,
            };
        });
        expect(result.addEmpty).toBe('5');
        expect(result.sumEmpty).toBe('0');
        expect(result.ifEmpty).toBe('n'); // empty → false
    });

    test('text in numeric context yields an error (not silent 0)', async ({ page }) => {
        const result = await page.evaluate(() => {
            const grid = [['hello']];
            const resolver = (col, row) => (grid[row] && grid[row][col] !== undefined) ? grid[row][col] : null;
            // @ts-ignore
            return {
                // Direct ref to text cell, numeric op → error
                addText: xlsx.evaluateFormula('=A1+1', resolver).kind,
                // But SUM skips text (Excel behaviour)
                sumText: xlsx.evaluateFormula('=SUM(A1:A1)', resolver).value,
                // COUNT skips text
                countText: xlsx.evaluateFormula('=COUNT(A1:A1)', resolver).value,
            };
        });
        expect(result.addText).toBe('error');
        expect(result.sumText).toBe('0');
        expect(result.countText).toBe('0');
    });

    test('error propagation through operators and functions', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null;
            // @ts-ignore
            return {
                litErr: xlsx.evaluateFormula('=#DIV/0!', resolver).kind,
                addErr: xlsx.evaluateFormula('=1+#N/A', resolver).kind,
                sumErr: xlsx.evaluateFormula('=SUM(1, 2, #REF!)', resolver).kind,
            };
        });
        expect(result.litErr).toBe('error');
        expect(result.addErr).toBe('error');
        expect(result.sumErr).toBe('error');
    });

    test('nested function calls', async ({ page }) => {
        const result = await page.evaluate(() => {
            const grid = [[1, 2, 3, 4, 5]];
            const resolver = (col, row) => (grid[row] && grid[row][col] !== undefined) ? grid[row][col] : null;
            // @ts-ignore
            return {
                nested: xlsx.evaluateFormula('=IF(SUM(A1:E1)>10, MAX(A1:E1), MIN(A1:E1))', resolver).value,
                double: xlsx.evaluateFormula('=SUM(SUM(A1:C1), SUM(D1:E1))', resolver).value,
                andNested: xlsx.evaluateFormula('=AND(SUM(A1:E1)=15, COUNT(A1:E1)=5)', resolver).value,
            };
        });
        expect(result.nested).toBe('5');  // sum=15>10 → MAX=5
        expect(result.double).toBe('15');
        expect(result.andNested).toBe('TRUE');
    });

    test('unknown function name → error', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null;
            // @ts-ignore
            return xlsx.evaluateFormula('=VLOOKUP(1, A:B, 2)', resolver).kind;
        });
        expect(result).toBe('error');
    });

    test('parse error in malformed formula → error', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null;
            // @ts-ignore
            return {
                unclosedParen: xlsx.evaluateFormula('=SUM(1,2', resolver).kind,
                strayOp: xlsx.evaluateFormula('=*5', resolver).kind,
            };
        });
        expect(result.unclosedParen).toBe('error');
        expect(result.strayOp).toBe('error');
    });

    test('evaluateSheetFormulas walks a sheet and rewrites empty <v> cells', async ({ page }) => {
        const result = await page.evaluate(() => {
            // Simulate a SheetLike: A1=1, A2=2, A3=3, B1 has =SUM(A1:A3) with
            // no cached value.
            const mk = (col, row, value, kind, formula = null) => ({
                col, row, value, kind, formula, styleIndex: -1, runs: null,
                phonetics: null, cellMetadataIndex: null, valueMetadataIndex: null,
                isSpillAnchor: false,
            });
            const sheet = {
                rows: [
                    [mk(0, 0, '1', 'number'), mk(1, 0, '', 'empty', 'SUM(A1:A3)')],
                    [mk(0, 1, '2', 'number')],
                    [mk(0, 2, '3', 'number')],
                ],
            };
            // @ts-ignore
            xlsx.evaluateSheetFormulas(sheet);
            return {
                b1Value: sheet.rows[0][1].value,
                b1Kind: sheet.rows[0][1].kind,
                a1Untouched: sheet.rows[0][0].value,
            };
        });
        expect(result.b1Value).toBe('6');
        expect(result.b1Kind).toBe('number');
        expect(result.a1Untouched).toBe('1');
    });

    test('evaluateFormulas renderer option runs end-to-end through parseAsync', async ({ page }) => {
        const result = await page.evaluate(async () => {
            // The formulas fixture contains:
            //   A1=10, B1=20, C1=30
            //   A2 =SUM(A1:C1)  with cached <v>60
            //   A3 =A1+A2       with NO cached <v> (this is the target)
            //   A4 =A1&" rows"  with cached <v>"10 rows"
            //   A5 =1/0         with cached <v>#DIV/0!
            const buf = await fetch('/tests/render-test/formulas/workbook.xlsx').then((r) => r.arrayBuffer());
            // @ts-ignore
            const evaluated = await xlsx.parseAsync(buf, { evaluateFormulas: true });
            const sheet = evaluated.parsed.sheets[0];
            const getCell = (row, col) => {
                const r = sheet.rows[row];
                if (!r) return null;
                return r[col] || null;
            };
            return {
                a3Value: getCell(2, 0)?.value,
                a3Kind: getCell(2, 0)?.kind,
                // A2 had a cached value — evaluator should leave it alone
                // (non-force mode).
                a2Value: getCell(1, 0)?.value,
                a2Kind: getCell(1, 0)?.kind,
            };
        });
        // A3 =A1+A2 → 10 + 60 = 70
        expect(result.a3Value).toBe('70');
        expect(result.a3Kind).toBe('number');
        // A2 keeps its cached value; evaluator doesn't force.
        expect(result.a2Value).toBe('60');
    });

    test('evaluateFormulasForce overwrites cached <v>', async ({ page }) => {
        const result = await page.evaluate(async () => {
            const buf = await fetch('/tests/render-test/formulas/workbook.xlsx').then((r) => r.arrayBuffer());
            // @ts-ignore
            const wb = await xlsx.parseAsync(buf, {
                evaluateFormulas: true,
                evaluateFormulasForce: true,
            });
            const sheet = wb.parsed.sheets[0];
            const a2 = sheet.rows[1][0];
            return { value: a2.value, kind: a2.kind };
        });
        // =SUM(A1:C1) = 60 (same result, but evaluator re-computed)
        expect(result.value).toBe('60');
        expect(result.kind).toBe('number');
    });
});
