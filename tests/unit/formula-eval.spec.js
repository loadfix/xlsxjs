// @ts-check
//
// Unit tests for the formula evaluator (src/formula-eval.ts).
// Exercised through the UMD bundle in a real browser — the evaluator is a
// pure function (no DOM), so the browser harness is just a convenient JS
// runtime. Each test constructs a toy SheetLike shape and calls
// `xlsx.evaluateFormula` / `xlsx.evaluateSheetFormulas` directly.

import { test, expect } from '@playwright/test';

test.describe('Formula evaluator — v2', () => {
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
                divZero: ev('=5/0').value,
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
        // v2: specific #DIV/0! rather than a generic #ERROR! sentinel.
        expect(result.divZero).toBe('#DIV/0!');
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
                // AVERAGE of the sole "text" cell (col C): text is skipped,
                // divisor 0 → #DIV/0!
                avgTextOnly: xlsx.evaluateFormula('=AVERAGE(C1:C1)', resolver).value,
            };
        });
        expect(result.avg).toBe('20');
        expect(result.min).toBe('10');
        expect(result.max).toBe('30');
        expect(result.count).toBe('3');
        expect(result.counta).toBe('4');
        expect(result.avgTextOnly).toBe('#DIV/0!');
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

    test('text in numeric context yields #VALUE! (not silent 0)', async ({ page }) => {
        const result = await page.evaluate(() => {
            const grid = [['hello']];
            const resolver = (col, row) => (grid[row] && grid[row][col] !== undefined) ? grid[row][col] : null;
            // @ts-ignore
            return {
                // Direct ref to text cell, numeric op → #VALUE!
                addText: xlsx.evaluateFormula('=A1+1', resolver).value,
                // But SUM skips text (Excel behaviour)
                sumText: xlsx.evaluateFormula('=SUM(A1:A1)', resolver).value,
                // COUNT skips text
                countText: xlsx.evaluateFormula('=COUNT(A1:A1)', resolver).value,
            };
        });
        expect(result.addText).toBe('#VALUE!');
        expect(result.sumText).toBe('0');
        expect(result.countText).toBe('0');
    });

    test('error propagation through operators and functions', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null;
            // @ts-ignore
            return {
                litErr: xlsx.evaluateFormula('=#DIV/0!', resolver).value,
                addErr: xlsx.evaluateFormula('=1+#N/A', resolver).value,
                sumErr: xlsx.evaluateFormula('=SUM(1, 2, #REF!)', resolver).value,
            };
        });
        // v2: the specific error code propagates, not a generic sentinel.
        expect(result.litErr).toBe('#DIV/0!');
        expect(result.addErr).toBe('#N/A');
        expect(result.sumErr).toBe('#REF!');
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

    test('unknown function name → #NAME?', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null;
            // @ts-ignore
            return xlsx.evaluateFormula('=BOGUSFUNC(1, 2)', resolver).value;
        });
        expect(result).toBe('#NAME?');
    });

    test('parse error in malformed formula → #VALUE!', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null;
            // @ts-ignore
            return {
                unclosedParen: xlsx.evaluateFormula('=SUM(1,2', resolver).value,
                strayOp: xlsx.evaluateFormula('=*5', resolver).value,
            };
        });
        expect(result.unclosedParen).toBe('#VALUE!');
        expect(result.strayOp).toBe('#VALUE!');
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
                return r.find((c) => c.col === col) || null;
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
            const a2 = sheet.rows[1].find((c) => c.col === 0);
            return { value: a2.value, kind: a2.kind };
        });
        // =SUM(A1:C1) = 60 (same result, but evaluator re-computed)
        expect(result.value).toBe('60');
        expect(result.kind).toBe('number');
    });

    // ── v2: new functions ─────────────────────────────────────────────────

    test('IFERROR / IFNA catch errors and substitute a fallback', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null;
            const ev = (f) => xlsx.evaluateFormula(f, resolver).value;
            return {
                iferrDiv: ev('=IFERROR(1/0, "oops")'),
                iferrPass: ev('=IFERROR(1+2, 999)'),
                ifnaHit: ev('=IFNA(#N/A, "nope")'),
                ifnaMiss: ev('=IFNA(#DIV/0!, "nope")'), // only #N/A caught
            };
        });
        expect(result.iferrDiv).toBe('oops');
        expect(result.iferrPass).toBe('3');
        expect(result.ifnaHit).toBe('nope');
        expect(result.ifnaMiss).toBe('#DIV/0!');
    });

    test('SWITCH picks first matching arm, else default', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null;
            const ev = (f) => xlsx.evaluateFormula(f, resolver).value;
            return {
                hit: ev('=SWITCH(2, 1, "one", 2, "two", 3, "three")'),
                miss: ev('=SWITCH(5, 1, "one", 2, "two", "other")'),
                noDefault: ev('=SWITCH(5, 1, "one", 2, "two")'),
            };
        });
        expect(result.hit).toBe('two');
        expect(result.miss).toBe('other');
        expect(result.noDefault).toBe('#N/A');
    });

    test('IFS returns first true arm', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null;
            const ev = (f) => xlsx.evaluateFormula(f, resolver).value;
            return {
                first: ev('=IFS(1=1, "yes", 2=2, "also")'),
                second: ev('=IFS(1=2, "no", 2=2, "yes")'),
                none: ev('=IFS(1=2, "no", 3=4, "still no")'),
            };
        });
        expect(result.first).toBe('yes');
        expect(result.second).toBe('yes');
        expect(result.none).toBe('#N/A');
    });

    test('ROUND / ROUNDUP / ROUNDDOWN', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null;
            const ev = (f) => xlsx.evaluateFormula(f, resolver).value;
            return {
                r2: ev('=ROUND(1.2345, 2)'),
                r0: ev('=ROUND(1.5, 0)'),
                rNeg: ev('=ROUND(1234.5, -2)'),  // round to nearest 100
                up: ev('=ROUNDUP(1.23, 1)'),
                down: ev('=ROUNDDOWN(1.29, 1)'),
                upNeg: ev('=ROUNDUP(-1.23, 1)'), // away from zero
            };
        });
        expect(result.r2).toBe('1.23');
        expect(result.r0).toBe('2');
        expect(result.rNeg).toBe('1200');
        expect(result.up).toBe('1.3');
        expect(result.down).toBe('1.2');
        expect(result.upNeg).toBe('-1.3');
    });

    test('ABS / SQRT / POWER / MOD / INT', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null;
            const ev = (f) => xlsx.evaluateFormula(f, resolver).value;
            return {
                abs: ev('=ABS(-5.5)'),
                sqrt: ev('=SQRT(16)'),
                sqrtNeg: ev('=SQRT(-1)'), // #NUM!
                pow: ev('=POWER(3, 4)'),
                mod: ev('=MOD(10, 3)'),
                modZero: ev('=MOD(5, 0)'), // #DIV/0!
                intPos: ev('=INT(3.7)'),
                intNeg: ev('=INT(-3.7)'),  // -4 per Excel INT = floor
            };
        });
        expect(result.abs).toBe('5.5');
        expect(result.sqrt).toBe('4');
        expect(result.sqrtNeg).toBe('#NUM!');
        expect(result.pow).toBe('81');
        expect(result.mod).toBe('1');
        expect(result.modZero).toBe('#DIV/0!');
        expect(result.intPos).toBe('3');
        expect(result.intNeg).toBe('-4');
    });

    test('SUMIF / COUNTIF with criteria strings', async ({ page }) => {
        const result = await page.evaluate(() => {
            // A: [10, 20, 30, 5, 15]
            // B: ["foo", "bar", "foo", "baz", "foo"]
            const a = [10, 20, 30, 5, 15];
            const b = ['foo', 'bar', 'foo', 'baz', 'foo'];
            const resolver = (col, row) => {
                if (col === 0) return a[row] ?? null;
                if (col === 1) return b[row] ?? null;
                return null;
            };
            const ev = (f) => xlsx.evaluateFormula(f, resolver).value;
            return {
                sumGt: ev('=SUMIF(A1:A5, ">=15")'),          // 20+30+15 = 65
                sumStr: ev('=SUMIF(B1:B5, "foo", A1:A5)'),   // 10+30+15 = 55
                cntGt: ev('=COUNTIF(A1:A5, ">10")'),         // 20,30,15 = 3
                cntStr: ev('=COUNTIF(B1:B5, "foo")'),        // 3
                cntNe: ev('=COUNTIF(B1:B5, "<>foo")'),       // 2
            };
        });
        expect(result.sumGt).toBe('65');
        expect(result.sumStr).toBe('55');
        expect(result.cntGt).toBe('3');
        expect(result.cntStr).toBe('3');
        expect(result.cntNe).toBe('2');
    });

    test('LEFT / RIGHT / MID / LEN / TRIM / UPPER / LOWER / CONCATENATE', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null;
            const ev = (f) => xlsx.evaluateFormula(f, resolver).value;
            return {
                left: ev('=LEFT("hello world", 5)'),
                right: ev('=RIGHT("hello world", 5)'),
                mid: ev('=MID("hello world", 7, 5)'),
                len: ev('=LEN("hello")'),
                trim: ev('=TRIM("  hello   world  ")'),
                upper: ev('=UPPER("hello")'),
                lower: ev('=LOWER("HELLO")'),
                concat: ev('=CONCATENATE("a", "b", "c")'),
                concatAlias: ev('=CONCAT("x", 1, "y")'),
            };
        });
        expect(result.left).toBe('hello');
        expect(result.right).toBe('world');
        expect(result.mid).toBe('world');
        expect(result.len).toBe('5');
        expect(result.trim).toBe('hello world');
        expect(result.upper).toBe('HELLO');
        expect(result.lower).toBe('hello');
        expect(result.concat).toBe('abc');
        expect(result.concatAlias).toBe('x1y');
    });

    test('VLOOKUP exact-match; FALSE (approximate) surfaces #VALUE!', async ({ page }) => {
        const result = await page.evaluate(() => {
            // Table: A:B
            // A1=1 B1="one"
            // A2=2 B2="two"
            // A3=3 B3="three"
            const table = [
                [1, 'one'],
                [2, 'two'],
                [3, 'three'],
            ];
            const resolver = (col, row) => {
                const r = table[row];
                if (!r) return null;
                const v = r[col];
                return v === undefined ? null : v;
            };
            const ev = (f) => xlsx.evaluateFormula(f, resolver).value;
            return {
                hit: ev('=VLOOKUP(2, A1:B3, 2)'),
                miss: ev('=VLOOKUP(99, A1:B3, 2, TRUE)'),
                col1: ev('=VLOOKUP(3, A1:B3, 1)'),
                approxPunt: ev('=VLOOKUP(2, A1:B3, 2, FALSE)'),
                outOfCol: ev('=VLOOKUP(2, A1:B3, 5)'),
            };
        });
        expect(result.hit).toBe('two');
        expect(result.miss).toBe('#N/A');
        expect(result.col1).toBe('3');
        // The 4th arg of VLOOKUP in this evaluator is named `exact`:
        // TRUE / omitted = exact-match, FALSE = approximate. Approximate
        // match is punted to a future wave; FALSE → #VALUE! so the gap is
        // visible to downstream consumers.
        expect(result.approxPunt).toBe('#VALUE!');
        expect(result.outOfCol).toBe('#REF!');
    });

    // ── v2: error taxonomy ───────────────────────────────────────────────

    test('error taxonomy: each code surfaces with its Excel sentinel', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null;
            const ev = (f) => xlsx.evaluateFormula(f, resolver).value;
            return {
                div0: ev('=5/0'),
                value: ev('="x"+1'),
                ref: ev('=#REF!'),
                name: ev('=NOSUCHFUNC(1)'),
                num: ev('=SQRT(-1)'),
                na: ev('=#N/A'),
                nullErr: ev('=#NULL!'),
            };
        });
        expect(result.div0).toBe('#DIV/0!');
        expect(result.value).toBe('#VALUE!');
        expect(result.ref).toBe('#REF!');
        expect(result.name).toBe('#NAME?');
        expect(result.num).toBe('#NUM!');
        expect(result.na).toBe('#N/A');
        expect(result.nullErr).toBe('#NULL!');
    });

    test('number-format routing: 0.1+0.2 cleans up to 0.3', async ({ page }) => {
        const result = await page.evaluate(() => {
            const resolver = () => null;
            const ev = (f) => xlsx.evaluateFormula(f, resolver).value;
            return {
                third: ev('=0.1+0.2'),
                thirdMinus: ev('=0.3-0.1'),
                integral: ev('=60'),
            };
        });
        // 0.1+0.2 must round-trip cleanly, not as 0.30000000000000004
        expect(result.third).toBe('0.3');
        expect(result.thirdMinus).toBe('0.2');
        expect(result.integral).toBe('60');
    });
});
