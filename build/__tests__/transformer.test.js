import { describe, it, expect } from 'vitest';
import {
  applySkipRules,
  applyUnsupportedRules,
  applySyntaxTransforms,
  applyFunctionRenames,
  applyEnumConversions,
  convertEnumRefsToLiterals,
  applyCreateObjectMock,
  trimEmptyLines,
} from '../converter/transformer.js';

describe('applySkipRules', () => {
  const skipRules = {
    blocks: [
      { start: '^VERSION\\s+', end: '^END$' },
      { start: '^\\s*#If\\s+', end: '^\\s*#End\\s+If' },
    ],
    lines: [
      { pattern: '^\\s*Attribute\\s+VB_' },
      { pattern: '^\\s*Option\\s+Explicit' },
    ],
  };

  it('should skip VERSION block', () => {
    const lines = [
      'VERSION 1.0 CLASS',
      'BEGIN',
      '  MultiUse = -1',
      'END',
      'Public Sub Test()',
    ];

    const result = applySkipRules(lines, skipRules);

    expect(result).toEqual(['Public Sub Test()']);
  });

  it('should skip #If/#End If blocks', () => {
    const lines = [
      'Dim x',
      '#If VBA7 Then',
      '  Private Declare PtrSafe Function Sleep Lib "kernel32"',
      '#Else',
      '  Private Declare Function Sleep Lib "kernel32"',
      '#End If',
      'Dim y',
    ];

    const result = applySkipRules(lines, skipRules);

    expect(result).toEqual(['Dim x', 'Dim y']);
  });

  it('should skip Attribute lines', () => {
    const lines = [
      'Attribute VB_Name = "Test"',
      'Attribute VB_Exposed = False',
      'Public Sub Test()',
    ];

    const result = applySkipRules(lines, skipRules);

    expect(result).toEqual(['Public Sub Test()']);
  });

  it('should skip Option Explicit', () => {
    const lines = ['Option Explicit', 'Dim x'];

    const result = applySkipRules(lines, skipRules);

    expect(result).toEqual(['Dim x']);
  });
});

describe('applyUnsupportedRules', () => {
  const unsupportedRules = {
    commentOut: [
      { pattern: '^\\s*GoSub\\s+', prefix: "' [VBS UNSUPPORTED] " },
      { pattern: '^\\s*Open\\s+.*\\s+For\\s+', prefix: "' [VBS UNSUPPORTED] " },
    ],
    warnings: [
      { pattern: '\\bLSet\\s+', prefix: "' [VBS WARNING: LSet] " },
    ],
  };

  it('should comment out GoSub', () => {
    const line = '  GoSub ErrorHandler';

    const result = applyUnsupportedRules(line, unsupportedRules);

    expect(result).toBe("' [VBS UNSUPPORTED]   GoSub ErrorHandler");
  });

  it('should comment out Open For', () => {
    const line = 'Open "file.txt" For Output As #1';

    const result = applyUnsupportedRules(line, unsupportedRules);

    expect(result).toBe("' [VBS UNSUPPORTED] Open \"file.txt\" For Output As #1");
  });

  it('should add warning for LSet', () => {
    const line = 'LSet buffer = value';

    const result = applyUnsupportedRules(line, unsupportedRules);

    expect(result).toBe("' [VBS WARNING: LSet] LSet buffer = value");
  });
});

describe('applySyntaxTransforms', () => {
  const syntaxRules = {
    transforms: [
      {
        name: 'debug-print',
        pattern: '\\bDebug\\.Print\\b',
        replacement: 'DebugPrint',
        scope: 'all',
        priority: 100,
      },
      {
        name: 'static-to-dim',
        pattern: '^(\\s*)Static\\s+',
        replacement: '$1Dim ',
        scope: 'all',
        priority: 130,
      },
      {
        name: 'thisworkbook-path',
        pattern: '\\bThisWorkbook\\.Path\\b',
        replacement: {
          module: 'GetScriptDir()',
          class: 'CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)',
        },
        scope: 'contextual',
        priority: 200,
      },
    ],
    optionalParams: [],
    typeRemoval: [
      {
        name: 'dim-as-type',
        pattern: '(\\bDim\\s+\\w+)\\s+As\\s+\\w+',
        replacement: '$1',
        priority: 502,
      },
    ],
  };

  it('should convert Debug.Print to DebugPrint', () => {
    const line = '  Debug.Print "Hello"';

    const result = applySyntaxTransforms(line, syntaxRules, false);

    expect(result).toBe('  DebugPrint "Hello"');
  });

  it('should convert Static to Dim', () => {
    const line = '  Static counter';

    const result = applySyntaxTransforms(line, syntaxRules, false);

    expect(result).toBe('  Dim counter');
  });

  it('should convert ThisWorkbook.Path for module', () => {
    const line = 'path = ThisWorkbook.Path';

    const result = applySyntaxTransforms(line, syntaxRules, false);

    expect(result).toBe('path = GetScriptDir()');
  });

  it('should convert ThisWorkbook.Path for class', () => {
    const line = 'path = ThisWorkbook.Path';

    const result = applySyntaxTransforms(line, syntaxRules, true);

    expect(result).toContain('CreateObject("Scripting.FileSystemObject")');
  });

  it('should remove type declarations', () => {
    const line = 'Dim x As Long';

    const result = applySyntaxTransforms(line, syntaxRules, false);

    expect(result).toBe('Dim x');
  });
});

describe('applyFunctionRenames', () => {
  const renameRules = {
    renames: [
      { from: '\\bLeft\\$\\(', to: 'Left(' },
      { from: '\\bMid\\$\\(', to: 'Mid(' },
      { from: '\\bTrim\\$\\(', to: 'Trim(' },
    ],
  };

  it('should rename Left$ to Left', () => {
    const line = 's = Left$(str, 5)';

    const result = applyFunctionRenames(line, renameRules);

    expect(result).toBe('s = Left(str, 5)');
  });

  it('should rename Mid$ to Mid', () => {
    const line = 's = Mid$(str, 2, 3)';

    const result = applyFunctionRenames(line, renameRules);

    expect(result).toBe('s = Mid(str, 2, 3)');
  });

  it('should handle multiple renames in one line', () => {
    const line = 's = Trim$(Left$(str, 10))';

    const result = applyFunctionRenames(line, renameRules);

    expect(result).toBe('s = Trim(Left(str, 10))');
  });
});

describe('applyEnumConversions', () => {
  const allEnums = new Map([
    ['Status', new Map([['Active', 1], ['Inactive', 0]])],
    ['Priority', new Map([['High', 1], ['Low', 2]])],
  ]);

  it('should convert EnumName.Member to EnumName_Member', () => {
    const line = 'x = Status.Active';

    const result = applyEnumConversions(line, allEnums, false);

    expect(result).toBe('x = Status_Active');
  });

  it('should convert standalone member in assignment', () => {
    const line = 'x = Active';

    const result = applyEnumConversions(line, allEnums, false);

    expect(result).toBe('x = Status_Active');
  });
});

describe('convertEnumRefsToLiterals', () => {
  const allEnums = new Map([
    ['Status', new Map([['Active', 1], ['Inactive', 0]])],
  ]);

  it('should convert enum references to literal values', () => {
    const content = 'x = Status_Active\ny = Status_Inactive';

    const result = convertEnumRefsToLiterals(content, allEnums);

    expect(result).toBe('x = 1\ny = 0');
  });
});

describe('applyCreateObjectMock', () => {
  it('should convert CreateObject to CreateObjectMock', () => {
    const line = 'Set obj = CreateObject("Excel.Application")';

    const result = applyCreateObjectMock(line);

    expect(result).toBe('Set obj = CreateObjectMock("Excel.Application")');
  });

  it('should not convert FSO', () => {
    const line = 'Set fso = CreateObject("Scripting.FileSystemObject")';

    const result = applyCreateObjectMock(line);

    expect(result).toBe('Set fso = CreateObject("Scripting.FileSystemObject")');
  });

  it('should not convert Dictionary', () => {
    const line = 'Set dict = CreateObject("Scripting.Dictionary")';

    const result = applyCreateObjectMock(line);

    expect(result).toBe('Set dict = CreateObject("Scripting.Dictionary")');
  });

  it('should not double-convert', () => {
    const line = 'Set obj = CreateObjectMock("Excel.Application")';

    const result = applyCreateObjectMock(line);

    expect(result).toBe('Set obj = CreateObjectMock("Excel.Application")');
  });
});

describe('trimEmptyLines', () => {
  it('should trim empty lines from start and end', () => {
    const lines = ['', '  ', 'content', '', '  '];

    const result = trimEmptyLines(lines);

    expect(result).toEqual(['content']);
  });

  it('should preserve empty lines in middle', () => {
    const lines = ['first', '', 'second'];

    const result = trimEmptyLines(lines);

    expect(result).toEqual(['first', '', 'second']);
  });
});
