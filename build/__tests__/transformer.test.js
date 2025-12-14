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

  it('should NOT convert enum in string literals', () => {
    const line = 'msg = "Status is Active"';

    const result = applyEnumConversions(line, allEnums, false);

    expect(result).toBe('msg = "Status is Active"');
  });

  it('should NOT convert enum in comments', () => {
    const line = "x = 1 ' Status.Active is the default";

    const result = applyEnumConversions(line, allEnums, false);

    expect(result).toBe("x = 1 ' Status.Active is the default");
  });

  it('should convert enum before comment but not inside', () => {
    const line = "x = Status.Active ' Set to Active";

    const result = applyEnumConversions(line, allEnums, false);

    expect(result).toBe("x = Status_Active ' Set to Active");
  });

  it('should NOT convert enum inside string even with code after', () => {
    const line = 'msg = "Active" : x = Status.Active';

    const result = applyEnumConversions(line, allEnums, false);

    expect(result).toBe('msg = "Active" : x = Status_Active');
  });

  it('should handle multiple enum references correctly', () => {
    const line = 'If s = Status.Active And p = Priority.High Then';

    const result = applyEnumConversions(line, allEnums, false);

    expect(result).toBe('If s = Status_Active And p = Priority_High Then');
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

describe('nested control structures', () => {
  describe('applySkipRules with nested blocks', () => {
    const skipRules = {
      blocks: [
        { start: '^\\s*#If\\s+', end: '^\\s*#End\\s+If' },
        { start: '^\\s*(Public\\s+|Private\\s+)?Enum\\s+', end: '^\\s*End\\s+Enum' },
        { start: '^\\s*(Public\\s+|Private\\s+)?Type\\s+', end: '^\\s*End\\s+Type' },
      ],
      lines: [],
    };

    it('should skip simple #If block', () => {
      const lines = [
        'Dim x',
        '#If VBA7 Then',
        '  Dim p As LongPtr',
        '#End If',
        'Dim y',
      ];

      const result = applySkipRules(lines, skipRules);

      expect(result).toEqual(['Dim x', 'Dim y']);
    });

    it('should skip #If with #Else', () => {
      const lines = [
        'Dim x',
        '#If VBA7 Then',
        '  Dim p As LongPtr',
        '#Else',
        '  Dim p As Long',
        '#End If',
        'Dim y',
      ];

      const result = applySkipRules(lines, skipRules);

      expect(result).toEqual(['Dim x', 'Dim y']);
    });

    it('should skip #If with #ElseIf', () => {
      const lines = [
        'Dim x',
        '#If Win64 Then',
        '  Dim p As LongLong',
        '#ElseIf VBA7 Then',
        '  Dim p As LongPtr',
        '#Else',
        '  Dim p As Long',
        '#End If',
        'Dim y',
      ];

      const result = applySkipRules(lines, skipRules);

      expect(result).toEqual(['Dim x', 'Dim y']);
    });

    it('should skip multiple sequential #If blocks', () => {
      const lines = [
        'Dim x',
        '#If VBA7 Then',
        '  Private Declare PtrSafe Function Sleep Lib "kernel32"',
        '#End If',
        '#If MAC Then',
        '  Private Declare Function Sleep Lib "libc"',
        '#End If',
        'Dim y',
      ];

      const result = applySkipRules(lines, skipRules);

      expect(result).toEqual(['Dim x', 'Dim y']);
    });

    it('should skip Enum block', () => {
      const lines = [
        'Dim x',
        'Public Enum Status',
        '  Active = 1',
        '  Inactive = 0',
        'End Enum',
        'Dim y',
      ];

      const result = applySkipRules(lines, skipRules);

      expect(result).toEqual(['Dim x', 'Dim y']);
    });

    it('should skip Type block', () => {
      const lines = [
        'Dim x',
        'Private Type Person',
        '  Name As String',
        '  Age As Long',
        'End Type',
        'Dim y',
      ];

      const result = applySkipRules(lines, skipRules);

      expect(result).toEqual(['Dim x', 'Dim y']);
    });

    it('should skip multiple sequential blocks of different types', () => {
      const lines = [
        'Dim x',
        'Public Enum Status',
        '  Active = 1',
        'End Enum',
        'Private Type Person',
        '  Name As String',
        'End Type',
        '#If VBA7 Then',
        '  Dim p As LongPtr',
        '#End If',
        'Dim y',
      ];

      const result = applySkipRules(lines, skipRules);

      expect(result).toEqual(['Dim x', 'Dim y']);
    });
  });

  describe('nested If/For in actual VBA code (not skip blocks)', () => {
    const syntaxRules = {
      transforms: [
        {
          name: 'debug-print',
          pattern: '\\bDebug\\.Print\\b',
          replacement: 'DebugPrint',
          scope: 'all',
          priority: 100,
        },
      ],
      optionalParams: [],
      typeRemoval: [],
    };

    it('should preserve nested If structure', () => {
      const lines = [
        'If x > 0 Then',
        '  If y > 0 Then',
        '    Debug.Print "Both positive"',
        '  End If',
        'End If',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      expect(result).toEqual([
        'If x > 0 Then',
        '  If y > 0 Then',
        '    DebugPrint "Both positive"',
        '  End If',
        'End If',
      ]);
    });

    it('should preserve nested For structure', () => {
      const lines = [
        'For i = 1 To 10',
        '  For j = 1 To 10',
        '    Debug.Print i * j',
        '  Next j',
        'Next i',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      expect(result).toEqual([
        'For i = 1 To 10',
        '  For j = 1 To 10',
        '    DebugPrint i * j',
        '  Next j',
        'Next i',
      ]);
    });

    it('should preserve mixed If/For nesting', () => {
      const lines = [
        'For i = 1 To 10',
        '  If i Mod 2 = 0 Then',
        '    For j = 1 To 5',
        '      Debug.Print i * j',
        '    Next j',
        '  End If',
        'Next i',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      expect(result).toEqual([
        'For i = 1 To 10',
        '  If i Mod 2 = 0 Then',
        '    For j = 1 To 5',
        '      DebugPrint i * j',
        '    Next j',
        '  End If',
        'Next i',
      ]);
    });

    it('should handle Select Case inside For loop', () => {
      const lines = [
        'For i = 1 To 3',
        '  Select Case i',
        '    Case 1',
        '      Debug.Print "One"',
        '    Case 2',
        '      Debug.Print "Two"',
        '    Case Else',
        '      Debug.Print "Other"',
        '  End Select',
        'Next i',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      expect(result).toEqual([
        'For i = 1 To 3',
        '  Select Case i',
        '    Case 1',
        '      DebugPrint "One"',
        '    Case 2',
        '      DebugPrint "Two"',
        '    Case Else',
        '      DebugPrint "Other"',
        '  End Select',
        'Next i',
      ]);
    });

    it('should handle Do While inside If', () => {
      const lines = [
        'If enabled Then',
        '  Do While count < 10',
        '    Debug.Print count',
        '    count = count + 1',
        '  Loop',
        'End If',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      expect(result).toEqual([
        'If enabled Then',
        '  Do While count < 10',
        '    DebugPrint count',
        '    count = count + 1',
        '  Loop',
        'End If',
      ]);
    });

    it('should handle With block inside For Each', () => {
      const lines = [
        'For Each item In collection',
        '  With item',
        '    Debug.Print .Name',
        '    Debug.Print .Value',
        '  End With',
        'Next item',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      expect(result).toEqual([
        'For Each item In collection',
        '  With item',
        '    DebugPrint .Name',
        '    DebugPrint .Value',
        '  End With',
        'Next item',
      ]);
    });

    it('should handle deeply nested structures (3+ levels)', () => {
      const lines = [
        'For i = 1 To 3',
        '  If i > 1 Then',
        '    For j = 1 To 2',
        '      If j = 1 Then',
        '        Debug.Print "i=" & i & ", j=" & j',
        '      End If',
        '    Next j',
        '  End If',
        'Next i',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      expect(result).toEqual([
        'For i = 1 To 3',
        '  If i > 1 Then',
        '    For j = 1 To 2',
        '      If j = 1 Then',
        '        DebugPrint "i=" & i & ", j=" & j',
        '      End If',
        '    Next j',
        '  End If',
        'Next i',
      ]);
    });

    it('should handle Sub/Function inside If (error handling pattern)', () => {
      const lines = [
        'Sub ProcessData()',
        '  On Error GoTo ErrorHandler',
        '  If data Is Nothing Then',
        '    Exit Sub',
        '  End If',
        '  Debug.Print "Processing..."',
        '  Exit Sub',
        'ErrorHandler:',
        '  Debug.Print Err.Description',
        'End Sub',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      expect(result[5]).toBe('  DebugPrint "Processing..."');
      expect(result[8]).toBe('  DebugPrint Err.Description');
    });
  });

  describe('enum conversion with nested structures', () => {
    const allEnums = new Map([
      ['Status', new Map([['Active', 1], ['Inactive', 0]])],
    ]);

    it('should convert enum in nested If', () => {
      const lines = [
        'If condition Then',
        '  If status = Status.Active Then',
        '    result = True',
        '  End If',
        'End If',
      ];

      const result = lines.map(line => applyEnumConversions(line, allEnums, false));

      expect(result[1]).toBe('  If status = Status_Active Then');
    });

    it('should convert enum in nested For', () => {
      const lines = [
        'For i = 1 To 10',
        '  For Each item In items',
        '    If item.Status = Status.Active Then',
        '      count = count + 1',
        '    End If',
        '  Next',
        'Next i',
      ];

      const result = lines.map(line => applyEnumConversions(line, allEnums, false));

      expect(result[2]).toBe('    If item.Status = Status_Active Then');
    });
  });

  describe('With block nesting', () => {
    const syntaxRules = {
      transforms: [
        {
          name: 'debug-print',
          pattern: '\\bDebug\\.Print\\b',
          replacement: 'DebugPrint',
          scope: 'all',
          priority: 100,
        },
      ],
      optionalParams: [],
      typeRemoval: [],
    };

    it('should preserve nested With blocks', () => {
      const lines = [
        'With objParent',
        '  .Name = "Parent"',
        '  With .Child',
        '    .Name = "Child"',
        '    Debug.Print .Name',
        '  End With',
        '  Debug.Print .Name',
        'End With',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      expect(result).toEqual([
        'With objParent',
        '  .Name = "Parent"',
        '  With .Child',
        '    .Name = "Child"',
        '    DebugPrint .Name',
        '  End With',
        '  DebugPrint .Name',
        'End With',
      ]);
    });

    it('should preserve deeply nested With blocks (3 levels)', () => {
      const lines = [
        'With Application',
        '  With .ActiveWorkbook',
        '    With .ActiveSheet',
        '      Debug.Print .Name',
        '    End With',
        '  End With',
        'End With',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      expect(result).toEqual([
        'With Application',
        '  With .ActiveWorkbook',
        '    With .ActiveSheet',
        '      DebugPrint .Name',
        '    End With',
        '  End With',
        'End With',
      ]);
    });

    it('should handle With inside For loop', () => {
      const lines = [
        'For Each ws In Worksheets',
        '  With ws',
        '    .Name = "Sheet" & i',
        '    With .Range("A1")',
        '      .Value = "Header"',
        '    End With',
        '  End With',
        'Next ws',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      expect(result).toEqual([
        'For Each ws In Worksheets',
        '  With ws',
        '    .Name = "Sheet" & i',
        '    With .Range("A1")',
        '      .Value = "Header"',
        '    End With',
        '  End With',
        'Next ws',
      ]);
    });

    it('should handle With with If inside', () => {
      const lines = [
        'With obj',
        '  If .Enabled Then',
        '    With .Settings',
        '      If .AutoSave Then',
        '        .Save',
        '      End If',
        '    End With',
        '  End If',
        'End With',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      expect(result).toEqual([
        'With obj',
        '  If .Enabled Then',
        '    With .Settings',
        '      If .AutoSave Then',
        '        .Save',
        '      End If',
        '    End With',
        '  End If',
        'End With',
      ]);
    });

    it('should handle member access with dot prefix preserved', () => {
      const lines = [
        'With rng',
        '  .Value = 100',
        '  .Font.Bold = True',
        '  .Interior.Color = RGB(255, 0, 0)',
        'End With',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      // Dot prefix should be preserved
      expect(result[1]).toBe('  .Value = 100');
      expect(result[2]).toBe('  .Font.Bold = True');
      expect(result[3]).toBe('  .Interior.Color = RGB(255, 0, 0)');
    });

    it('should not confuse With member access with object method calls', () => {
      const lines = [
        'With obj',
        '  result = .GetValue()',
        '  other = someObject.Method()',
        '  Debug.Print .Name & " - " & other.Name',
        'End With',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      expect(result[1]).toBe('  result = .GetValue()');
      expect(result[2]).toBe('  other = someObject.Method()');
      expect(result[3]).toBe('  DebugPrint .Name & " - " & other.Name');
    });
  });

  describe('sequential nested blocks independence', () => {
    const syntaxRules = {
      transforms: [
        {
          name: 'debug-print',
          pattern: '\\bDebug\\.Print\\b',
          replacement: 'DebugPrint',
          scope: 'all',
          priority: 100,
        },
      ],
      optionalParams: [],
      typeRemoval: [],
    };

    it('should treat sequential 2-level If blocks as separate', () => {
      const lines = [
        'If a Then',
        '  If b Then',
        '    Debug.Print "First block"',
        '  End If',
        'End If',
        'If c Then',
        '  If d Then',
        '    Debug.Print "Second block"',
        '  End If',
        'End If',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      // Both blocks should be preserved independently
      expect(result).toEqual([
        'If a Then',
        '  If b Then',
        '    DebugPrint "First block"',
        '  End If',
        'End If',
        'If c Then',
        '  If d Then',
        '    DebugPrint "Second block"',
        '  End If',
        'End If',
      ]);
    });

    it('should treat sequential 2-level For blocks as separate', () => {
      const lines = [
        'For i = 1 To 5',
        '  For j = 1 To 3',
        '    Debug.Print "Block 1: " & i & "," & j',
        '  Next j',
        'Next i',
        'For k = 1 To 5',
        '  For l = 1 To 3',
        '    Debug.Print "Block 2: " & k & "," & l',
        '  Next l',
        'Next k',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      expect(result).toEqual([
        'For i = 1 To 5',
        '  For j = 1 To 3',
        '    DebugPrint "Block 1: " & i & "," & j',
        '  Next j',
        'Next i',
        'For k = 1 To 5',
        '  For l = 1 To 3',
        '    DebugPrint "Block 2: " & k & "," & l',
        '  Next l',
        'Next k',
      ]);
    });

    it('should treat sequential 2-level With blocks as separate', () => {
      const lines = [
        'With obj1',
        '  With .Child1',
        '    Debug.Print .Name',
        '  End With',
        'End With',
        'With obj2',
        '  With .Child2',
        '    Debug.Print .Value',
        '  End With',
        'End With',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      expect(result).toEqual([
        'With obj1',
        '  With .Child1',
        '    DebugPrint .Name',
        '  End With',
        'End With',
        'With obj2',
        '  With .Child2',
        '    DebugPrint .Value',
        '  End With',
        'End With',
      ]);
    });

    it('should handle mixed sequential nested blocks', () => {
      const lines = [
        'If x Then',
        '  If y Then',
        '    Debug.Print "If block"',
        '  End If',
        'End If',
        'For i = 1 To 3',
        '  For j = 1 To 2',
        '    Debug.Print "For block"',
        '  Next j',
        'Next i',
        'With obj',
        '  With .Child',
        '    Debug.Print "With block"',
        '  End With',
        'End With',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      expect(result).toEqual([
        'If x Then',
        '  If y Then',
        '    DebugPrint "If block"',
        '  End If',
        'End If',
        'For i = 1 To 3',
        '  For j = 1 To 2',
        '    DebugPrint "For block"',
        '  Next j',
        'Next i',
        'With obj',
        '  With .Child',
        '    DebugPrint "With block"',
        '  End With',
        'End With',
      ]);
    });

    it('should handle code between sequential nested blocks', () => {
      const lines = [
        'If a Then',
        '  If b Then',
        '    x = 1',
        '  End If',
        'End If',
        'Debug.Print "Between blocks"',
        'y = 2',
        'If c Then',
        '  If d Then',
        '    z = 3',
        '  End If',
        'End If',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      expect(result[5]).toBe('DebugPrint "Between blocks"');
      expect(result[6]).toBe('y = 2');
    });

    it('should handle 3 sequential 2-level blocks correctly', () => {
      const lines = [
        'If a Then',
        '  If a2 Then',
        '    Debug.Print "A"',
        '  End If',
        'End If',
        'If b Then',
        '  If b2 Then',
        '    Debug.Print "B"',
        '  End If',
        'End If',
        'If c Then',
        '  If c2 Then',
        '    Debug.Print "C"',
        '  End If',
        'End If',
      ];

      const result = lines.map(line => applySyntaxTransforms(line, syntaxRules, false));

      // All 3 blocks should be separate
      expect(result[2]).toBe('    DebugPrint "A"');
      expect(result[7]).toBe('    DebugPrint "B"');
      expect(result[12]).toBe('    DebugPrint "C"');
    });
  });

  describe('skip blocks sequential independence', () => {
    const skipRules = {
      blocks: [
        { start: '^\\s*#If\\s+', end: '^\\s*#End\\s+If' },
      ],
      lines: [],
    };

    it('should skip sequential #If blocks independently', () => {
      const lines = [
        'Dim a',
        '#If VBA7 Then',
        '  #If Win64 Then',
        '    Dim p64 As LongLong',
        '  #End If',
        '#End If',
        'Dim b',
        '#If MAC Then',
        '  #If DEBUG Then',
        '    Dim debug',
        '  #End If',
        '#End If',
        'Dim c',
      ];

      const result = applySkipRules(lines, skipRules);

      // Both #If blocks should be skipped, preserving Dim statements
      expect(result).toEqual(['Dim a', 'Dim b', 'Dim c']);
    });

    it('should not merge sequential Enum blocks', () => {
      const skipRulesWithEnum = {
        blocks: [
          { start: '^\\s*(Public\\s+|Private\\s+)?Enum\\s+', end: '^\\s*End\\s+Enum' },
        ],
        lines: [],
      };

      const lines = [
        'Dim x',
        'Public Enum Status',
        '  Active = 1',
        '  Inactive = 0',
        'End Enum',
        'Dim y',
        'Public Enum Priority',
        '  High = 1',
        '  Low = 2',
        'End Enum',
        'Dim z',
      ];

      const result = applySkipRules(lines, skipRulesWithEnum);

      expect(result).toEqual(['Dim x', 'Dim y', 'Dim z']);
    });
  });
});
