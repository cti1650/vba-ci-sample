import { describe, it, expect } from 'vitest';
import {
  collectEnumDefinitions,
  collectApiDeclarations,
  mergeEnumDefinitions,
  mergeApiDeclarations,
  enumMapToObject,
} from '../converter/parser.js';

describe('collectEnumDefinitions', () => {
  it('should collect enum with explicit values', () => {
    const lines = [
      'Public Enum MyEnum',
      '    First = 1',
      '    Second = 2',
      '    Third = 3',
      'End Enum',
    ];

    const result = collectEnumDefinitions(lines);

    expect(result.has('MyEnum')).toBe(true);
    expect(result.get('MyEnum').get('First')).toBe(1);
    expect(result.get('MyEnum').get('Second')).toBe(2);
    expect(result.get('MyEnum').get('Third')).toBe(3);
  });

  it('should collect enum with auto-increment values', () => {
    const lines = [
      'Private Enum Status',
      '    Pending',
      '    Active',
      '    Completed',
      'End Enum',
    ];

    const result = collectEnumDefinitions(lines);

    expect(result.has('Status')).toBe(true);
    expect(result.get('Status').get('Pending')).toBe(0);
    expect(result.get('Status').get('Active')).toBe(1);
    expect(result.get('Status').get('Completed')).toBe(2);
  });

  it('should handle mixed explicit and auto-increment values', () => {
    const lines = [
      'Enum Mixed',
      '    A = 10',
      '    B',
      '    C = 20',
      '    D',
      'End Enum',
    ];

    const result = collectEnumDefinitions(lines);

    expect(result.get('Mixed').get('A')).toBe(10);
    expect(result.get('Mixed').get('B')).toBe(11);
    expect(result.get('Mixed').get('C')).toBe(20);
    expect(result.get('Mixed').get('D')).toBe(21);
  });

  it('should handle negative values', () => {
    const lines = [
      'Enum Negative',
      '    Error = -1',
      '    None = 0',
      'End Enum',
    ];

    const result = collectEnumDefinitions(lines);

    expect(result.get('Negative').get('Error')).toBe(-1);
    expect(result.get('Negative').get('None')).toBe(0);
  });

  it('should handle multiple enums', () => {
    const lines = [
      'Enum First',
      '    A = 1',
      'End Enum',
      '',
      'Enum Second',
      '    B = 2',
      'End Enum',
    ];

    const result = collectEnumDefinitions(lines);

    expect(result.has('First')).toBe(true);
    expect(result.has('Second')).toBe(true);
    expect(result.get('First').get('A')).toBe(1);
    expect(result.get('Second').get('B')).toBe(2);
  });
});

describe('collectApiDeclarations', () => {
  it('should collect basic Declare Function', () => {
    const lines = [
      'Private Declare Function GetTickCount Lib "kernel32" () As Long',
    ];

    const result = collectApiDeclarations(lines);

    expect(result.has('GetTickCount')).toBe(true);
    expect(result.get('GetTickCount').lib).toBe('kernel32');
    expect(result.get('GetTickCount').alias).toBe('');
    expect(result.get('GetTickCount').isPtrSafe).toBe(false);
  });

  it('should collect PtrSafe Declare', () => {
    const lines = [
      'Private Declare PtrSafe Function Sleep Lib "kernel32" (ByVal ms As Long)',
    ];

    const result = collectApiDeclarations(lines);

    expect(result.has('Sleep')).toBe(true);
    expect(result.get('Sleep').isPtrSafe).toBe(true);
  });

  it('should collect Declare with Alias', () => {
    const lines = [
      'Private Declare Function GetUserNameA Lib "advapi32" Alias "GetUserNameA" (ByRef lpBuffer As String, ByRef nSize As Long) As Long',
    ];

    const result = collectApiDeclarations(lines);

    expect(result.has('GetUserNameA')).toBe(true);
    expect(result.get('GetUserNameA').alias).toBe('GetUserNameA');
  });

  it('should collect Declare Sub', () => {
    const lines = [
      'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)',
    ];

    const result = collectApiDeclarations(lines);

    expect(result.has('CopyMemory')).toBe(true);
    expect(result.get('CopyMemory').type).toBe('sub');
    expect(result.get('CopyMemory').alias).toBe('RtlMoveMemory');
  });

  it('should handle multiple declarations', () => {
    const lines = [
      'Private Declare Function GetTickCount Lib "kernel32" () As Long',
      'Private Declare Function Sleep Lib "kernel32" (ByVal ms As Long)',
      'Private Declare Function MessageBoxA Lib "user32" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long) As Long',
    ];

    const result = collectApiDeclarations(lines);

    expect(result.size).toBe(3);
    expect(result.has('GetTickCount')).toBe(true);
    expect(result.has('Sleep')).toBe(true);
    expect(result.has('MessageBoxA')).toBe(true);
  });
});

describe('mergeEnumDefinitions', () => {
  it('should merge multiple enum maps', () => {
    const map1 = new Map([
      ['Enum1', new Map([['A', 1]])],
    ]);
    const map2 = new Map([
      ['Enum2', new Map([['B', 2]])],
    ]);

    const result = mergeEnumDefinitions([map1, map2]);

    expect(result.has('Enum1')).toBe(true);
    expect(result.has('Enum2')).toBe(true);
  });

  it('should merge members of same enum from different files', () => {
    const map1 = new Map([
      ['Shared', new Map([['A', 1]])],
    ]);
    const map2 = new Map([
      ['Shared', new Map([['B', 2]])],
    ]);

    const result = mergeEnumDefinitions([map1, map2]);

    expect(result.get('Shared').get('A')).toBe(1);
    expect(result.get('Shared').get('B')).toBe(2);
  });
});

describe('mergeApiDeclarations', () => {
  it('should merge multiple API maps', () => {
    const map1 = new Map([
      ['Func1', { lib: 'lib1', alias: '', isPtrSafe: false, type: 'function' }],
    ]);
    const map2 = new Map([
      ['Func2', { lib: 'lib2', alias: '', isPtrSafe: true, type: 'function' }],
    ]);

    const result = mergeApiDeclarations([map1, map2]);

    expect(result.has('Func1')).toBe(true);
    expect(result.has('Func2')).toBe(true);
  });
});

describe('enumMapToObject', () => {
  it('should convert enum map to plain object', () => {
    const map = new Map([
      ['MyEnum', new Map([['A', 1], ['B', 2]])],
    ]);

    const result = enumMapToObject(map);

    expect(result).toEqual({
      MyEnum: { A: 1, B: 2 },
    });
  });
});
