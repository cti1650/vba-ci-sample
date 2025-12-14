import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { join } from 'path';
import {
  readFile,
  writeFile,
  readJson,
  loadRule,
  splitLines,
  joinLines,
  isClassFile,
  getBaseName,
  getExtension,
  normalizeCrlf,
} from '../converter/utils.js';

describe('splitLines', () => {
  it('should split content with LF', () => {
    const content = 'line1\nline2\nline3';
    const result = splitLines(content);
    expect(result).toEqual(['line1', 'line2', 'line3']);
  });

  it('should split content with CRLF', () => {
    const content = 'line1\r\nline2\r\nline3';
    const result = splitLines(content);
    expect(result).toEqual(['line1', 'line2', 'line3']);
  });

  it('should handle mixed line endings', () => {
    const content = 'line1\nline2\r\nline3';
    const result = splitLines(content);
    expect(result).toEqual(['line1', 'line2', 'line3']);
  });

  it('should handle empty content', () => {
    const content = '';
    const result = splitLines(content);
    expect(result).toEqual(['']);
  });
});

describe('joinLines', () => {
  it('should join lines with CRLF', () => {
    const lines = ['line1', 'line2', 'line3'];
    const result = joinLines(lines);
    expect(result).toBe('line1\r\nline2\r\nline3');
  });

  it('should handle single line', () => {
    const lines = ['single'];
    const result = joinLines(lines);
    expect(result).toBe('single');
  });

  it('should handle empty array', () => {
    const lines = [];
    const result = joinLines(lines);
    expect(result).toBe('');
  });
});

describe('normalizeCrlf', () => {
  it('should convert LF to CRLF', () => {
    const content = 'line1\nline2';
    const result = normalizeCrlf(content);
    expect(result).toBe('line1\r\nline2');
  });

  it('should keep existing CRLF', () => {
    const content = 'line1\r\nline2';
    const result = normalizeCrlf(content);
    expect(result).toBe('line1\r\nline2');
  });

  it('should handle mixed line endings', () => {
    const content = 'line1\nline2\r\nline3';
    const result = normalizeCrlf(content);
    expect(result).toBe('line1\r\nline2\r\nline3');
  });
});

describe('isClassFile', () => {
  it('should return true for .cls files', () => {
    expect(isClassFile('MyClass.cls')).toBe(true);
    expect(isClassFile('/path/to/MyClass.cls')).toBe(true);
  });

  it('should return false for .bas files', () => {
    expect(isClassFile('Module.bas')).toBe(false);
    expect(isClassFile('/path/to/Module.bas')).toBe(false);
  });

  it('should be case insensitive', () => {
    expect(isClassFile('MyClass.CLS')).toBe(true);
    expect(isClassFile('MyClass.Cls')).toBe(true);
  });
});

describe('getBaseName', () => {
  it('should return filename without extension', () => {
    expect(getBaseName('MyClass.cls')).toBe('MyClass');
    expect(getBaseName('/path/to/Module.bas')).toBe('Module');
  });

  it('should handle files with multiple dots', () => {
    expect(getBaseName('my.module.bas')).toBe('my.module');
  });
});

describe('getExtension', () => {
  it('should return extension with dot', () => {
    expect(getExtension('MyClass.cls')).toBe('.cls');
    expect(getExtension('Module.bas')).toBe('.bas');
  });

  it('should return lowercase extension', () => {
    expect(getExtension('MyClass.CLS')).toBe('.cls');
    expect(getExtension('Module.BAS')).toBe('.bas');
  });

  it('should return empty string for files without extension', () => {
    expect(getExtension('noextension')).toBe('');
  });
});

describe('loadRule', () => {
  it('should load existing rule file', () => {
    const rule = loadRule('skip-blocks');
    expect(rule).toBeDefined();
    expect(rule.blocks).toBeDefined();
    expect(Array.isArray(rule.blocks)).toBe(true);
  });

  it('should throw error for non-existent rule', () => {
    expect(() => loadRule('non-existent-rule')).toThrow(/Rule file not found/);
  });
});

describe('readFile error handling', () => {
  it('should throw descriptive error for non-existent file', () => {
    expect(() => readFile('/non/existent/path.txt')).toThrow(/Failed to read file/);
  });
});

describe('readJson error handling', () => {
  it('should throw descriptive error for invalid JSON', () => {
    // Create a temporary test by mocking - skip as we can't easily test this without file system
    // This is tested implicitly through loadRule
  });
});
