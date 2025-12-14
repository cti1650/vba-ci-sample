import { describe, it, expect, vi, beforeEach } from 'vitest';
import { generateEnumsFile, generateSummary, printSummary } from '../converter/generator.js';

// Mock utils
vi.mock('../converter/utils.js', () => ({
  writeFile: vi.fn(),
  joinLines: (lines) => lines.join('\r\n'),
  normalizeCrlf: (content) => content.replace(/\r?\n/g, '\r\n'),
}));

describe('generateEnumsFile', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('should return null for empty enums', () => {
    const result = generateEnumsFile('/output', new Map());
    expect(result).toBeNull();
  });

  it('should generate enum file content', async () => {
    const { writeFile } = await import('../converter/utils.js');

    const allEnums = new Map([
      ['Status', new Map([['Active', 1], ['Inactive', 0]])],
      ['Priority', new Map([['High', 1], ['Medium', 2], ['Low', 3]])],
    ]);

    const result = generateEnumsFile('/output', allEnums);

    expect(result).toBe('/output/_enums.vbs');
    expect(writeFile).toHaveBeenCalledWith(
      '/output/_enums.vbs',
      expect.stringContaining('Status_Active = 1')
    );
    expect(writeFile).toHaveBeenCalledWith(
      '/output/_enums.vbs',
      expect.stringContaining('Priority_High = 1')
    );
  });
});

describe('generateSummary', () => {
  it('should generate correct summary', () => {
    const options = {
      convertedFiles: ['Calculator.vbs', 'Utils.vbs'],
      allEnums: new Map([
        ['Status', new Map([['Active', 1], ['Inactive', 0]])],
      ]),
      allApis: new Map([
        ['GetTickCount', { lib: 'kernel32' }],
        ['CustomApi', { lib: 'custom.dll' }],
      ]),
      mockedApis: new Set(['GetTickCount']),
    };

    const result = generateSummary(options);

    expect(result.convertedCount).toBe(2);
    expect(result.convertedFiles).toEqual(['Calculator.vbs', 'Utils.vbs']);
    expect(result.enums).toHaveLength(1);
    expect(result.enums[0].name).toBe('Status');
    expect(result.enums[0].members).toEqual(['Active', 'Inactive']);
    expect(result.apis.mocked).toHaveLength(1);
    expect(result.apis.mocked[0].name).toBe('GetTickCount');
    expect(result.apis.unmocked).toHaveLength(1);
    expect(result.apis.unmocked[0].name).toBe('CustomApi');
  });
});

describe('printSummary', () => {
  it('should print summary without errors', () => {
    const consoleSpy = vi.spyOn(console, 'log').mockImplementation(() => {});

    const summary = {
      convertedCount: 2,
      convertedFiles: ['file1.vbs', 'file2.vbs'],
      enums: [
        { name: 'Status', members: ['Active', 'Inactive'] },
      ],
      apis: {
        mocked: [{ name: 'GetTickCount', lib: 'kernel32' }],
        unmocked: [{ name: 'CustomApi', lib: 'custom.dll' }],
      },
    };

    printSummary(summary);

    expect(consoleSpy).toHaveBeenCalled();
    consoleSpy.mockRestore();
  });
});
