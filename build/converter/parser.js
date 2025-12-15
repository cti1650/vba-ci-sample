import { readFile, splitLines, isClassFile, getBaseName } from './utils.js';

/**
 * Parse VBA file and extract metadata
 * @param {string} filePath - Path to the VBA file
 * @returns {object} Parsed VBA metadata
 */
export function parseVbaFile(filePath) {
  const content = readFile(filePath);
  const lines = splitLines(content);
  const isClass = isClassFile(filePath);
  const fileName = getBaseName(filePath);

  return {
    filePath,
    fileName,
    isClass,
    content,
    lines,
    enums: collectEnumDefinitions(lines),
    apiDeclarations: collectApiDeclarations(lines),
  };
}

/**
 * Collect Enum definitions from VBA lines
 * @param {string[]} lines - Array of VBA lines
 * @returns {Map<string, Map<string, number>>} Map of enum name -> Map of member name -> value
 */
export function collectEnumDefinitions(lines) {
  const enums = new Map();
  let currentEnumName = '';
  let inEnum = false;
  let autoValue = 0;

  for (const line of lines) {
    // Start of Enum block
    const enumStartMatch = line.match(/^\s*(Public\s+|Private\s+)?Enum\s+(\w+)/i);
    if (enumStartMatch) {
      inEnum = true;
      currentEnumName = enumStartMatch[2];
      autoValue = 0;
      if (!enums.has(currentEnumName)) {
        enums.set(currentEnumName, new Map());
      }
      continue;
    }

    if (inEnum) {
      // End of Enum block
      if (/^\s*End\s+Enum/i.test(line)) {
        inEnum = false;
        currentEnumName = '';
        continue;
      }

      // Enum member with explicit value: MemberName = Value
      const memberWithValueMatch = line.match(/^\s*(\w+)\s*=\s*(-?\d+)/);
      if (memberWithValueMatch) {
        const memberName = memberWithValueMatch[1];
        const value = parseInt(memberWithValueMatch[2], 10);
        enums.get(currentEnumName).set(memberName, value);
        autoValue = value + 1;
        continue;
      }

      // Enum member without value (auto-increment)
      const memberOnlyMatch = line.match(/^\s*(\w+)\s*$/);
      if (memberOnlyMatch && memberOnlyMatch[1] !== '') {
        const memberName = memberOnlyMatch[1];
        enums.get(currentEnumName).set(memberName, autoValue);
        autoValue++;
      }
    }
  }

  return enums;
}

/**
 * Collect API (Declare) statements from VBA lines
 * @param {string[]} lines - Array of VBA lines
 * @returns {Map<string, object>} Map of function name -> declaration info
 */
export function collectApiDeclarations(lines) {
  const apis = new Map();

  for (const line of lines) {
    // Match Declare statements
    // Examples:
    // Private Declare Function GetTickCount Lib "kernel32" () As Long
    // Private Declare PtrSafe Function Sleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
    const declareMatch = line.match(
      /^\s*(Public\s+|Private\s+)?Declare\s+(PtrSafe\s+)?(Function|Sub)\s+(\w+)\s+Lib\s+"([^"]+)"/i
    );

    if (declareMatch) {
      const funcName = declareMatch[4];
      const libName = declareMatch[5];

      // Check for Alias
      const aliasMatch = line.match(/Alias\s+"([^"]+)"/i);
      const alias = aliasMatch ? aliasMatch[1] : '';

      apis.set(funcName, {
        lib: libName,
        alias,
        originalLine: line.trim(),
        isPtrSafe: !!declareMatch[2],
        type: declareMatch[3].toLowerCase(), // 'function' or 'sub'
      });
    }
  }

  return apis;
}

/**
 * Merge enum definitions from multiple files
 * @param {Map<string, Map<string, number>>[]} enumMaps - Array of enum maps
 * @returns {Map<string, Map<string, number>>} Merged enum map
 */
export function mergeEnumDefinitions(enumMaps) {
  const merged = new Map();

  for (const enumMap of enumMaps) {
    for (const [enumName, members] of enumMap) {
      if (!merged.has(enumName)) {
        merged.set(enumName, new Map());
      }
      for (const [memberName, value] of members) {
        merged.get(enumName).set(memberName, value);
      }
    }
  }

  return merged;
}

/**
 * Merge API declarations from multiple files
 * @param {Map<string, object>[]} apiMaps - Array of API maps
 * @returns {Map<string, object>} Merged API map
 */
export function mergeApiDeclarations(apiMaps) {
  const merged = new Map();

  for (const apiMap of apiMaps) {
    for (const [funcName, info] of apiMap) {
      if (!merged.has(funcName)) {
        merged.set(funcName, info);
      }
    }
  }

  return merged;
}

/**
 * Convert enum Map to plain object for JSON serialization
 * @param {Map<string, Map<string, number>>} enumMap - Enum map
 * @returns {object} Plain object representation
 */
export function enumMapToObject(enumMap) {
  const result = {};
  for (const [enumName, members] of enumMap) {
    result[enumName] = {};
    for (const [memberName, value] of members) {
      result[enumName][memberName] = value;
    }
  }
  return result;
}

/**
 * Convert API Map to plain object for JSON serialization
 * @param {Map<string, object>} apiMap - API map
 * @returns {object} Plain object representation
 */
export function apiMapToObject(apiMap) {
  const result = {};
  for (const [funcName, info] of apiMap) {
    result[funcName] = info;
  }
  return result;
}
