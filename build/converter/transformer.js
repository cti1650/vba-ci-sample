import { loadRule, splitLines, joinLines } from './utils.js';

/**
 * Transform VBA content to VBS
 * @param {object} options - Transformation options
 * @param {string} options.content - VBA content
 * @param {boolean} options.isClass - Whether the file is a class
 * @param {string} options.className - Class name (for .cls files)
 * @param {Map<string, Map<string, number>>} options.allEnums - All enum definitions
 * @param {Map<string, object>} options.allApis - All API declarations
 * @param {Set<string>} options.mockedApis - Set of mocked API function names
 * @param {boolean} options.useMockCreateObject - Whether to mock CreateObject calls
 * @returns {string} Transformed VBS content
 */
export function transformVbaToVbs(options) {
  const {
    content,
    isClass,
    className,
    allEnums = new Map(),
    allApis = new Map(),
    mockedApis = new Set(),
    useMockCreateObject = false,
  } = options;

  // Load rules
  const skipBlocksRule = loadRule('skip-blocks');
  const syntaxTransformsRule = loadRule('syntax-transforms');
  const functionRenamesRule = loadRule('function-renames');
  const unsupportedRule = loadRule('unsupported');

  let lines = splitLines(content);

  // Step 1: Skip blocks and lines
  lines = applySkipRules(lines, skipBlocksRule);

  // Step 2: Apply line-by-line transformations
  lines = lines.map((line) => {
    let transformed = line;

    // Apply unsupported syntax rules (comment out)
    transformed = applyUnsupportedRules(transformed, unsupportedRule);

    // Apply syntax transforms
    transformed = applySyntaxTransforms(transformed, syntaxTransformsRule, isClass);

    // Apply function renames
    transformed = applyFunctionRenames(transformed, functionRenamesRule);

    // Apply enum reference conversions
    transformed = applyEnumConversions(transformed, allEnums, isClass);

    // Apply API call warnings for unmocked APIs
    transformed = applyApiWarnings(transformed, allApis, mockedApis);

    // Apply CreateObject mock if enabled
    if (useMockCreateObject) {
      transformed = applyCreateObjectMock(transformed);
    }

    return transformed;
  });

  // Step 3: Trim empty lines at start and end
  lines = trimEmptyLines(lines);

  // Step 4: Convert enum references to literals in class files
  let body = joinLines(lines);
  if (isClass) {
    body = convertEnumRefsToLiterals(body, allEnums);
  }

  // Step 5: Wrap class files
  if (isClass && className) {
    return `Class ${className}\r\n${body}\r\nEnd Class`;
  }

  return body;
}

/**
 * Apply skip rules to filter out blocks and lines
 * @param {string[]} lines - Input lines
 * @param {object} skipRules - Skip rules configuration
 * @returns {string[]} Filtered lines
 */
export function applySkipRules(lines, skipRules) {
  const result = [];
  let skipUntilPattern = null;

  for (const line of lines) {
    // Check if we're in a skip block
    if (skipUntilPattern) {
      if (new RegExp(skipUntilPattern, 'i').test(line)) {
        skipUntilPattern = null;
      }
      continue;
    }

    // Check for block start
    let shouldSkipBlock = false;
    for (const block of skipRules.blocks) {
      if (new RegExp(block.start, 'i').test(line)) {
        skipUntilPattern = block.end;
        shouldSkipBlock = true;
        break;
      }
    }
    if (shouldSkipBlock) continue;

    // Check for single line skip
    let shouldSkipLine = false;
    for (const lineRule of skipRules.lines) {
      if (new RegExp(lineRule.pattern, 'i').test(line)) {
        shouldSkipLine = true;
        break;
      }
    }
    if (shouldSkipLine) continue;

    result.push(line);
  }

  return result;
}

/**
 * Apply unsupported syntax rules (comment out lines)
 * @param {string} line - Input line
 * @param {object} unsupportedRules - Unsupported rules configuration
 * @returns {string} Transformed line
 */
export function applyUnsupportedRules(line, unsupportedRules) {
  let result = line;

  // Apply comment out rules
  for (const rule of unsupportedRules.commentOut) {
    if (new RegExp(rule.pattern, 'i').test(result)) {
      // Check exclude pattern if present
      if (rule.excludePattern && new RegExp(rule.excludePattern, 'i').test(result)) {
        continue;
      }
      result = rule.prefix + result;
      break; // Only apply one comment rule per line
    }
  }

  // Apply warning rules (prepend warning but keep line)
  for (const rule of unsupportedRules.warnings) {
    if (new RegExp(rule.pattern, 'i').test(result) && !result.startsWith("'")) {
      result = rule.prefix + result;
      break;
    }
  }

  return result;
}

/**
 * Apply syntax transformation rules
 * @param {string} line - Input line
 * @param {object} syntaxRules - Syntax transformation rules
 * @param {boolean} isClass - Whether processing a class file
 * @returns {string} Transformed line
 */
export function applySyntaxTransforms(line, syntaxRules, isClass) {
  let result = line;

  // Sort transforms by priority
  const allTransforms = [
    ...(syntaxRules.transforms || []),
    ...(syntaxRules.optionalParams || []),
    ...(syntaxRules.typeRemoval || []),
  ].sort((a, b) => (a.priority || 0) - (b.priority || 0));

  for (const transform of allTransforms) {
    const pattern = new RegExp(transform.pattern, 'gi');

    if (transform.scope === 'contextual' && typeof transform.replacement === 'object') {
      // Contextual replacement (different for module vs class)
      const replacement = isClass ? transform.replacement.class : transform.replacement.module;
      result = result.replace(pattern, replacement);
    } else {
      // Standard replacement
      result = result.replace(pattern, transform.replacement);
    }
  }

  return result;
}

/**
 * Apply function rename rules
 * @param {string} line - Input line
 * @param {object} renameRules - Function rename rules
 * @returns {string} Transformed line
 */
export function applyFunctionRenames(line, renameRules) {
  let result = line;

  for (const rename of renameRules.renames) {
    const pattern = new RegExp(rename.from, 'g');
    result = result.replace(pattern, rename.to);
  }

  return result;
}

/**
 * Apply enum reference conversions
 * @param {string} line - Input line
 * @param {Map<string, Map<string, number>>} allEnums - All enum definitions
 * @param {boolean} isClass - Whether processing a class file
 * @returns {string} Transformed line
 */
export function applyEnumConversions(line, allEnums, isClass) {
  let result = line;

  for (const [enumName, members] of allEnums) {
    for (const [memberName] of members) {
      // Convert EnumName.MemberName -> EnumName_MemberName
      const dotPattern = new RegExp(`\\b${enumName}\\.${memberName}\\b`, 'g');
      result = result.replace(dotPattern, `${enumName}_${memberName}`);
    }
  }

  // Convert standalone enum member references (for same-file enums)
  // This handles cases like: = MemberName -> = EnumName_MemberName
  for (const [enumName, members] of allEnums) {
    for (const [memberName] of members) {
      // Only in assignment context to avoid false positives
      const assignPattern = new RegExp(`=\\s*\\b${memberName}\\b(?!\\s*[.(])`, 'g');
      result = result.replace(assignPattern, `= ${enumName}_${memberName}`);
    }
  }

  return result;
}

/**
 * Convert enum references to literal values (for class files)
 * @param {string} content - Content to transform
 * @param {Map<string, Map<string, number>>} allEnums - All enum definitions
 * @returns {string} Transformed content
 */
export function convertEnumRefsToLiterals(content, allEnums) {
  let result = content;

  for (const [enumName, members] of allEnums) {
    for (const [memberName, value] of members) {
      const pattern = new RegExp(`\\b${enumName}_${memberName}\\b`, 'g');
      result = result.replace(pattern, String(value));
    }
  }

  return result;
}

/**
 * Apply API call warnings for unmocked APIs
 * @param {string} line - Input line
 * @param {Map<string, object>} allApis - All API declarations
 * @param {Set<string>} mockedApis - Set of mocked API function names
 * @returns {string} Transformed line
 */
export function applyApiWarnings(line, allApis, mockedApis) {
  let result = line;

  for (const [apiName, info] of allApis) {
    const pattern = new RegExp(`\\b${apiName}\\s*\\(`, 'g');
    if (pattern.test(result) && !mockedApis.has(apiName)) {
      if (!result.includes("' [VBS API MOCK]")) {
        result = `' [VBS API MOCK: ${apiName} from ${info.lib} - returns default value] ${result}`;
      }
    }
  }

  return result;
}

/**
 * Apply CreateObject mock transformation
 * @param {string} line - Input line
 * @returns {string} Transformed line
 */
export function applyCreateObjectMock(line) {
  // Don't transform if already mocked
  if (line.includes('CreateObjectMock')) {
    return line;
  }

  // Match CreateObject("ProgID")
  const match = line.match(/CreateObject\s*\(\s*"([^"]+)"/);
  if (match) {
    const progId = match[1];
    // Don't mock FSO and Dictionary (basic infrastructure)
    if (!/^Scripting\.(FileSystemObject|Dictionary)$/i.test(progId)) {
      return line.replace(/\bCreateObject\s*\(/, 'CreateObjectMock(');
    }
  }

  return line;
}

/**
 * Trim empty lines from start and end of array
 * @param {string[]} lines - Input lines
 * @returns {string[]} Trimmed lines
 */
export function trimEmptyLines(lines) {
  let result = [...lines];

  // Trim start
  while (result.length > 0 && /^\s*$/.test(result[0])) {
    result.shift();
  }

  // Trim end
  while (result.length > 0 && /^\s*$/.test(result[result.length - 1])) {
    result.pop();
  }

  return result;
}

/**
 * Collect local enum definitions from content
 * @param {string} content - VBA content
 * @returns {Array<{varName: string, value: number}>} Array of enum variable definitions
 */
export function collectLocalEnumDefs(content) {
  const lines = splitLines(content);
  const result = [];
  let currentEnumName = '';
  let inEnum = false;
  let autoValue = 0;

  for (const line of lines) {
    const enumStartMatch = line.match(/^\s*(Public\s+|Private\s+)?Enum\s+(\w+)/i);
    if (enumStartMatch) {
      inEnum = true;
      currentEnumName = enumStartMatch[2];
      autoValue = 0;
      continue;
    }

    if (inEnum) {
      if (/^\s*End\s+Enum/i.test(line)) {
        inEnum = false;
        currentEnumName = '';
        continue;
      }

      const memberWithValueMatch = line.match(/^\s*(\w+)\s*=\s*(-?\d+)/);
      if (memberWithValueMatch) {
        const varName = `${currentEnumName}_${memberWithValueMatch[1]}`;
        const value = parseInt(memberWithValueMatch[2], 10);
        result.push({ varName, value });
        autoValue = value + 1;
        continue;
      }

      const memberOnlyMatch = line.match(/^\s*(\w+)\s*$/);
      if (memberOnlyMatch && memberOnlyMatch[1] !== '') {
        const varName = `${currentEnumName}_${memberOnlyMatch[1]}`;
        result.push({ varName, value: autoValue });
        autoValue++;
      }
    }
  }

  return result;
}
