import { loadRule, splitLines, joinLines } from './utils.js';
import { collectEnumDefinitions } from './parser.js';

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

  // Step 1.5: Remove comment-only lines
  lines = removeCommentOnlyLines(lines);

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
  let currentBlock = null; // { start: pattern, end: pattern }
  let nestDepth = 0;

  for (const line of lines) {
    // Check if we're in a skip block
    if (currentBlock) {
      // Check for nested block start (same type)
      if (new RegExp(currentBlock.start, 'i').test(line)) {
        nestDepth++;
      }
      // Check for block end
      else if (new RegExp(currentBlock.end, 'i').test(line)) {
        if (nestDepth > 0) {
          nestDepth--;
        } else {
          currentBlock = null;
        }
      }
      continue;
    }

    // Check for block start
    let shouldSkipBlock = false;
    for (const block of skipRules.blocks) {
      if (new RegExp(block.start, 'i').test(line)) {
        currentBlock = block;
        nestDepth = 0;
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
 * Check if a position in a line is inside a string literal or comment
 * @param {string} line - The line to check
 * @param {number} position - Position in the line
 * @returns {boolean} True if position is inside string or comment
 */
function isInsideStringOrComment(line, position) {
  // Check if the position is after a comment marker
  const commentPos = line.indexOf("'");
  if (commentPos !== -1 && position > commentPos) {
    return true;
  }

  // Check if the position is inside a string literal
  let inString = false;
  for (let i = 0; i < position && i < line.length; i++) {
    if (line[i] === '"') {
      inString = !inString;
    }
  }

  return inString;
}

/**
 * Apply enum reference conversions (safe: skips strings and comments)
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
      result = safeReplace(result, dotPattern, `${enumName}_${memberName}`);
    }
  }

  // Convert standalone enum member references (for same-file enums)
  // This handles cases like: = MemberName -> = EnumName_MemberName
  for (const [enumName, members] of allEnums) {
    for (const [memberName] of members) {
      // Only in assignment context to avoid false positives
      const assignPattern = new RegExp(`=\\s*\\b${memberName}\\b(?!\\s*[.(])`, 'g');
      result = safeReplace(result, assignPattern, `= ${enumName}_${memberName}`, true);
    }
  }

  return result;
}

/**
 * Replace pattern only if not inside string or comment
 * @param {string} line - Input line
 * @param {RegExp} pattern - Pattern to match
 * @param {string} replacement - Replacement string
 * @param {boolean} preserveSpacing - If true, preserves original spacing around =
 * @returns {string} Transformed line
 */
function safeReplace(line, pattern, replacement, preserveSpacing = false) {
  let result = line;
  let match;
  const regex = new RegExp(pattern.source, 'g');

  // Find all matches and filter out those in strings/comments
  const matches = [];
  while ((match = regex.exec(line)) !== null) {
    if (!isInsideStringOrComment(line, match.index)) {
      matches.push({
        index: match.index,
        length: match[0].length,
        match: match[0],
      });
    }
  }

  // Replace from end to start to preserve indices
  for (let i = matches.length - 1; i >= 0; i--) {
    const m = matches[i];
    let replaceWith = replacement;

    // For assignment patterns, preserve original spacing
    if (preserveSpacing && m.match.startsWith('=')) {
      const originalSpacing = m.match.match(/^=(\s*)/)[1];
      replaceWith = '=' + originalSpacing + replacement.replace(/^=\s*/, '');
    }

    result = result.slice(0, m.index) + replaceWith + result.slice(m.index + m.length);
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
 * Remove lines that contain only comments (no actual code)
 * @param {string[]} lines - Input lines
 * @returns {string[]} Lines with comment-only lines removed
 */
export function removeCommentOnlyLines(lines) {
  return lines.filter((line) => {
    const trimmed = line.trim();
    // Keep empty lines (for readability)
    if (trimmed === '') {
      return true;
    }
    // Remove lines that start with a comment marker
    if (trimmed.startsWith("'")) {
      return false;
    }
    // Remove lines that start with Rem (VBA comment keyword)
    if (/^Rem\s/i.test(trimmed)) {
      return false;
    }
    return true;
  });
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
 * Uses collectEnumDefinitions from parser.js and converts to array format
 * @param {string} content - VBA content
 * @returns {Array<{varName: string, value: number}>} Array of enum variable definitions
 */
export function collectLocalEnumDefs(content) {
  const lines = splitLines(content);
  const enumMap = collectEnumDefinitions(lines);

  // Convert Map format to array format for backward compatibility
  const result = [];
  for (const [enumName, members] of enumMap) {
    for (const [memberName, value] of members) {
      result.push({ varName: `${enumName}_${memberName}`, value });
    }
  }

  return result;
}
