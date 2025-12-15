import { writeFile, joinLines, normalizeCrlf } from './utils.js';
import { join } from 'path';

/**
 * Generate VBS output file from transformed content
 * @param {string} outputDir - Output directory path
 * @param {string} fileName - Output file name (without extension)
 * @param {string} content - Transformed VBS content
 */
export function generateVbsFile(outputDir, fileName, content) {
  const outputPath = join(outputDir, `${fileName}.vbs`);
  const normalizedContent = normalizeCrlf(content);
  writeFile(outputPath, normalizedContent);
  return outputPath;
}

/**
 * Generate _enums.vbs file with all enum definitions
 * @param {string} outputDir - Output directory path
 * @param {Map<string, Map<string, number>>} allEnums - All enum definitions
 * @returns {string|null} Output path or null if no enums
 */
export function generateEnumsFile(outputDir, allEnums) {
  if (allEnums.size === 0) {
    return null;
  }

  const lines = ["' Auto-generated Enum constants"];

  for (const [enumName, members] of allEnums) {
    for (const [memberName, value] of members) {
      lines.push(`${enumName}_${memberName} = ${value}`);
    }
  }

  const outputPath = join(outputDir, '_enums.vbs');
  const content = normalizeCrlf(joinLines(lines));
  writeFile(outputPath, content);

  return outputPath;
}

/**
 * Generate conversion summary
 * @param {object} options - Summary options
 * @param {string[]} options.convertedFiles - List of converted file paths
 * @param {Map<string, Map<string, number>>} options.allEnums - All enum definitions
 * @param {Map<string, object>} options.allApis - All API declarations
 * @param {Set<string>} options.mockedApis - Set of mocked API function names
 * @returns {object} Summary object
 */
export function generateSummary(options) {
  const { convertedFiles, allEnums, allApis, mockedApis } = options;

  const summary = {
    convertedCount: convertedFiles.length,
    convertedFiles,
    enums: [],
    apis: {
      mocked: [],
      unmocked: [],
    },
  };

  // Enum summary
  for (const [enumName, members] of allEnums) {
    summary.enums.push({
      name: enumName,
      members: Array.from(members.keys()),
    });
  }

  // API summary
  for (const [apiName, info] of allApis) {
    if (mockedApis.has(apiName)) {
      summary.apis.mocked.push({ name: apiName, lib: info.lib });
    } else {
      summary.apis.unmocked.push({ name: apiName, lib: info.lib });
    }
  }

  return summary;
}

/**
 * Print conversion summary to console
 * @param {object} summary - Summary object from generateSummary
 */
export function printSummary(summary) {
  console.log('=========================================');
  console.log('VBA to VBS Converter');
  console.log('=========================================');
  console.log('');

  if (summary.enums.length > 0) {
    console.log('[INFO] Collected Enums:');
    for (const e of summary.enums) {
      console.log(`  - ${e.name}: ${e.members.join(', ')}`);
    }
    console.log('');
  }

  if (summary.apis.mocked.length > 0 || summary.apis.unmocked.length > 0) {
    console.log('[INFO] Collected API Declarations:');
    for (const api of summary.apis.mocked) {
      console.log(`  - ${api.name} (${api.lib}) [MOCKED]`);
    }
    for (const api of summary.apis.unmocked) {
      console.log(`  - ${api.name} (${api.lib}) [NOT MOCKED - will use default return]`);
    }
    console.log('');
    console.log(`[INFO] API Summary: ${summary.apis.mocked.length} mocked, ${summary.apis.unmocked.length} not mocked`);
    console.log('');
  }

  for (const file of summary.convertedFiles) {
    console.log(`[CONVERTED] ${file}`);
  }

  console.log('');
  console.log('=========================================');
  console.log(`Converted ${summary.convertedCount} file(s)`);
  console.log('=========================================');
}
