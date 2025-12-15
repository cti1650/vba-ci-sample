#!/usr/bin/env node
import yargs from 'yargs';
import { hideBin } from 'yargs/helpers';
import { glob } from 'glob';
import { resolve, basename } from 'path';
import { existsSync, mkdirSync } from 'fs';

import { parseVbaFile, mergeEnumDefinitions, mergeApiDeclarations } from './parser.js';
import { transformVbaToVbs } from './transformer.js';
import { generateVbsFile, generateEnumsFile, generateSummary, printSummary } from './generator.js';
import { getMockedApiFunctions, generateVbaCompat } from './mock-generator.js';
import { getBaseName, isClassFile } from './utils.js';

/**
 * Main conversion function
 */
async function convert(inputDirs, outputDir, options = {}) {
  const { useMockCreateObject = false, generateMocks = true } = options;

  // Ensure output directory exists
  if (!existsSync(outputDir)) {
    mkdirSync(outputDir, { recursive: true });
  }

  // Step 1: Generate vba-compat.vbs if requested
  if (generateMocks) {
    await generateVbaCompat();
  }

  // Step 2: Collect all VBA files
  const allFiles = [];
  for (const dir of inputDirs) {
    const resolvedDir = resolve(dir);
    if (!existsSync(resolvedDir)) {
      console.warn(`[WARN] Directory not found: ${dir}`);
      continue;
    }

    const files = await glob('**/*.{bas,cls}', { cwd: resolvedDir, absolute: true });
    allFiles.push(...files);
  }

  if (allFiles.length === 0) {
    console.warn('[WARN] No VBA files found');
    return;
  }

  // Step 3: Parse all files and collect metadata
  const parsedFiles = allFiles.map(parseVbaFile);
  const allEnums = mergeEnumDefinitions(parsedFiles.map(f => f.enums));
  const allApis = mergeApiDeclarations(parsedFiles.map(f => f.apiDeclarations));
  const mockedApis = getMockedApiFunctions();

  // Step 4: Generate _enums.vbs
  const enumsPath = generateEnumsFile(outputDir, allEnums);
  if (enumsPath) {
    console.log(`[GENERATED] _enums.vbs`);
  }

  // Step 5: Transform and generate each file
  const convertedFiles = [];

  for (const parsed of parsedFiles) {
    const className = getBaseName(parsed.filePath);

    const transformed = transformVbaToVbs({
      content: parsed.content,
      isClass: parsed.isClass,
      className,
      allEnums,
      allApis,
      mockedApis,
      useMockCreateObject,
    });

    const outputPath = generateVbsFile(outputDir, className, transformed);
    convertedFiles.push(`${basename(parsed.filePath)} -> ${basename(outputPath)}`);
  }

  // Step 6: Print summary
  const summary = generateSummary({
    convertedFiles,
    allEnums,
    allApis,
    mockedApis,
  });

  printSummary(summary);
}

// CLI setup
const argv = yargs(hideBin(process.argv))
  .usage('Usage: $0 --input <dirs...> --output <dir>')
  .option('input', {
    alias: 'i',
    type: 'array',
    description: 'Input directories containing VBA files',
    demandOption: true,
  })
  .option('output', {
    alias: 'o',
    type: 'string',
    description: 'Output directory for VBS files',
    demandOption: true,
  })
  .option('mock-create-object', {
    type: 'boolean',
    description: 'Convert CreateObject calls to CreateObjectMock',
    default: false,
  })
  .option('skip-mocks', {
    type: 'boolean',
    description: 'Skip vba-compat.vbs generation',
    default: false,
  })
  .help()
  .alias('help', 'h')
  .parseSync();

// Run conversion
convert(argv.input, argv.output, {
  useMockCreateObject: argv.mockCreateObject,
  generateMocks: !argv.skipMocks,
}).catch((err) => {
  console.error('Error:', err.message);
  process.exit(1);
});
