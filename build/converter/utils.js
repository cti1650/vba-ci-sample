import { readFileSync, writeFileSync, existsSync, mkdirSync } from 'fs';
import { dirname, join, resolve, basename, extname } from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

/**
 * Get the build directory path
 */
export function getBuildDir() {
  return resolve(__dirname, '..');
}

/**
 * Get the rules directory path
 */
export function getRulesDir() {
  return join(getBuildDir(), 'rules');
}

/**
 * Get the mocks directory path
 */
export function getMocksDir() {
  return join(getBuildDir(), 'mocks');
}

/**
 * Read a file and return its content
 * @param {string} filePath - Path to the file
 * @param {string} encoding - File encoding (default: utf8)
 * @returns {string} File content
 * @throws {Error} If file cannot be read
 */
export function readFile(filePath, encoding = 'utf8') {
  try {
    return readFileSync(filePath, { encoding });
  } catch (error) {
    throw new Error(`Failed to read file: ${filePath}\n${error.message}`);
  }
}

/**
 * Write content to a file (BOM-less UTF-8)
 * @param {string} filePath - Path to the file
 * @param {string} content - Content to write
 * @throws {Error} If file cannot be written
 */
export function writeFile(filePath, content) {
  try {
    const dir = dirname(filePath);
    if (!existsSync(dir)) {
      mkdirSync(dir, { recursive: true });
    }
    writeFileSync(filePath, content, { encoding: 'utf8' });
  } catch (error) {
    throw new Error(`Failed to write file: ${filePath}\n${error.message}`);
  }
}

/**
 * Read and parse a JSON file
 * @param {string} filePath - Path to the JSON file
 * @returns {object} Parsed JSON object
 * @throws {Error} If file cannot be read or parsed
 */
export function readJson(filePath) {
  const content = readFile(filePath);
  try {
    return JSON.parse(content);
  } catch (error) {
    throw new Error(`Failed to parse JSON file: ${filePath}\n${error.message}`);
  }
}

/**
 * Load a rule file from the rules directory
 * @param {string} ruleName - Name of the rule file (without .json extension)
 * @returns {object} Parsed rule object
 * @throws {Error} If rule file cannot be loaded
 */
export function loadRule(ruleName) {
  const rulePath = join(getRulesDir(), `${ruleName}.json`);
  if (!existsSync(rulePath)) {
    throw new Error(`Rule file not found: ${ruleName}.json (expected at ${rulePath})`);
  }
  return readJson(rulePath);
}

/**
 * Load a mock definition from the mocks directory
 * @param {string} mockName - Name of the mock file (without .json extension)
 * @returns {object} Parsed mock definition
 * @throws {Error} If mock definition file cannot be loaded
 */
export function loadMockDefinition(mockName) {
  const mockPath = join(getMocksDir(), `${mockName}.json`);
  if (!existsSync(mockPath)) {
    throw new Error(`Mock definition not found: ${mockName}.json (expected at ${mockPath})`);
  }
  return readJson(mockPath);
}

/**
 * Load a VBS template file
 * @param {string} category - Template category (classes or helpers)
 * @param {string} templateName - Template file name (without .vbs extension)
 * @returns {string} Template content
 * @throws {Error} If template file cannot be loaded
 */
export function loadVbsTemplate(category, templateName) {
  const templatePath = join(getMocksDir(), category, `${templateName}.vbs`);
  if (!existsSync(templatePath)) {
    throw new Error(`VBS template not found: ${category}/${templateName}.vbs (expected at ${templatePath})`);
  }
  return readFile(templatePath);
}

/**
 * Check if a file exists
 * @param {string} filePath - Path to check
 * @returns {boolean} True if file exists
 */
export function fileExists(filePath) {
  return existsSync(filePath);
}

/**
 * Get file name without extension
 * @param {string} filePath - Path to the file
 * @returns {string} File name without extension
 */
export function getBaseName(filePath) {
  return basename(filePath, extname(filePath));
}

/**
 * Get file extension
 * @param {string} filePath - Path to the file
 * @returns {string} File extension (including dot)
 */
export function getExtension(filePath) {
  return extname(filePath).toLowerCase();
}

/**
 * Check if a file is a class file (.cls)
 * @param {string} filePath - Path to the file
 * @returns {boolean} True if file is a class
 */
export function isClassFile(filePath) {
  return getExtension(filePath) === '.cls';
}

/**
 * Split content into lines (handles both CRLF and LF)
 * @param {string} content - Content to split
 * @returns {string[]} Array of lines
 */
export function splitLines(content) {
  return content.split(/\r?\n/);
}

/**
 * Join lines with CRLF (Windows line endings for VBS)
 * @param {string[]} lines - Array of lines
 * @returns {string} Joined content
 */
export function joinLines(lines) {
  return lines.join('\r\n');
}

/**
 * Normalize line endings to CRLF
 * @param {string} content - Content to normalize
 * @returns {string} Content with CRLF line endings
 */
export function normalizeCrlf(content) {
  return content.replace(/\r?\n/g, '\r\n');
}
