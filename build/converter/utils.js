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
 */
export function readFile(filePath, encoding = 'utf8') {
  return readFileSync(filePath, { encoding });
}

/**
 * Write content to a file (BOM-less UTF-8)
 * @param {string} filePath - Path to the file
 * @param {string} content - Content to write
 */
export function writeFile(filePath, content) {
  const dir = dirname(filePath);
  if (!existsSync(dir)) {
    mkdirSync(dir, { recursive: true });
  }
  writeFileSync(filePath, content, { encoding: 'utf8' });
}

/**
 * Read and parse a JSON file
 * @param {string} filePath - Path to the JSON file
 * @returns {object} Parsed JSON object
 */
export function readJson(filePath) {
  const content = readFile(filePath);
  return JSON.parse(content);
}

/**
 * Load a rule file from the rules directory
 * @param {string} ruleName - Name of the rule file (without .json extension)
 * @returns {object} Parsed rule object
 */
export function loadRule(ruleName) {
  const rulePath = join(getRulesDir(), `${ruleName}.json`);
  return readJson(rulePath);
}

/**
 * Load a mock definition from the mocks directory
 * @param {string} mockName - Name of the mock file (without .json extension)
 * @returns {object} Parsed mock definition
 */
export function loadMockDefinition(mockName) {
  const mockPath = join(getMocksDir(), `${mockName}.json`);
  return readJson(mockPath);
}

/**
 * Load a VBS template file
 * @param {string} category - Template category (classes or helpers)
 * @param {string} templateName - Template file name (without .vbs extension)
 * @returns {string} Template content
 */
export function loadVbsTemplate(category, templateName) {
  const templatePath = join(getMocksDir(), category, `${templateName}.vbs`);
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
