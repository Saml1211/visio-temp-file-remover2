const express = require('express');
const path = require('path');
const fs = require('fs');
const { exec } = require('child_process');
const chalk = require('chalk');
const os = require('os');

// Function to resolve PowerShell executable path
function getPowerShellExecutable() {
  // Try to find PowerShell executable
  const powerShellPaths = [
    'powershell.exe',
    'pwsh.exe',
    'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe',
    'C:\\Windows\\SysWOW64\\WindowsPowerShell\\v1.0\\powershell.exe'
  ];
  
  for (const psPath of powerShellPaths) {
    try {
      // Use 'where' command on Windows to check if executable exists
      const { execSync } = require('child_process');
      const result = execSync(`where "${psPath}"`, { stdio: 'pipe' });
      if (result && result.toString().trim()) {
        return psPath;
      }
    } catch (error) {
      // Continue to next path
    }
  }
  
  // Fallback to 'powershell'
  return 'powershell';
}

// Security function to validate file paths
function isValidFilePath(filePath, allowedRoots) {
  try {
    // Resolve the file path to handle relative paths and prevent directory traversal
    const resolvedPath = path.resolve(filePath);
    
    // Check if the file path is under any of the allowed roots
    for (const root of allowedRoots) {
      const resolvedRoot = path.resolve(root);
      // Ensure the resolved path starts with the resolved root
      if (resolvedPath.startsWith(resolvedRoot)) {
        return true;
      }
    }
    
    return false;
  } catch (error) {
    log(LOG_LEVELS.ERROR, CATEGORIES.DELETE, `Error validating file path: ${error.message}`);
    return false;
  }
}

// Add a function to get allowed root directories (can be extended with config)
function getAllowedRoots() {
  // For now, allow common Visio template directories
  // This could be expanded to read from config file
  return [
    'Z:\\ENGINEERING TEMPLATES\\VISIO SHAPES 2025',
    'C:\\Users',
    'C:\\Program Files',
    'C:\\Program Files (x86)'
  ];
}
const LOG_LEVELS = {
  INFO: 'INFO',
  SUCCESS: 'SUCCESS',
  WARN: 'WARN',
  ERROR: 'ERROR',
  DEBUG: 'DEBUG', // For developers
  DETAIL: 'DETAIL', // For verbose operational steps
};

const CATEGORIES = {
  CONFIG: 'CONFIG',
  SERVER: 'SERVER',
  SCAN: 'SCAN',
  DELETE: 'DELETE',
  POWERSHELL: 'PS_CMD',
  PARSING: 'PARSING',
  HTTP: 'HTTP',
  GENERAL: 'GENERAL',
};

const ICONS = {
  INFO: 'ℹ️',
  SUCCESS: '✅',
  WARN: '⚠️',
  ERROR: '❌',
  DEBUG: '⚙️',
  DETAIL: '🔹', // Using DETAIL for "working" or detailed steps
  POWERSHELL: '>',
};

function getTimestamp() {
  return new Date().toLocaleTimeString();
}

function log(level, category, message, ...args) {
  const timestamp = chalk.dim(`[${getTimestamp()}]`);
  let coloredMessage = message;
  let icon = '';
  let categoryString = `[${category.padEnd(7, ' ')}]`;

  switch (level) {
    case LOG_LEVELS.INFO:
      icon = ICONS.INFO;
      coloredMessage = chalk.blue(message);
      categoryString = chalk.bold.blue(categoryString);
      break;
    case LOG_LEVELS.SUCCESS:
      icon = ICONS.SUCCESS;
      coloredMessage = chalk.bold.green(message);
      categoryString = chalk.bold.green(categoryString);
      break;
    case LOG_LEVELS.WARN:
      icon = ICONS.WARN;
      coloredMessage = chalk.yellow(message);
      categoryString = chalk.bold.yellow(categoryString);
      break;
    case LOG_LEVELS.ERROR:
      icon = ICONS.ERROR;
      coloredMessage = chalk.bold.red(message);
      categoryString = chalk.bold.red(categoryString);
      break;
    case LOG_LEVELS.DEBUG:
      icon = ICONS.DEBUG;
      coloredMessage = chalk.dim.cyan(message);
      categoryString = chalk.bold.cyan(categoryString);
      break;
    case LOG_LEVELS.DETAIL:
      icon = ICONS.DETAIL;
      coloredMessage = chalk.dim(message); // Keep details subtle
      // Category color for DETAIL will depend on the category itself
      break;
    default:
      icon = '';
  }
  
  // Specific category styling overrides/enhancements
  switch (category) {
    case CATEGORIES.CONFIG:
      categoryString = chalk.bold.blue(categoryString);
      if (level === LOG_LEVELS.DETAIL) coloredMessage = chalk.dim.blue(message);
      break;
    case CATEGORIES.SERVER:
      categoryString = chalk.bold.green(categoryString);
      if (level === LOG_LEVELS.INFO) coloredMessage = chalk.blue(message); // Server info can be blue
      if (level === LOG_LEVELS.SUCCESS) coloredMessage = chalk.bold.green(message);
      break;
    case CATEGORIES.SCAN:
      categoryString = chalk.bold.yellow(categoryString);
      if (level === LOG_LEVELS.DETAIL) coloredMessage = chalk.dim.yellow(message);
      else if (level !== LOG_LEVELS.WARN && level !== LOG_LEVELS.ERROR) coloredMessage = chalk.yellow(message);
      break;
    case CATEGORIES.DELETE:
      categoryString = chalk.bold.magenta(categoryString);
      if (level === LOG_LEVELS.DETAIL) coloredMessage = chalk.dim.magenta(message);
      else if (level !== LOG_LEVELS.WARN && level !== LOG_LEVELS.ERROR) coloredMessage = chalk.magenta(message);
      break;
    case CATEGORIES.POWERSHELL:
      icon = ICONS.POWERSHELL; // Ensure PS icon is always used for this category
      categoryString = chalk.bold.cyan(categoryString);
      coloredMessage = chalk.dim.cyan(message); // PS commands are always details
      break;
    case CATEGORIES.PARSING:
      categoryString = chalk.bold.yellow(categoryString); // Often associated with warnings or errors
      // Keep message color as per level (e.g. red for error)
      break;
    case CATEGORIES.HTTP:
      categoryString = chalk.bold.cyan(categoryString);
      if (level === LOG_LEVELS.DETAIL) coloredMessage = chalk.dim.cyan(message);
      else if (level !== LOG_LEVELS.WARN && level !== LOG_LEVELS.ERROR) coloredMessage = chalk.cyan(message);
      break;
    default:
      // For GENERAL or unstyled categories, ensure the level color applies if not detailed
      if (level !== LOG_LEVELS.DETAIL) {
         // Re-apply level default color if not handled by category specific
        switch (level) {
            case LOG_LEVELS.INFO: coloredMessage = chalk.blue(message); break;
            case LOG_LEVELS.SUCCESS: coloredMessage = chalk.bold.green(message); break;
            case LOG_LEVELS.WARN: coloredMessage = chalk.yellow(message); break;
            case LOG_LEVELS.ERROR: coloredMessage = chalk.bold.red(message); break;
            case LOG_LEVELS.DEBUG: coloredMessage = chalk.dim.cyan(message); break;
        }
      }
      categoryString = chalk.dim(categoryString);
  }

  const finalMessage = `${timestamp} ${icon} ${categoryString} ${coloredMessage}`;
  
  const formattedArgs = args.map(arg => {
    if (typeof arg === 'object') {
      try {
        return chalk.dim(`\n${JSON.stringify(arg, null, 2)}`);
      } catch (e) {
        return chalk.dim(`\n${String(arg)}`);
      }
    }
    // Apply the base color of the message to its arguments, if not an object
    // This logic might need to be smarter if args need their own colors
    if (level === LOG_LEVELS.ERROR) return chalk.bold.red(String(arg));
    if (level === LOG_LEVELS.WARN) return chalk.yellow(String(arg));
    return chalk.dim(String(arg)); // Default for other args
  });

  if (level === LOG_LEVELS.ERROR) {
    console.error(finalMessage, ...formattedArgs);
  } else if (level === LOG_LEVELS.WARN) {
    console.warn(finalMessage, ...formattedArgs);
  } else {
    console.log(finalMessage, ...formattedArgs);
  }
}

function logSeparator(character = '─', length = 60) {
  console.log(chalk.dim.white(character.repeat(length)));
}

// --- End Logger Utility ---

// Function to get local IP addresses
function getLocalIPs() {
  const interfaces = os.networkInterfaces();
  const ips = [];
  
  for (const interfaceName in interfaces) {
    const iface = interfaces[interfaceName];
    for (const connection of iface) {
      // Skip internal (localhost) and non-IPv4 addresses
      if (!connection.internal && connection.family === 'IPv4') {
        ips.push(connection.address);
      }
    }
  }
  
  return ips;
}

const app = express();
const PORT = process.env.PORT || 3000;
const HOST = process.env.HOST || 'localhost'; // Default to localhost for security
const NODE_ENV = process.env.NODE_ENV || 'development';
const DEFAULT_SCAN_DIR_DISPLAY = 'Z:\\\\ENGINEERING TEMPLATES\\\\VISIO SHAPES 2025'; // For display

// Middleware
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Serve the main page
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'views', 'index.html'));
});

// API to scan for files
app.post('/api/scan', (req, res) => {
  const targetDir = req.body.directory || DEFAULT_SCAN_DIR_DISPLAY;
  
  log(LOG_LEVELS.INFO, CATEGORIES.SCAN, `Request to scan directory: ${targetDir}`);
  
  const escapedPath = targetDir.replace(/'/g, "''");
  const powershellExecutable = getPowerShellExecutable();
  const powershellCommand = `Get-ChildItem -Path '${escapedPath}' -Recurse -File -Force -Include "~\\$*.vssx","~\\$*.vsdx","~\\$*.vstx","~\\$*.vsdm","~\\$*.vsd" | Select-Object -Property FullName,Name | ConvertTo-Json`;
  
  log(LOG_LEVELS.DETAIL, CATEGORIES.POWERSHELL, `Executing: ${powershellCommand}`);
  
  exec(`${powershellExecutable} -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "${powershellCommand}"`, { maxBuffer: 1024 * 1024 }, (error, stdout, stderr) => {
    if (error) {
      log(LOG_LEVELS.ERROR, CATEGORIES.SCAN, `Error scanning files: ${error.message}`);
      return res.status(500).json({ 
        error: error.message,
        details: 'Error executing PowerShell command to scan for files'
      });
    }
    
    if (stderr && stderr.trim() !== '') {
      log(LOG_LEVELS.WARN, CATEGORIES.SCAN, `Warning during file scan: ${stderr.trim()}`);
    }
    
    if (!stdout || stdout.trim() === '') {
      log(LOG_LEVELS.INFO, CATEGORIES.SCAN, `No files found in directory: ${targetDir}`);
      return res.json({ files: [], message: 'No matching files found' });
    }
    
    try {
      log(LOG_LEVELS.DEBUG, CATEGORIES.POWERSHELL, `Raw output (first 200 chars): ${stdout.substring(0, 200)}${stdout.length > 200 ? '...' : ''}`);
      const files = JSON.parse(stdout);
      
      if (!files || (files && Object.keys(files).length === 0)) {
        log(LOG_LEVELS.INFO, CATEGORIES.SCAN, `No files found after parsing JSON.`);
        return res.json({ files: [], message: 'No matching files found' });
      }
      
      const fileArray = Array.isArray(files) ? files : [files];
      
      const processedFiles = fileArray.map(file => {
        if (!file.Name && file.FullName) {
          file.Name = file.FullName.split('\\\\').pop().split('/').pop();
        }
        return file;
      });
      
      log(LOG_LEVELS.SUCCESS, CATEGORIES.SCAN, `Found ${processedFiles.length} files matching the pattern.`);
      res.json({ 
        files: processedFiles,
        message: `Found ${processedFiles.length} file(s)`,
        scannedDirectory: targetDir
      });
    } catch (e) {
      log(LOG_LEVELS.ERROR, CATEGORIES.PARSING, `Error parsing PowerShell output: ${e.message}`);
      log(LOG_LEVELS.DEBUG, CATEGORIES.POWERSHELL, `Raw output on parsing error (first 200 chars): ${stdout.substring(0, 200)}${stdout.length > 200 ? '...' : ''}`);
      res.status(500).json({ 
        error: `Failed to parse file list: ${e.message}`, 
        details: 'The PowerShell command executed successfully but returned invalid JSON',
        files: [] 
      });
    }
  });
});

// API to delete files
app.post('/api/delete', (req, res) => {
  const filesToDelete = req.body.files;
  
  if (!filesToDelete || !Array.isArray(filesToDelete) || filesToDelete.length === 0) {
    log(LOG_LEVELS.WARN, CATEGORIES.DELETE, `No files specified for deletion in request.`);
    return res.status(400).json({ 
      error: 'No files specified for deletion',
      details: 'The request must include a "files" array with at least one file path'
    });
  }
  
  log(LOG_LEVELS.INFO, CATEGORIES.DELETE, `Attempting to delete ${filesToDelete.length} files.`);
  
  const invalidFiles = filesToDelete.filter(file => typeof file !== 'string' || file.trim() === '');
  if (invalidFiles.length > 0) {
    log(LOG_LEVELS.ERROR, CATEGORIES.DELETE, `Invalid file paths in request: ${invalidFiles.length} invalid entries.`);
    return res.status(400).json({
      error: 'Invalid file paths provided',
      details: 'All file paths must be non-empty strings',
      invalidCount: invalidFiles.length
    });
  }
  
  // Validate file paths for security
  const allowedRoots = getAllowedRoots();
  const unauthorizedFiles = filesToDelete.filter(file => !isValidFilePath(file, allowedRoots));
  if (unauthorizedFiles.length > 0) {
    log(LOG_LEVELS.ERROR, CATEGORIES.DELETE, `Unauthorized file paths in request: ${unauthorizedFiles.length} files outside allowed roots.`);
    return res.status(403).json({
      error: 'Unauthorized file paths',
      details: 'All file paths must be within allowed directories for security',
      unauthorizedCount: unauthorizedFiles.length,
      allowedRoots: allowedRoots
    });
  }
  
  const fileListString = filesToDelete.map(file => `'${file.replace(/'/g, "''")}'`).join(',');
  
  const powershellExecutable = getPowerShellExecutable();
  const powershellCommand = `@(${fileListString}) | ForEach-Object { Remove-Item -Path $_ -Force }`;
  
  log(LOG_LEVELS.DETAIL, CATEGORIES.POWERSHELL, `Executing (truncated): ${powershellCommand.substring(0, 100)}...`);
  
  exec(`${powershellExecutable} -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command "${powershellCommand}"`, { maxBuffer: 1024 * 1024 }, (error, stdout, stderr) => {
    if (error) {
      log(LOG_LEVELS.ERROR, CATEGORIES.DELETE, `Error deleting files: ${error.message}`);
      return res.status(500).json({ 
        error: error.message,
        details: 'Error executing PowerShell command to delete files',
        filesAttempted: filesToDelete.length
      });
    }
    
    if (stderr && stderr.trim() !== '') {
      log(LOG_LEVELS.WARN, CATEGORIES.DELETE, `Issues during file deletion: ${stderr.trim()}`);
      
      return res.status(207).json({
        partialSuccess: true,
        message: `Some files may not have been deleted successfully`,
        details: stderr,
        filesAttempted: filesToDelete.length
      });
    }
    
    log(LOG_LEVELS.SUCCESS, CATEGORIES.DELETE, `Successfully processed delete command for ${filesToDelete.length} files.`);
    if (stdout && stdout.trim() !== '') {
      log(LOG_LEVELS.DEBUG, CATEGORIES.POWERSHELL, `Output (first 200 chars): ${stdout.substring(0, 200)}${stdout.length > 200 ? '...' : ''}`);
    }
    
    res.json({ 
      success: true, 
      message: `${filesToDelete.length} files deleted successfully`
    });
  });
});

// Middleware for handling errors
app.use((err, req, res, next) => {
  log(LOG_LEVELS.ERROR, CATEGORIES.SERVER, `Unhandled exception on path ${req.path}: ${err.message}`);
  if (err.stack) {
    log(LOG_LEVELS.DEBUG, CATEGORIES.SERVER, `Stack trace:\\n${err.stack}`);
  }
  res.status(500).json({
    error: 'Internal server error',
    message: err.message,
    path: req.path
  });
});

// Start the server
app.listen(PORT, HOST, () => {
  const serverName = "Visio Temp File Remover Server";
  const version = require('./package.json').version;
  const line = `═`.repeat(serverName.length + version.length + 8);

  console.log(chalk.bold.cyan(`\n╔${line}╗`));
  console.log(chalk.bold.cyan(`║  ${serverName} v${version}  ║`));
  console.log(chalk.bold.cyan(`╚${line}╝\n`));
  
  log(LOG_LEVELS.INFO, CATEGORIES.CONFIG, `Environment: ${NODE_ENV}`);
  log(LOG_LEVELS.INFO, CATEGORIES.CONFIG, `Default scan directory: ${DEFAULT_SCAN_DIR_DISPLAY}`);
  logSeparator();
  
  // Show all available access URLs
  log(LOG_LEVELS.SUCCESS, CATEGORIES.SERVER, `Server running on port ${PORT} and accessible at:`);
  log(LOG_LEVELS.INFO, CATEGORIES.SERVER, `  • Local:    http://localhost:${PORT}`);
  log(LOG_LEVELS.INFO, CATEGORIES.SERVER, `  • Local:    http://127.0.0.1:${PORT}`);
  
  const localIPs = getLocalIPs();
  if (localIPs.length > 0) {
    localIPs.forEach(ip => {
      log(LOG_LEVELS.INFO, CATEGORIES.SERVER, `  • Network:  http://${ip}:${PORT}`);
    });
  } else {
    log(LOG_LEVELS.WARN, CATEGORIES.SERVER, `  • Network:  No local network IPs detected`);
  }
  
  logSeparator();
  log(LOG_LEVELS.INFO, CATEGORIES.SERVER, `API endpoints available at all above URLs:`);
  log(LOG_LEVELS.INFO, CATEGORIES.SERVER, `  - POST /api/scan`);
  log(LOG_LEVELS.INFO, CATEGORIES.SERVER, `  - POST /api/delete`);
  logSeparator();
  log(LOG_LEVELS.INFO, CATEGORIES.SERVER, `Press Ctrl+C to stop the server.`);
});
