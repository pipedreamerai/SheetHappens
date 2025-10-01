// Simple Node.js script to copy dist/ to docs/addin/
// Replaces rsync for Vercel compatibility

const fs = require('fs');
const path = require('path');

// Source and destination paths (relative to this script's location)
const srcDir = path.join(__dirname, '..', 'dist');
const destDir = path.join(__dirname, '..', '..', 'docs', 'addin');

// Recursively copy directory
function copyRecursive(src, dest) {
  // Create destination directory if it doesn't exist
  if (!fs.existsSync(dest)) {
    fs.mkdirSync(dest, { recursive: true });
  }

  // Read source directory
  const entries = fs.readdirSync(src, { withFileTypes: true });

  for (const entry of entries) {
    const srcPath = path.join(src, entry.name);
    const destPath = path.join(dest, entry.name);

    if (entry.isDirectory()) {
      // Recursively copy subdirectory
      copyRecursive(srcPath, destPath);
    } else {
      // Copy file
      fs.copyFileSync(srcPath, destPath);
    }
  }
}

// Remove destination directory contents first (equivalent to rsync --delete)
function removeDir(dir) {
  if (fs.existsSync(dir)) {
    fs.readdirSync(dir).forEach((file) => {
      const curPath = path.join(dir, file);
      if (fs.lstatSync(curPath).isDirectory()) {
        removeDir(curPath);
      } else {
        fs.unlinkSync(curPath);
      }
    });
    fs.rmdirSync(dir);
  }
}

console.log('Cleaning destination directory...');
removeDir(destDir);

console.log('Copying files from dist/ to docs/addin/...');
copyRecursive(srcDir, destDir);

console.log('Done! Files copied successfully.');

