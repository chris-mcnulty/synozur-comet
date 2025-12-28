# Quick Fix: Node.js Not Found

## Immediate Solution

**Close PowerShell completely and open a new window.** This refreshes the PATH environment variable.

Then try:
```powershell
node --version
```

## If That Doesn't Work

### Option 1: Install Node.js 18 LTS (Recommended for SPFx)

1. Download Node.js 18 LTS from: https://nodejs.org/en/download/
2. Choose the **Windows Installer (.msi)** for 64-bit
3. Run the installer (it will add Node.js to PATH automatically)
4. **Close and reopen PowerShell**
5. Verify: `node --version` (should show v18.x.x)

### Option 2: Install nvm-windows (Better for managing multiple versions)

1. Download nvm-windows: https://github.com/coreybutler/nvm-windows/releases
2. Install `nvm-setup.exe`
3. **Close and reopen PowerShell**
4. Install Node.js 18:
   ```powershell
   nvm install 18.20.4
   nvm use 18.20.4
   ```
5. Verify: `node --version`

## After Node.js is Working

1. Navigate to your project:
   ```powershell
   cd C:\Users\cmcnu\OneDrive\Documents\Cursor\Comet\synozur-comet
   ```

2. Verify dependencies are installed:
   ```powershell
   npm install
   ```

3. Run gulp:
   ```powershell
   gulp clean
   gulp bundle --ship
   ```

## Why This Happens

- PATH environment variable needs to be refreshed in PowerShell
- Node.js might have been uninstalled or moved
- PowerShell session was opened before Node.js was installed

**The simplest fix is usually to just close and reopen PowerShell!**

