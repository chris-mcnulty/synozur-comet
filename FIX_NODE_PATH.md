# Fix Node.js PATH Issue

## Problem
After installing/switching Node.js versions, PowerShell can't find `node.exe`.

## Solution 1: Close and Reopen PowerShell (Easiest)

1. **Close the current PowerShell window completely**
2. **Open a new PowerShell window**
3. **Run:**
   ```powershell
   node --version
   gulp clean
   ```

When you switch Node versions with nvm, you need to restart PowerShell for the PATH changes to take effect.

## Solution 2: If Using nvm-windows

1. **Check installed versions:**
   ```powershell
   nvm list
   ```

2. **Use a specific version:**
   ```powershell
   nvm use 18.20.4
   ```

3. **If Node.js still not found, refresh PATH manually:**
   ```powershell
   $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
   ```

## Solution 3: Verify Node.js Installation

If you're not using nvm, verify Node.js is installed:

1. Check if Node.js is in Program Files:
   ```powershell
   Test-Path "C:\Program Files\nodejs\node.exe"
   ```

2. If it exists, add it to PATH for this session:
   ```powershell
   $env:Path += ";C:\Program Files\nodejs"
   ```

3. Verify:
   ```powershell
   node --version
   ```

## Recommended Action

**Simply close PowerShell and open a new window** - this is the quickest fix after installing or switching Node.js versions.

