# Node.js Version Compatibility Guide

## Issue
SPFx 1.18.2 requires Node.js 18.17.1 - 18.x.x, but you're currently using Node.js v22.14.0.

## Solution 1: Use Node Version Manager (Recommended)

### For Windows (nvm-windows):

1. **Download and Install nvm-windows:**
   - Download from: https://github.com/coreybutler/nvm-windows/releases
   - Install the latest `nvm-setup.exe`

2. **After installation, open a NEW PowerShell/Command Prompt** and run:
   ```powershell
   nvm install 18.20.4
   nvm use 18.20.4
   ```

3. **Verify the version:**
   ```powershell
   node -v
   ```
   Should show: `v18.20.4`

4. **Now you can run gulp commands:**
   ```powershell
   npm install
   gulp clean
   gulp bundle --ship
   ```

### Switch between Node versions:
- Use Node 18 for SPFx: `nvm use 18.20.4`
- Use Node 22 for other projects: `nvm use 22.14.0`

## Solution 2: Update to SPFx 1.22.1 (May Support Newer Node)

If you want to try a newer SPFx version, we can update the project. However, this requires updating all dependencies and may have breaking changes.

## Quick Check

To check which Node versions are compatible with SPFx versions:
- SPFx 1.18.2: Node 18.17.1 - 18.x.x
- SPFx 1.19.0: Node 18.17.1 - 18.x.x  
- SPFx 1.20.0+: Node 18.x or 20.x (check specific version requirements)

