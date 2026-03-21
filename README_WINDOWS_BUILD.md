# Building the Windows Installer

## Why not directly on your Linux server?

Cross-compiling Qt6 apps for Windows on Linux requires:
- MinGW-w64 cross-compiler (~200 MB)
- Qt 6.5.3 Windows binaries (~1.5 GB)

Your build server has restricted outbound internet and cannot download these.
The solution below uses **GitHub Actions** — a free Windows VM provided by GitHub
that builds your installer automatically on every push.

---

## Step-by-step: GitHub Actions (Free, ~8 minutes per build)

### 1. Install git on your server (if not present)
```bash
apt-get install -y git
```

### 2. Initialise the repo and push to GitHub
```bash
cd ~/QtSpreadsheet
git init
git add .
git commit -m "Initial commit"
```

Go to https://github.com/new and create a **new empty repository** (no README).
Then push:
```bash
git remote add origin https://github.com/YOUR_USERNAME/QtSpreadsheet.git
git branch -M main
git push -u origin main
```

### 3. Watch the build
- Go to your repo on GitHub
- Click the **Actions** tab
- The "Windows Installer" workflow starts automatically
- Wait ~8 minutes

### 4. Download your installer
- Click the completed workflow run
- Under **Artifacts**, download `QtSpreadsheet-Windows-Installer`
- Unzip it — you have `QtSpreadsheet-Setup.exe`

### 5. What the installer does on Windows
- Installs to `C:\Program Files\QtSpreadsheet\`
- Creates a Start Menu shortcut
- Creates a Desktop shortcut
- Registers `.xlsx` file association
- Appears in "Add or Remove Programs" with an uninstaller

---

## Triggering a build manually (without a push)

In GitHub → Actions → "Windows Installer" → click **"Run workflow"** button.

---

## Changing the app version

Edit `installer/QtSpreadsheet.nsi`, line:
```
!define APP_VERSION   "1.0.0"
```

---

## If you DO get MinGW access later (native cross-compile)

```bash
# Install tools
apt-get install mingw-w64 nsis cmake

# Download Qt 6.5.3 for Windows (MinGW build) from:
# https://download.qt.io/official_releases/qt/6.5/6.5.3/

# Build
mkdir build-win && cd build-win
cmake .. \
  -DCMAKE_TOOLCHAIN_FILE=../cmake/mingw-w64-x86_64.cmake \
  -DCMAKE_PREFIX_PATH=/path/to/Qt/6.5.3/mingw_64

cmake --build . --parallel

# Deploy + package
/path/to/Qt/6.5.3/mingw_64/bin/windeployqt QtSpreadsheet.exe
makensis ../installer/QtSpreadsheet.nsi
```
