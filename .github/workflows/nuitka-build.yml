name: Build with Nuitka (Standalone Folder Mode - Final Diagnostic)

on:
  push:
    branches: [ "main", "master" ]
  workflow_dispatch:

jobs:
  build-windows:
    strategy:
      matrix:
        include:
          - python-version: '3.8.10'
            architecture: 'x86'
            artifact-name: 'boyin-Windows-x86-for-Win7-NUITKA'
            vlc-folder: 'vlc_lib_x86'

          - python-version: '3.8.10'
            architecture: 'x64'
            artifact-name: 'boyin-Windows-x64-for-Win7-NUITKA'
            vlc-folder: 'vlc_lib_x64'

          - python-version: '3.11'
            architecture: 'x64'
            artifact-name: 'boyin-Windows-x64-for-Win8.1+-NUITKA'
            vlc-folder: 'vlc_lib_x64'

    runs-on: windows-latest

    steps:
      - name: Check out repository code
        uses: actions/checkout@v4

      - name: Set up Python ${{ matrix.python-version }} (${{ matrix.architecture }})
        uses: actions/setup-python@v5
        with:
          python-version: '${{ matrix.python-version }}'
          architecture: '${{ matrix.architecture }}'
          
      - name: Cache pip dependencies
        uses: actions/cache@v4
        with:
          path: ~\AppData\Local\pip\Cache
          # 强制刷新缓存到 v3
          key: v3-${{ runner.os }}-nuitka-${{ matrix.python-version }}-${{ matrix.architecture }}-pip-${{ hashFiles('**/requirements_nuitka.txt') }}
          restore-keys: |
            v3-${{ runner.os }}-nuitka-${{ matrix.python-version }}-${{ matrix.architecture }}-pip-

      - name: Install dependencies with Strict Error Checking
        run: |
          # --- ↓↓↓ 核心修改：强制 PowerShell 在遇到任何错误时立即停止 ↓↓↓ ---
          $ErrorActionPreference = "Stop"

          python -m pip install --upgrade pip
          Write-Host "=== Installing from requirements_nuitka.txt ==="
          pip install -r requirements_nuitka.txt
          
          Write-Host "=== Listing all installed packages for verification ==="
          pip list
          
          Write-Host "=== Installing Nuitka ==="
          pip install nuitka
        shell: pwsh

      - name: Build and Archive application with Nuitka
        run: |
          $ErrorActionPreference = "Stop"
          Write-Host "=== 构建 Nuitka 编译版本 (Standalone 文件夹模式) ==="
          $outputDir = "boyin.dist"
          $exeName = "创翔多功能定时播音旗舰版.exe"
          $zipName = "${{ matrix.artifact-name }}.zip"

          $nuitkaArgs = @(
              "--standalone",
              "--windows-console-mode=disable",
              "--lto=yes",
              "--assume-yes-for-downloads",
              "--output-dir=$outputDir",
              "--output-filename=$exeName",
              "--windows-icon-from-ico=icon.ico",
              "--product-name=创翔多功能定时播音旗舰版",
              "--file-version=1.0.0",
              "--company-name=创翔",
              "--enable-plugin=tk-inter",
              "--enable-plugin=pyside6",
              "--enable-plugin=anti-bloat",
              "--include-package-data=pygame",
              "--include-package-data=ttkbootstrap",
              "--include-data-file=icon.ico=icon.ico",
              "--include-data-dir=${{ matrix.vlc-folder }}=${{ matrix.vlc-folder }}",
              "boyin.py"
          )
          
          python -m nuitka $nuitkaArgs

          Write-Host "=== 正在压缩构建产物... ==="
          Compress-Archive -Path "$outputDir\*" -DestinationPath $zipName -Force
        shell: pwsh

      - name: Upload Nuitka artifact
        uses: actions/upload-artifact@v4
        with:
          name: ${{ matrix.artifact-name }}
          path: ${{ matrix.artifact-name }}.zip
