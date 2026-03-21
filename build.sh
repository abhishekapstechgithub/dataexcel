#!/usr/bin/env bash
# ============================================================
#  QtSpreadsheet — Linux/macOS Build Script
#  Requirements: Qt 6.5+, CMake 3.25+, GCC 12+ or Clang 16+
# ============================================================
set -euo pipefail

BUILD_TYPE="${1:-Release}"
QT_DIR="${2:-/opt/Qt/6.5.3/gcc_64}"
BUILD_DIR="build-linux"
JOBS=$(nproc 2>/dev/null || sysctl -n hw.ncpu 2>/dev/null || echo 4)

echo ""
echo "  =================================================="
echo "   QtSpreadsheet Linux/macOS Builder"
echo "   Build type : $BUILD_TYPE"
echo "   Qt path    : $QT_DIR"
echo "   Jobs       : $JOBS"
echo "  =================================================="
echo ""

# Detect Qt automatically if not at given path
if [[ ! -f "$QT_DIR/lib/cmake/Qt6/Qt6Config.cmake" ]]; then
    for candidate in \
        /usr/lib/qt6 /usr/lib/x86_64-linux-gnu/cmake/Qt6 \
        ~/Qt/6.*/gcc_64 ~/Qt/6.*/macos; do
        expanded=( $candidate )
        for e in "${expanded[@]}"; do
            if [[ -f "$e/lib/cmake/Qt6/Qt6Config.cmake" ]]; then
                QT_DIR="$e"; break 2
            fi
        done
    done
fi

echo "[1/3] Configuring..."
# Remove stale build dir to prevent cached CMake errors blocking a fresh build
rm -rf "$BUILD_DIR"
cmake -B "$BUILD_DIR" \
    -DCMAKE_BUILD_TYPE="$BUILD_TYPE" \
    -DCMAKE_PREFIX_PATH="$QT_DIR" \
    -DQt6_DIR="$QT_DIR/lib/cmake/Qt6"

echo "[2/3] Building ($JOBS parallel jobs)..."
cmake --build "$BUILD_DIR" --config "$BUILD_TYPE" --parallel "$JOBS"

echo "[3/3] Done."
echo ""
echo "  =================================================="
echo "   BUILD SUCCESSFUL"
echo "   Executable: $BUILD_DIR/bin/QtSpreadsheet"
echo "  =================================================="
