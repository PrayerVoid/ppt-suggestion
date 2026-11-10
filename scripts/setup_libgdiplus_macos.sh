#!/usr/bin/env bash
# 在没有 Homebrew 的 macOS（含 Apple 芯片）上编译 libgdiplus，让 Aspose.Slides 能渲染预览图。
# 不需要管理员权限；所有东西都装在 $GDIP_HOME（默认 ~/.comate-engine/gdip）。
#
# 用法:  bash setup_libgdiplus_macos.sh
# 变量:  GDIP_HOME   安装位置，默认 ~/.comate-engine/gdip
#        BUILD_ENV   依赖环境位置，默认 ~/.comate-engine/gdip-build
#        FORCE=1     已安装也重装
# 前提:  Xcode Command Line Tools（xcode-select --install）。
# 耗时:  约 3-5 分钟（下载依赖约 200MB + 编译）。
#
# 思路: conda-forge 没有 libgdiplus，但有它的全部依赖（cairo/pango/glib/libpng 等）。
# 用 micromamba 装依赖，再用 Mono 官方源码包只编译 src/ 里的库本身。
# 注意要用 download.mono-project.com 的发布包，GitHub 源码包缺子模块，make 会失败；
# tests/ 目录需要 cmake 拉 googletest，也不用编，所以只 make -C src。
set -euo pipefail

GDIP_HOME="${GDIP_HOME:-$HOME/.comate-engine/gdip}"
BUILD_ENV="${BUILD_ENV:-$HOME/.comate-engine/gdip-build}"
MM_DIR="$HOME/.comate-engine/mm"
CONDA_MIRROR="https://mirrors.tuna.tsinghua.edu.cn/anaconda/cloud/conda-forge"
LG_URL="https://download.mono-project.com/sources/libgdiplus/libgdiplus-6.1.tar.gz"

if [ "$(uname -s)" != "Darwin" ]; then
  echo "只用于 macOS。Linux 请 apt-get install libgdiplus。" >&2; exit 1
fi
if [ -f "$GDIP_HOME/lib/libgdiplus.dylib" ] && [ "${FORCE:-0}" != "1" ]; then
  echo "已安装: $GDIP_HOME/lib/libgdiplus.dylib（FORCE=1 可重装）"; exit 0
fi
xcode-select -p >/dev/null 2>&1 || { echo "缺 Xcode Command Line Tools: xcode-select --install" >&2; exit 1; }

case "$(uname -m)" in arm64) PLAT=osx-arm64;; *) PLAT=osx-64;; esac

# 1) micromamba（单文件，从清华 conda-forge 镜像取）
if [ ! -x "$MM_DIR/bin/micromamba" ]; then
  mkdir -p "$MM_DIR"
  curl -fsSL -o "$MM_DIR/mm.tar.bz2" "$CONDA_MIRROR/$PLAT/micromamba-2.0.5-0.tar.bz2"
  tar -xjf "$MM_DIR/mm.tar.bz2" -C "$MM_DIR" bin/micromamba
fi

# 2) 依赖环境
"$MM_DIR/bin/micromamba" create -y -p "$BUILD_ENV" -c "$CONDA_MIRROR" --override-channels \
  cairo glib libpng giflib libjpeg-turbo libtiff freetype fontconfig pango libexif \
  pkg-config zlib pixman expat bzip2 >/dev/null

# 3) 源码编译（只编 src/）
WORK="$(mktemp -d)"; trap 'rm -rf "$WORK"' EXIT
curl -fsSL -o "$WORK/lg.tar.gz" "$LG_URL"
tar -xzf "$WORK/lg.tar.gz" -C "$WORK"
cd "$WORK"/libgdiplus-*
export PATH="$BUILD_ENV/bin:$PATH"
export PKG_CONFIG_PATH="$BUILD_ENV/lib/pkgconfig:$BUILD_ENV/share/pkgconfig"
export CPPFLAGS="-I$BUILD_ENV/include"
export LDFLAGS="-L$BUILD_ENV/lib -Wl,-rpath,$BUILD_ENV/lib"
./configure --prefix="$GDIP_HOME" --without-x11 >/dev/null
make -C src -j"$(sysctl -n hw.ncpu)" >/dev/null
make -C src install >/dev/null

echo "完成: $GDIP_HOME/lib/libgdiplus.dylib"
echo "render_preview.py 会自动加载它；手动使用时: export DYLD_LIBRARY_PATH=$GDIP_HOME/lib"
