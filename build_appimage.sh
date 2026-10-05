#!/usr/bin/env bash
set -e

# =============================================================================
# BUILD_APPIMAGE.SH - Gerador de Pacote Linux AppImage (.image / .AppImage)
# Projeto: Convert_IMCA (v1.6.0)
# =============================================================================

PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APPDIR="${PROJECT_DIR}/build/AppDir"
OUTPUT_NAME="Convert_IMCA-x86_64.AppImage"

PYTHON_BIN="python3"
if command -v python3.10 >/dev/null 2>&1; then
    PYTHON_BIN="python3.10"
fi

echo "============================================================"
echo " 🚀 INICIANDO GERAÇÃO DO PACOTE LINUX APPIMAGE"
echo "============================================================"

# 1. Compilar executável com PyInstaller caso ainda não exista
echo "📦 [1/5] Compilando executável da interface com PyInstaller (${PYTHON_BIN})..."
${PYTHON_BIN} -m PyInstaller --noconfirm --onedir --windowed \
    --name "convert_imca_gui" \
    --add-data "assets:assets" \
    --distpath "${PROJECT_DIR}/dist" \
    --workpath "${PROJECT_DIR}/build" \
    "${PROJECT_DIR}/convert_imca_gui.py"

# 2. Criar estrutura do AppDir
echo "📂 [2/5] Estruturando AppDir..."
rm -rf "${APPDIR}"
mkdir -p "${APPDIR}/usr/bin"
mkdir -p "${APPDIR}/usr/share/icons/hicolor/256x256/apps"

# Copia os arquivos do pacote executável
cp -r "${PROJECT_DIR}/dist/convert_imca_gui/"* "${APPDIR}/usr/bin/"

# Copia ícones
if [ -f "${PROJECT_DIR}/assets/icon.png" ]; then
    cp "${PROJECT_DIR}/assets/icon.png" "${APPDIR}/convert_imca.png"
    cp "${PROJECT_DIR}/assets/icon.png" "${APPDIR}/usr/share/icons/hicolor/256x256/apps/convert_imca.png"
    cp "${PROJECT_DIR}/assets/icon.png" "${APPDIR}/.DirIcon"
fi

# Copia o binário Rust da CLI se estiver compilado
if [ -f "${PROJECT_DIR}/target/release/convert_imca" ]; then
    cp "${PROJECT_DIR}/target/release/convert_imca" "${APPDIR}/usr/bin/convert_imca"
fi

# 3. Criar arquivo .desktop
echo "📝 [3/5] Gerando convert_imca.desktop..."
cat << 'EOF' > "${APPDIR}/convert_imca.desktop"
[Desktop Entry]
Type=Application
Name=Convert_IMCA
GenericName=Conversor de Planilhas
Comment=Conversor de Planilhas Excel para Formato Leve .IMCA (cabeceira-pwa1)
Exec=convert_imca_gui
Icon=convert_imca
Categories=Office;Utility;
Terminal=false
StartupNotify=true
EOF

# 4. Criar script AppRun
echo "⚙️ [4/5] Gerando script AppRun..."
cat << 'EOF' > "${APPDIR}/AppRun"
#!/bin/sh
set -e
HERE="$(dirname "$(readlink -f "${0}")")"
export PATH="${HERE}/usr/bin:${PATH}"
export LD_LIBRARY_PATH="${HERE}/usr/bin:${LD_LIBRARY_PATH}"
export TCL_LIBRARY="${HERE}/usr/bin/_internal/tcl"
export TK_LIBRARY="${HERE}/usr/bin/_internal/tk"
exec "${HERE}/usr/bin/convert_imca_gui" "$@"
EOF
chmod +x "${APPDIR}/AppRun"

# 5. Executar appimagetool para empacotar
echo "🛠️ [5/5] Empacotando com appimagetool..."
APPIMAGETOOL="/tmp/appimage_tools/squashfs-root/AppRun"

if [ ! -f "${APPIMAGETOOL}" ]; then
    if which appimagetool >/dev/null 2>&1; then
        APPIMAGETOOL="appimagetool"
    else
        echo "Baixando appimagetool..."
        mkdir -p /tmp/appimage_tools && cd /tmp/appimage_tools
        curl -sL -o appimagetool "https://github.com/AppImage/AppImageKit/releases/download/continuous/appimagetool-x86_64.AppImage"
        chmod +x appimagetool
        ./appimagetool --appimage-extract >/dev/null 2>&1 || true
        APPIMAGETOOL="/tmp/appimage_tools/squashfs-root/AppRun"
        cd "${PROJECT_DIR}"
    fi
fi

mkdir -p "${PROJECT_DIR}/dist"
ARCH=x86_64 "${APPIMAGETOOL}" "${APPDIR}" "${PROJECT_DIR}/dist/${OUTPUT_NAME}"

# Criar atalho com extensão .image para conveniência
cp "${PROJECT_DIR}/dist/${OUTPUT_NAME}" "${PROJECT_DIR}/dist/Convert_IMCA.image"

echo "============================================================"
echo "✨ SUCESSO! Pacote Linux AppImage gerado:"
echo "   📍 dist/${OUTPUT_NAME}"
echo "   📍 dist/Convert_IMCA.image"
ls -lh "${PROJECT_DIR}/dist/${OUTPUT_NAME}"
echo "============================================================"
