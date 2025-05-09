# detected system type, set different icon paths
echo "Detected OS: ${OSTYPE}"
if [[ "$OSTYPE" == "linux-gnu"* ]]; then
    ICON_PATH="assets/logo.png"
elif [[ "$OSTYPE" == "darwin"* ]]; then
    ICON_PATH="assets/logo.icns"
elif [[ "$OSTYPE" == "msys" || "$OSTYPE" == "cygwin" ]]; then
    ICON_PATH="assets/logo.ico"
else
    echo "Unsupported OS type: $OSTYPE"
    exit 1
fi
echo "Using icon path: ${ICON_PATH}"
echo "Running PyInstaller to create executable..."
pyinstaller --name QMPlusScoring \
            --windowed \
            --onefile \
            --icon=${ICON_PATH} \
            app.py