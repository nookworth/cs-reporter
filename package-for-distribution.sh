#!/bin/bash
# CS Reporter - Distribution Package Creator
# Creates a clean zip file ready to send to users

set -e

echo "========================================="
echo "  CS Reporter - Package for Distribution"
echo "========================================="
echo ""

# Get version/date for filename
VERSION=$(date +%Y%m%d)
PACKAGE_NAME="cs-reporter-${VERSION}.zip"

# Create temporary directory
TEMP_DIR=$(mktemp -d)
DIST_DIR="${TEMP_DIR}/cs-reporter"

echo "Creating distribution package..."
echo ""

# Create directory structure
mkdir -p "${DIST_DIR}"

# Copy essential files
echo "Copying files..."
cp -r src "${DIST_DIR}/"
cp -r config "${DIST_DIR}/"
cp -r templates "${DIST_DIR}/"
cp setup.py "${DIST_DIR}/"
cp setup.sh "${DIST_DIR}/"
cp setup.bat "${DIST_DIR}/"
cp run-reporter.bat "${DIST_DIR}/"
cp requirements.txt "${DIST_DIR}/"
cp QUICKSTART.txt "${DIST_DIR}/"
cp README.md "${DIST_DIR}/"
cp USER_GUIDE.md "${DIST_DIR}/"

# Create empty output directory with .gitkeep
mkdir -p "${DIST_DIR}/output"
touch "${DIST_DIR}/output/.gitkeep"

# Clean up Python cache files
echo "Cleaning up..."
find "${DIST_DIR}" -type d -name "__pycache__" -exec rm -rf {} + 2>/dev/null || true
find "${DIST_DIR}" -type f -name "*.pyc" -delete 2>/dev/null || true
find "${DIST_DIR}" -type f -name ".DS_Store" -delete 2>/dev/null || true

# Create zip file
echo "Creating zip archive..."
cd "${TEMP_DIR}"
zip -r -q "${PACKAGE_NAME}" cs-reporter

# Move to current directory
mv "${PACKAGE_NAME}" "${OLDPWD}/"

# Clean up temp directory
rm -rf "${TEMP_DIR}"

echo ""
echo "========================================="
echo "  Package Created! ✓"
echo "========================================="
echo ""
echo "Distribution package: ${PACKAGE_NAME}"
echo ""
echo "Next steps:"
echo "1. Send '${PACKAGE_NAME}' to your user"
echo "2. User unzips the file"
echo "3. User runs: ./setup.sh"
echo "4. User runs: ./run-reporter.sh"
echo ""
