#!/bin/bash
# CS Reporter - Setup Script
# This script sets up the cs-reporter tool for first-time use

set -e  # Exit on any error

echo "========================================="
echo "  CS Reporter - Setup"
echo "========================================="
echo ""

# Check if Python 3 is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ Error: Python 3 is not installed."
    echo "Please install Python 3 from https://www.python.org/downloads/"
    exit 1
fi

echo "✓ Python 3 found: $(python3 --version)"
echo ""

# Create virtual environment
echo "Creating virtual environment..."
if [ -d ".venv" ]; then
    echo "  Virtual environment already exists, skipping..."
else
    python3 -m venv .venv
    echo "  ✓ Virtual environment created"
fi
echo ""

# Activate virtual environment
echo "Activating virtual environment..."
source .venv/bin/activate
echo "  ✓ Activated"
echo ""

# Check tkinter is available (needed for the file-picker dialogs).
# Homebrew Python does not bundle Tk; it must be installed separately.
echo "Checking tkinter..."
if ! python -c "import tkinter" 2>/dev/null; then
    PYVER=$(python -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')")
    echo "❌ Error: tkinter is not available for this Python (${PYVER})."
    echo "   cs-reporter uses tkinter for its file-selection dialogs."
    echo "   Install it, then re-run ./setup.sh:"
    echo "     brew install python-tk@${PYVER}"
    exit 1
fi
echo "  ✓ tkinter available"
echo ""

# Upgrade pip
echo "Upgrading pip..."
pip install --upgrade pip --quiet
echo "  ✓ pip upgraded"
echo ""

# Install cs-reporter and all dependencies
echo "Installing cs-reporter..."
pip install -e . --quiet
echo "  ✓ cs-reporter installed"
echo ""

# Create launcher script
echo "Creating launcher script..."
cat > run-reporter.sh << 'LAUNCHER_EOF'
#!/bin/bash
# CS Reporter - Launcher Script
# This script activates the virtual environment and runs the reporter

cd "$(dirname "$0")"
source .venv/bin/activate
reporter
LAUNCHER_EOF

chmod +x run-reporter.sh
echo "  ✓ Launcher script created"
echo ""

echo "========================================="
echo "  Setup Complete! ✓"
echo "========================================="
echo ""
echo "To run the reporter, use one of these commands:"
echo ""
echo "  Option 1 (Recommended):"
echo "    ./run-reporter.sh"
echo ""
echo "  Option 2:"
echo "    source .venv/bin/activate"
echo "    reporter"
echo ""
echo "The reporter will prompt you to select your Excel files."
echo "Generated reports will be saved in the 'output' folder."
echo ""
