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

# Upgrade pip
echo "Upgrading pip..."
pip install --upgrade pip --quiet
echo "  ✓ pip upgraded"
echo ""

# Install dependencies
echo "Installing dependencies..."
pip install -r requirements.txt --quiet
echo "  ✓ Dependencies installed"
echo ""

# Install package in editable mode
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
