#!/bin/bash
# Installation script for excel.nvim

set -e

echo "======================================"
echo "excel.nvim Installation Script"
echo "======================================"
echo ""

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Check Python
echo "Checking Python installation..."
if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
    echo -e "${GREEN}✓${NC} Python3 found: $(python3 --version)"
elif command -v python &> /dev/null; then
    PYTHON_CMD="python"
    echo -e "${GREEN}✓${NC} Python found: $(python --version)"
else
    echo -e "${RED}✗${NC} Python not found. Please install Python 3.7+"
    exit 1
fi

# Check Neovim
echo "Checking Neovim installation..."
if command -v nvim &> /dev/null; then
    echo -e "${GREEN}✓${NC} Neovim found: $(nvim --version | head -n1)"
else
    echo -e "${RED}✗${NC} Neovim not found. Please install Neovim 0.8.0+"
    exit 1
fi

# Install Python dependencies
echo ""
echo "Installing Python dependencies..."
echo "Running: $PYTHON_CMD -m pip install openpyxl pandas"

if $PYTHON_CMD -m pip install openpyxl pandas --user; then
    echo -e "${GREEN}✓${NC} Python dependencies installed successfully"
else
    echo -e "${YELLOW}⚠${NC} Some packages may have failed to install"
    echo "  Try manually: pip install openpyxl pandas"
fi

# Check LibreOffice (optional)
echo ""
echo "Checking for LibreOffice (optional, for formula recalculation)..."
if command -v libreoffice &> /dev/null || command -v soffice &> /dev/null; then
    echo -e "${GREEN}✓${NC} LibreOffice found"
else
    echo -e "${YELLOW}⚠${NC} LibreOffice not found"
    echo "  Formula recalculation will not be available"
    echo "  Install with:"
    echo "    - Ubuntu/Debian: sudo apt-get install libreoffice"
    echo "    - macOS: brew install libreoffice"
fi

# Create data directory
echo ""
echo "Setting up data directory..."
DATA_DIR="$HOME/.local/share/nvim/excel_nvim"
mkdir -p "$DATA_DIR"
echo -e "${GREEN}✓${NC} Data directory created: $DATA_DIR"

# Copy Python handler
echo ""
echo "Installing Python handler..."
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
if [ -f "$SCRIPT_DIR/python/excel_handler.py" ]; then
    cp "$SCRIPT_DIR/python/excel_handler.py" "$DATA_DIR/"
    chmod +x "$DATA_DIR/excel_handler.py"
    echo -e "${GREEN}✓${NC} Python handler installed"
else
    echo -e "${YELLOW}⚠${NC} Python handler not found in expected location"
fi

# Installation instructions
echo ""
echo "======================================"
echo "Installation Complete!"
echo "======================================"
echo ""
echo "Next steps:"
echo ""
echo "1. Add to your Neovim configuration:"
echo ""
echo "   Using lazy.nvim:"
echo "   {"
echo "     'excel.nvim',"
echo "     dir = '$SCRIPT_DIR',"
echo "     config = function()"
echo "       require('excel').setup()"
echo "     end,"
echo "   }"
echo ""
echo "   Or manually add to runtimepath:"
echo "   set runtimepath+=$SCRIPT_DIR"
echo ""
echo "2. Restart Neovim"
echo ""
echo "3. Open an Excel file:"
echo "   nvim myfile.xlsx"
echo ""
echo "4. See README.md for usage instructions and keybindings"
echo ""
echo "======================================"

# Test Python script
echo ""
echo "Testing Python script..."
if [ -f "$DATA_DIR/excel_handler.py" ]; then
    if $PYTHON_CMD "$DATA_DIR/excel_handler.py" 2>&1 | grep -q "error"; then
        echo -e "${RED}✗${NC} Python script test failed"
        $PYTHON_CMD "$DATA_DIR/excel_handler.py"
    else
        echo -e "${GREEN}✓${NC} Python script is working"
    fi
fi

echo ""
echo "Installation script finished!"
