#!/bin/bash

ve() {
    # Check if Python is installed:
    if command -v python3 &> /dev/null; then
        echo "Python3  is installed."
        python3 --version
    else
        echo "Python3 is not installed. Please install Python first."
        exit
    fi


    # Check is venv module is available:
    if [ ! command -v &> /dev/null ]; then
        echo "Python venv module is not available. Ensure Python >= 3.3 is installed."
        exit
    fi


    # Create virtual environment if it doesn't exist:
    if [ ! -d "venv" ]; then
        echo "Creating virtual environment..."
        python3 -m venv ./venv
    else
        echo "Virtual environment venv already exists"
    fi


    # Activate virtual environment:
    echo "Activating virtual environment..."
    DIR=$(realpath "$(dirname "${BASH_SOURCE[0]}")")
    echo $DIR/venv/bin/activate
    source $DIR/venv/bin/activate

    # Install requirements:
    if [ -f "requirements.txt" ]; then
        echo "Installing dependencies from requirements.txt..."
        pip install -r requirements.txt
    else
        echo "requirements.txt not found."
    fi
    
    echo "Installed dependencies:"
    pip freeze
    
    echo "Virtual environment setup complete."
}
