#!/bin/bash

# Python Boilerplate Setup Script
# This script installs uv and sets up a Python development environment

set -e  # Exit on any error

# Parse command line arguments
NON_INTERACTIVE=false
SHOW_HELP=false

for arg in "$@"; do
    case $arg in
        -y|--yes)
            NON_INTERACTIVE=true
            shift
            ;;
        -h|--help)
            SHOW_HELP=true
            shift
            ;;
        *)
            # Unknown option
            ;;
    esac
done

# Show help if requested
if [ "$SHOW_HELP" = true ]; then
    echo "Python Boilerplate Setup Script"
    echo "Usage: ./setup.sh [-y] [-h]"
    echo "  -y, --yes   : Non-interactive mode (use defaults)"
    echo "  -h, --help  : Show this help message"
    exit 0
fi

# Function to convert string to valid Python module name
convert_to_python_module_name() {
    local name="$1"
    # Convert to lowercase, replace spaces and hyphens with underscores, remove invalid chars
    name=$(echo "$name" | tr '[:upper:]' '[:lower:]' | sed 's/[[:space:]-]/_/g' | sed 's/[^a-z0-9_]//g')
    # Ensure it doesn't start with a number
    if [[ $name =~ ^[0-9] ]]; then
        name="module_$name"
    fi
    # Capitalize first letter for class-style naming
    echo "${name^}"
}

# Function to get current module name from pyproject.toml
get_current_module_name() {
    if [ -f "pyproject.toml" ]; then
        grep -o 'name = "[^"]*"' pyproject.toml 2>/dev/null | sed 's/name = "\([^"]*\)"/\1/' || echo ""
    else
        echo ""
    fi
}

echo "🚀 Setting up Python development environment with uv..."

# Check if uv is already installed
if command -v uv &> /dev/null; then
    echo "✅ uv is already installed"
else
    echo "📦 Installing uv..."
    # Install uv using the official installer
    curl -LsSf https://astral.sh/uv/install.sh | sh
    
    # Add uv to PATH for current session
    export PATH="$HOME/.cargo/bin:$PATH"
    
    # Verify installation
    if command -v uv &> /dev/null; then
        echo "✅ uv installed successfully"
    else
        echo "❌ Failed to install uv"
        exit 1
    fi
fi

# Infer module name from directory name
directory_name=$(basename "$(pwd)")
inferred_module_name=$(convert_to_python_module_name "$directory_name")
current_module_name=$(get_current_module_name)

# Determine the module name to use
module_name=""
if [ -n "$current_module_name" ] && [ "$current_module_name" != "MyModule" ]; then
    # Already configured with a custom name, keep it
    module_name="$current_module_name"
    echo "📝 Using existing module name: $module_name"
elif [ "$NON_INTERACTIVE" = true ]; then
    # Non-interactive mode, use inferred name
    module_name="$inferred_module_name"
    echo "📝 Using inferred module name: $module_name (non-interactive mode)"
else
    # Interactive mode - ask user
    echo "📝 Module name configuration"
    echo "   Current directory: $directory_name"
    echo "   Suggested module name: $inferred_module_name"
    if [ -n "$current_module_name" ] && [ "$current_module_name" != "MyModule" ]; then
        echo "   Current pyproject.toml name: $current_module_name"
    fi
    echo ""
    read -p "Enter module name (press Enter for '$inferred_module_name'): " user_input
    if [ -z "$user_input" ]; then
        module_name="$inferred_module_name"
    else
        module_name=$(convert_to_python_module_name "$user_input")
    fi
    echo "✅ Using module name: $module_name"
fi

# Check if we're in an existing project or need to refresh
if [ -f "pyproject.toml" ] && [ -d "src" ]; then
    echo "🔄 Existing project detected. Refreshing..."
    
    # Remove existing virtual environment if it exists
    if [ -d ".venv" ]; then
        echo "🗑️  Removing existing virtual environment..."
        rm -rf .venv
    fi
    
    # Clean up any cached files
    find . -name "__pycache__" -type d -exec rm -rf {} + 2>/dev/null || true
    find . -name "*.pyc" -delete 2>/dev/null || true
else
    echo "🆕 Setting up new project structure..."
fi

# Update module name in pyproject.toml if it's different
if [ -f "pyproject.toml" ] && [ "$module_name" != "$current_module_name" ]; then
    echo "🔧 Updating module name in pyproject.toml..."
    sed -i.bak "s/name = \"[^\"]*\"/name = \"$module_name\"/" pyproject.toml
    rm -f pyproject.toml.bak 2>/dev/null || true
fi

# Sync src directory structure with pyproject.toml name
if [ -d "src" ]; then
    # Get the final module name (could be from pyproject.toml if manually edited)
    final_module_name=$(get_current_module_name)
    
    # Find existing module directories in src/ (excluding __pycache__)
    existing_dir=$(find src -maxdepth 1 -type d ! -name "src" ! -name "__pycache__" | head -1 | xargs basename 2>/dev/null || echo "")
    
    if [ -n "$existing_dir" ] && [ "$existing_dir" != "$final_module_name" ]; then
        echo "🔧 Syncing module directory: $existing_dir → $final_module_name..."
        
        if [ -d "src/$final_module_name" ]; then
            # Target already exists, remove the old one
            rm -rf "src/$existing_dir"
            echo "   Removed duplicate directory: src/$existing_dir"
        else
            # Rename to match pyproject.toml
            mv "src/$existing_dir" "src/$final_module_name"
            echo "   Renamed: src/$existing_dir → src/$final_module_name"
        fi
    fi
    
    # Update the module_name variable for later use
    module_name="$final_module_name"
fi

# Create virtual environment and install dependencies
echo "🔧 Creating virtual environment and installing dependencies..."
uv sync

# Activate virtual environment and show Python version
echo "🐍 Virtual environment ready!"
echo "To activate the environment, run: source .venv/bin/activate (Linux/Mac) or .venv\Scripts\activate (Windows)"

# Check for requirements.txt and migrate if found
if [ -f "requirements.txt" ]; then
    echo "📋 Found requirements.txt - migrating to uv..."
    if uv add --requirements requirements.txt; then
        echo "✅ Successfully migrated requirements.txt to pyproject.toml"
        echo "💡 You can now delete requirements.txt (optional)"
    else
        echo "⚠️  Could not migrate requirements.txt - you may need to add dependencies manually:"
        echo "    uv add --requirements requirements.txt"
    fi
fi

# Show installed packages
echo "📦 Installed packages:"
uv pip list

echo "✨ Setup complete! Your Python development environment is ready."
echo "🧪 Run tests with: uv run pytest"
echo "🚀 Start developing in the src/$module_name directory"

# Handle git repository initialization
echo ""
if [ -d ".git" ]; then
    echo "🔍 Git repository detected."
    echo "⚠️  To use this as a new project, you should remove the boilerplate git history."
    echo ""
    echo "Options:"
    echo "  1. Remove git history and start fresh (git init)"
    echo "  2. Remove git history, no git init (for git clone later)"
    echo "  3. Keep current git history"
    echo "  4. Skip for now (you can run this script again later)"
    echo ""
    read -p "Choose option (1/2/3/4): " git_choice
    
    case $git_choice in
        1)
            echo "🗑️  Removing boilerplate git history..."
            rm -rf .git
            echo "🆕 Initializing new git repository..."
            git init
            git add .
            echo "📝 Ready for your initial commit!"
            echo "    Run: git commit -m 'Initial commit'"
            echo "    Then add your remote: git remote add origin <your-repo-url>"
            ;;
        2)
            echo "🗑️  Removing boilerplate git history..."
            rm -rf .git
            echo "📝 Ready for git clone! You can now:"
            echo "    git clone <your-repo-url> temp && mv temp/.git . && rm -rf temp"
            echo "    git reset --hard HEAD  # to sync with remote"
            echo "    Or initialize later with: git init"
            ;;
        3)
            echo "✅ Keeping current git history."
            ;;
        4)
            echo "⏭️  Skipped git setup. You can run this script again later."
            ;;
        *)
            echo "❓ Invalid choice. Skipping git setup."
            ;;
    esac
else
    echo "📝 No git repository found. You can:"
    echo "    Initialize: git init && git add . && git commit -m 'Initial commit'"
    echo "    Or clone: git clone <your-repo-url> temp && mv temp/.git . && rm -rf temp"
    echo "    Then: git reset --hard HEAD"
fi

# Add project setup files and documentation to .gitignore
echo ""
echo "🚫 Adding project setup files to .gitignore..."
gitignore_entries="
# Project setup files (added by ProjectSetup scripts)
ProjectSetup.ps1
ProjectSetup.sh
HOW2USEIT.md
uv.lock"

if [ -f ".gitignore" ]; then
    # Check if entries are already there
    if ! grep -q "ProjectSetup.ps1" .gitignore; then
        echo "$gitignore_entries" >> .gitignore
        echo "   ✅ Added setup files to .gitignore"
    else
        echo "   ✅ Setup files already in .gitignore"
    fi
else
    echo "   ⚠️  No .gitignore found - setup files not excluded from git"
fi