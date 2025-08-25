#!/bin/bash

# Colors for output
GREEN='\033[0;32m'
BLUE='\033[0;34m'
RED='\033[0;31m'
NC='\033[0m' # No Color

# Azure DevOps Configuration
AZURE_ORG="monarch360"
AZURE_PROJECT="Monarch360"
AZURE_REPO="monarch-ai-search"
AZURE_REPO_URL="https://monarch360@dev.azure.com/monarch360/Monarch360/_git/monarch-ai-search"

# Local paths
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Check if git is installed
if ! command -v git &> /dev/null; then
    echo -e "${RED}Error: git is not installed. Please install it using Homebrew:${NC}"
    echo "brew install git"
    exit 1
fi

# Check if Azure CLI is installed
if ! command -v az &> /dev/null; then
    echo -e "${RED}Error: Azure CLI is not installed. Please install it using Homebrew:${NC}"
    echo "brew update && brew install azure-cli"
    exit 1
fi

# Function to check if we're in a git repository
is_git_repo() {
    git rev-parse --is-inside-work-tree &> /dev/null
}

# Initialize repository if needed
if ! is_git_repo; then
    echo -e "${BLUE}Initializing git repository...${NC}"
    git init
fi

# Configure git if needed
if [[ -z $(git config --get user.email) ]]; then
    echo -e "${BLUE}Please enter your git email:${NC}"
    read GIT_EMAIL
    git config user.email "$GIT_EMAIL"
fi

if [[ -z $(git config --get user.name) ]]; then
    echo -e "${BLUE}Please enter your git name:${NC}"
    read GIT_NAME
    git config user.name "$GIT_NAME"
fi

# Check if Azure remote exists
if ! git remote | grep -q "azure"; then
    echo -e "${BLUE}Adding Azure DevOps remote...${NC}"
    git remote add azure "$AZURE_REPO_URL"
else
    echo -e "${BLUE}Updating Azure DevOps remote URL...${NC}"
    git remote set-url azure "$AZURE_REPO_URL"
fi

# Stage all files
echo -e "${BLUE}Staging files...${NC}"
git add .

# Commit if there are changes
if git diff --staged --quiet; then
    echo -e "${BLUE}No changes to commit${NC}"
else
    echo -e "${BLUE}Enter commit message:${NC}"
    read COMMIT_MSG
    git commit -m "$COMMIT_MSG"
fi

# Push to Azure DevOps
echo -e "${BLUE}Pushing to Azure DevOps...${NC}"
if git push azure master; then
    echo -e "${GREEN}Successfully pushed to Azure DevOps!${NC}"
    echo -e "${GREEN}Repository URL: $AZURE_REPO_URL${NC}"
else
    echo -e "${RED}Failed to push to Azure DevOps. You might need to:${NC}"
    echo "1. Set up SSH keys or"
    echo "2. Use Azure DevOps Personal Access Token (PAT)"
    echo -e "${BLUE}Would you like to set up PAT authentication? (y/n)${NC}"
    read SETUP_PAT
    
    if [[ "$SETUP_PAT" == "y" ]]; then
        echo -e "${BLUE}Please enter your Azure DevOps PAT:${NC}"
        read -s AZURE_PAT
        
        # Update remote URL with PAT
        PAT_URL="https://$AZURE_PAT@dev.azure.com/$AZURE_ORG/$AZURE_PROJECT/_git/$AZURE_REPO"
        git remote set-url azure "$PAT_URL"
        
        echo -e "${BLUE}Trying to push again...${NC}"
        if git push azure master; then
            echo -e "${GREEN}Successfully pushed to Azure DevOps!${NC}"
            # Reset URL to HTTPS for security (don't store PAT in git config)
            git remote set-url azure "$AZURE_REPO_URL"
        else
            echo -e "${RED}Push failed. Please check your PAT and try again.${NC}"
        fi
    fi
fi

echo -e "${BLUE}Done!${NC}"
