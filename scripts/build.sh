#!/bin/bash
# Build the Thunderbird TbSync MCP extension

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"
EXTENSION_DIR="$PROJECT_DIR/extension"
DIST_DIR="$PROJECT_DIR/dist"

echo "Building Thunderbird TbSync MCP extension..."

mkdir -p "$DIST_DIR"

cd "$EXTENSION_DIR"
zip -r "$DIST_DIR/thunderbird-tbsync-mcp.xpi" . -x "*.DS_Store" -x "*.git*"

echo "Built: $DIST_DIR/thunderbird-tbsync-mcp.xpi"
