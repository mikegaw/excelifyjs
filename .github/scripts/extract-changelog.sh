#!/bin/bash
# Extract release notes for a specific version from CHANGELOG.md
# Usage: ./extract-changelog.sh <version>
# Example: ./extract-changelog.sh 0.1.3

set -e

VERSION="$1"

if [ -z "$VERSION" ]; then
  echo "Error: Version number required" >&2
  echo "Usage: $0 <version>" >&2
  exit 1
fi

if [ ! -f "CHANGELOG.md" ]; then
  echo "Error: CHANGELOG.md not found" >&2
  exit 1
fi

# Extract content between ## VERSION and next ## or EOF
# Using awk for reliable multi-line extraction
NOTES=$(awk -v ver="$VERSION" '
  /^## / {
    if (found) exit;
    if ($2 == ver) { found=1; next }
  }
  found { print }
' CHANGELOG.md)

if [ -z "$NOTES" ]; then
  echo "Error: Version $VERSION not found in CHANGELOG.md" >&2
  exit 1
fi

echo "$NOTES"
