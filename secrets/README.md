# Secrets Directory

This directory contains sensitive files that are excluded from Git.

## Security Rules
- Files with `.pem`, `.key`, `.crt` extensions are automatically ignored
- Never commit sensitive data to version control
- This README.md is tracked to preserve the directory structure

## Contents
- `deploy-key.pem` - SSH private key for server deployment (IGNORED by Git)
