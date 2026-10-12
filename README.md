# IFEXA Academy — Certificate Automation

## Overview

This repository contains the Google Apps Script source code for IFEXA Academy's certificate automation workflow.

The system supports certificate generation, PDF storage, QR-code verification, student tracking, and certificate delivery by email.

## Current Architecture

- **Google Sheets:** Student and certificate tracker.
- **Google Apps Script:** Automation logic.
- **Google Slides:** Certificate template.
- **Google Drive:** Generated PDF certificates.
- **Email:** Certificate delivery to students.
- **Verification web app:** Used by recipients to verify certificates.

## Important Tracker Fields

The `Certificate Tracker` sheet uses the following columns:

| Column | Purpose |
|---|---|
| A | Student ID |
| B | Full Name |
| C | Email |
| D | Attendance |
| E | Points |
| F | Student Payment Status |
| G | IFEXA Payment Verification |
| H | Certificate Eligibility |
| I | Certificate Status |
| J | Certificate No. |
| K | Issued Date |
| L | Certificate Link |
| M | QR Status |
| N | Certificate Last Sent |
| O | Selected Batch Status |

The registration form responses are stored in the `Form Responses 1` tab.

## Development Rules

1. Do not modify the production automation without testing changes first.
2. Never upload student personal data, credentials, access tokens, or private keys.
3. Keep production configuration separate from public source code.
4. Avoid generating duplicate certificates for the same student.
5. Verify student eligibility and payment status according to IFEXA's approved process.
6. Test certificate generation and email delivery using test records before production use.
7. Document significant changes in GitHub commits and pull requests.

## Deployment Notes

This repository is a source-code backup. Changes committed to GitHub do not automatically update the deployed Google Apps Script project.

Before deploying a change, review it, test it, and confirm that the production spreadsheet, certificate template, Drive folders, and verification service are configured correctly.

## Maintainer

IFEXA Academy — IFEXA Group
