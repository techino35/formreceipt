# FormReceipt

A Google Apps Script add-on that automatically generates PDFs from Google Form responses and emails them to respondents.

## Features
- Auto PDF generation on form submit
- Google Docs template with `{{placeholder}}` substitution
- Email receipt with PDF attachment
- Drive auto-save: `/FormReceipt/[form-name]/[date]_[seq].pdf`
- Free: 10/month | Pro: Unlimited + logo insertion

## Setup
```bash
cp .clasp.json.example .clasp.json
npx @google/clasp push
```

## Placeholders
- `{{Receipt Number}}` - Auto-generated (2024-01-0001 format)
- `{{Submission Date/Time}}` - Form submission timestamp
- `{{Submission Date}}` - Form submission date

## License
- Free: 10/month
- Pro: Unlimited + logo (key format: FR-XXXX-XXXX-XXXX)
