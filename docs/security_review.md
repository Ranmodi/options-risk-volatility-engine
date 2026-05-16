# Security Review

This repository was sanitized for public portfolio usage.

## Removed or replaced

- Real e-mail recipients were replaced by example addresses.
- Internal paths were replaced by placeholders or sample paths.
- Company/provider-specific references were generalized where practical.
- No client datasets are included.
- No production Excel workbooks are included.
- No credentials, API keys or service accounts are included.

## Before making changes public

Check for:

```text
.env
*.json credentials
service-account files
API keys
private URLs
client data
account numbers
internal file paths
real e-mail recipients
production spreadsheets
```

## Recommended production pattern

Use environment variables or a secrets manager for:

- E-mail recipients;
- API keys;
- Paths to production files;
- Service credentials;
- Database connection strings.
