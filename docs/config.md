# SMTP Configuration Guide

## Gmail Configuration

1. **Enable 2-Factor Authentication** (Recommended)
   - Go to your Google Account settings
   - Enable 2-Step Verification

2. **Generate an App Password**
   - Go to https://myaccount.google.com/apppasswords
   - Generate a password for "Mail"
   - Use this password in the application instead of your regular password

3. **SMTP Settings**
   - SMTP Server: `smtp.gmail.com`
   - Port: `587`
   - TLS: Enabled

## Outlook/Hotmail Configuration

- SMTP Server: `smtp.office365.com`
- Port: `587`
- TLS: Enabled

## Yahoo Mail Configuration

- SMTP Server: `smtp.mail.yahoo.com`
- Port: `587`
- TLS: Enabled

## Other Email Providers

Check your email provider's documentation for their SMTP settings. Common settings include:

- SMTP Server: usually `smtp.yourprovider.com`
- Port: `587` (with TLS) or `465` (with SSL)
- Authentication: Required (your email and password)

## Troubleshooting

### Connection Issues
- Verify your SMTP server and port settings
- Check if your firewall is blocking the connection
- Try using a different network connection

### Authentication Errors
- Double-check your username and password
- For Gmail, ensure you're using an App Password if 2FA is enabled
- Some providers require enabling "Less secure app access"

### Sending Limitations
- Most email providers have daily sending limits
- Gmail: 500 emails per day for free accounts
- Avoid sending too many emails too quickly to prevent being flagged as spam
