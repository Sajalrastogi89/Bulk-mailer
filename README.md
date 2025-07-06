# Email Sender Application

A comprehensive web application that allows users to send single emails or bulk emails from spreadsheets using Node.js, Express, and Nodemailer.

## Features

- **Single Email Sending**: Send individual emails with a modern web form
- **Bulk Email Sending**: Upload spreadsheets (Excel/CSV) and send emails to multiple recipients
- **Custom HTML Templates**: Create and send beautiful HTML email templates with personalization
- **Pre-built Templates**: Choose from newsletter, announcement, and promotional templates
- **Template Preview**: Live preview of email templates before sending
- **Asynchronous Processing**: Emails are sent with random intervals (5-30 seconds) to avoid spam detection
- **Real-time Status Monitoring**: Track the progress of bulk email campaigns
- **Email Personalization**: Automatically personalizes emails with recipient names and other placeholders
- **Error Handling**: Comprehensive error handling and reporting
- **Support for Multiple Email Providers**: Works with Gmail, Outlook, and custom SMTP servers
- **File Upload Validation**: Supports .xlsx, .xls, and .csv file formats
- **Progress Tracking**: Visual progress bar and detailed statistics
- **Template Management**: Save and load custom templates

## Spreadsheet Format

Your spreadsheet should contain at least an email column. The application automatically detects common column names:

### Email Columns (any of these):
- `email`, `Email`, `EMAIL`
- `mail`, `Mail`, `MAIL`  
- `e-mail`, `E-mail`, `E-MAIL`

### Name Columns (optional, for personalization):
- `name`, `Name`, `NAME`
- `first_name`, `firstName`, `First Name`
- `full_name`, `fullName`, `Full Name`

### Example CSV Format:
```csv
Name,Email
John Doe,john.doe@example.com
Jane Smith,jane.smith@example.com
Mike Johnson,mike.johnson@example.com
```

A sample file `sample_emails.csv` is included for reference.

## HTML Email Templates

The application supports custom HTML email templates with personalization placeholders:

### Available Placeholders:
- `{{name}}` - Recipient's name from the spreadsheet
- `{{email}}` - Recipient's email address
- `{{company}}` - Company name (if available in spreadsheet)

### Pre-built Templates:
1. **Newsletter Template**: Professional newsletter layout
2. **Announcement Template**: Eye-catching announcement design
3. **Promotional Template**: Marketing-focused promotional email

### Template Features:
- **Live Preview**: See how your email will look before sending
- **Custom HTML**: Write your own HTML templates
- **Responsive Design**: Templates work on mobile and desktop
- **Placeholder Insertion**: Easy buttons to insert common placeholders
- **Template Saving**: Save templates as JSON files for reuse
- **Plain Text Version**: Auto-generated or custom plain text versions

### Example HTML Template:
```html
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; }
        .container { max-width: 600px; margin: 0 auto; }
        .header { background: #667eea; color: white; padding: 20px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Hello {{name}}!</h1>
        </div>
        <div class="content">
            <p>This is a personalized email for {{email}}</p>
        </div>
    </div>
</body>
</html>
```

## Setup Instructions

### 1. Install Dependencies

```bash
npm install
```

### 2. Configure Email Settings

Edit the `.env` file and replace the placeholder values with your actual email credentials:

```env
EMAIL_USER=your-email@gmail.com
EMAIL_PASS=your-app-password
```

**For Gmail users:**
1. Enable 2-factor authentication on your Google account
2. Generate an "App Password" for this application
3. Use the app password (not your regular password) in the `EMAIL_PASS` field

**For other email providers:**
- Update the transporter configuration in `server.js`
- Set the appropriate SMTP settings in `.env`

### 3. Run the Application

```bash
npm start
```

Or for development with auto-restart:

```bash
npm run dev
```

### 4. Access the Application

Open your browser and go to: `http://localhost:3000`

## Email Provider Configuration

### Gmail
```javascript
service: 'gmail',
auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS
}
```

### Outlook/Hotmail
```javascript
service: 'hotmail',
auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS
}
```

### Custom SMTP
```javascript
host: 'smtp.your-provider.com',
port: 587,
secure: false,
auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS
}
```

## Security Notes

- Never commit your `.env` file to version control
- Use app-specific passwords when available
- Enable 2-factor authentication on your email account
- Consider using environment variables in production

## Troubleshooting

### Authentication Errors
- Check that your email and password are correct
- For Gmail, make sure you're using an app password
- Enable "Less secure app access" if required by your provider

### Network Errors
- Check your internet connection
- Verify SMTP settings for your email provider
- Check firewall settings

## File Structure

```
├── index.html      # Frontend form
├── server.js       # Express server with Nodemailer
├── package.json    # Node.js dependencies
├── .env           # Environment variables (not in git)
└── README.md      # This file
```
