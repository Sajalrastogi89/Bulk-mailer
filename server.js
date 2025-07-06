const express = require('express');
const nodemailer = require('nodemailer');
const path = require('path');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Email queue and status tracking
let emailQueue = [];
let emailStatus = {
    total: 0,
    sent: 0,
    failed: 0,
    pending: 0,
    inProgress: false,
    startTime: null,
    results: []
};

// Configure multer for file uploads
const upload = multer({
    dest: 'uploads/',
    fileFilter: (req, file, cb) => {
        const allowedTypes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
            'application/vnd.ms-excel', // .xls
            'text/csv' // .csv
        ];
        if (allowedTypes.includes(file.mimetype)) {
            cb(null, true);
        } else {
            cb(new Error('Only Excel (.xlsx, .xls) and CSV files are allowed'), false);
        }
    },
    limits: {
        fileSize: 5 * 1024 * 1024 // 5MB limit
    }
});

// Middleware
app.use(express.json());
app.use(express.static(path.join(__dirname)));

// Create transporter for nodemailer
// You'll need to configure this with your email provider
const createTransporter = () => {
    return nodemailer.createTransport({
        // For Gmail
        service: 'gmail',
        auth: {
            user: process.env.EMAIL_USER, // Your email
            pass: process.env.EMAIL_PASS  // Your app password (not regular password)
        }
        
        // Alternative configuration for other email providers:
        /*
        host: process.env.SMTP_HOST,
        port: process.env.SMTP_PORT,
        secure: process.env.SMTP_SECURE === 'true',
        auth: {
            user: process.env.EMAIL_USER,
            pass: process.env.EMAIL_PASS
        }
        */
    });
};

// Route to serve the HTML file
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Route to handle email sending
app.post('/send-email', async (req, res) => {
    try {
        const { to, subject, message } = req.body;
        
        // Validate input
        if (!to || !subject || !message) {
            return res.status(400).json({
                error: 'Please provide all required fields: to, subject, and message'
            });
        }
        
        // Email validation
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(to)) {
            return res.status(400).json({
                error: 'Please provide a valid email address'
            });
        }
        
        const transporter = createTransporter();
        
        // Mail options
        const mailOptions = {
            from: process.env.EMAIL_USER,
            to: to,
            subject: subject,
            html: `
                <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
                    <h2 style="color: #333;">New Message</h2>
                    <div style="background-color: #f9f9f9; padding: 20px; border-radius: 5px;">
                        <p style="margin: 0; line-height: 1.6;">${message.replace(/\n/g, '<br>')}</p>
                    </div>
                    <hr style="margin: 20px 0; border: none; border-top: 1px solid #ddd;">
                    <p style="color: #666; font-size: 12px;">
                        This email was sent from the Email Sender application.
                    </p>
                </div>
            `,
            text: message // Plain text version
        };
        
        // Send email
        const info = await transporter.sendMail(mailOptions);
        
        console.log('Email sent successfully:', info.messageId);
        
        res.json({
            success: true,
            message: 'Email sent successfully!',
            messageId: info.messageId
        });
        
    } catch (error) {
        console.error('Error sending email:', error);
        
        let errorMessage = 'Failed to send email. Please try again.';
        
        // Handle specific error types
        if (error.code === 'EAUTH') {
            errorMessage = 'Authentication failed. Please check your email credentials.';
        } else if (error.code === 'ENOTFOUND') {
            errorMessage = 'Network error. Please check your internet connection.';
        } else if (error.responseCode === 535) {
            errorMessage = 'Invalid email credentials. Please check your username and password.';
        }
        
        res.status(500).json({
            error: errorMessage,
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// Utility functions
const getRandomDelay = (min = 5000, max = 30000) => {
    return Math.floor(Math.random() * (max - min + 1)) + min;
};

const parseSpreadsheet = (filePath) => {
    try {
        const workbook = xlsx.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet);
        
        // Extract emails from various possible column names
        const emailColumns = ['email', 'Email', 'EMAIL', 'mail', 'Mail', 'MAIL', 'e-mail', 'E-mail', 'E-MAIL'];
        const nameColumns = ['name', 'Name', 'NAME', 'first_name', 'firstName', 'First Name', 'full_name', 'fullName', 'Full Name'];
        
        const emails = data.map((row, index) => {
            let email = null;
            let name = null;
            
            // Find email column
            for (const col of emailColumns) {
                if (row[col]) {
                    email = row[col];
                    break;
                }
            }
            
            // Find name column
            for (const col of nameColumns) {
                if (row[col]) {
                    name = row[col];
                    break;
                }
            }
            
            return {
                email: email ? email.toString().trim() : null,
                name: name ? name.toString().trim() : null,
                rowIndex: index + 1
            };
        }).filter(item => item.email && validateEmail(item.email));
        
        return emails;
    } catch (error) {
        throw new Error(`Failed to parse spreadsheet: ${error.message}`);
    }
};

const validateEmail = (email) => {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
};

// Function to add email-safe CSS to prevent automatic link styling
const addEmailSafeCSS = (htmlTemplate) => {
    const emailSafeCSS = `
    <style>
        /* Prevent email clients from auto-styling text as links */
        a, a:link, a:visited, a:hover, a:active {
            color: inherit !important;
            text-decoration: none !important;
        }
        /* Force text color consistency */
        p, div, span, td, th, h1, h2, h3, h4, h5, h6 {
            color: inherit !important;
        }
        /* Prevent automatic link detection styling */
        .no-link, .prevent-auto-link {
            color: inherit !important;
            text-decoration: none !important;
            pointer-events: none !important;
        }
        /* Override email client default link colors */
        body a, body a:link, body a:visited {
            color: inherit !important;
            text-decoration: none !important;
        }
        /* Specific fix for text after images */
        img + p, img + div, img + span {
            color: inherit !important;
        }
    </style>
    `;
    
    // Check if the HTML already has a <head> section
    if (htmlTemplate.includes('<head>')) {
        // Insert the CSS before the closing </head> tag
        return htmlTemplate.replace('</head>', emailSafeCSS + '</head>');
    } else if (htmlTemplate.includes('<html>')) {
        // If there's an <html> tag but no <head>, add a head section
        return htmlTemplate.replace('<html>', '<html><head>' + emailSafeCSS + '</head>');
    } else {
        // If it's just HTML content without proper structure, wrap it
        return `<!DOCTYPE html><html><head>${emailSafeCSS}</head><body>${htmlTemplate}</body></html>`;
    }
};

const sendBulkEmails = async (emails, subject, message, htmlTemplate = null) => {
    emailStatus.inProgress = true;
    emailStatus.total = emails.length;
    emailStatus.sent = 0;
    emailStatus.failed = 0;
    emailStatus.pending = emails.length;
    emailStatus.startTime = new Date();
    emailStatus.results = [];
    
    const transporter = createTransporter();
    
    for (let i = 0; i < emails.length; i++) {
        const emailData = emails[i];
        
        try {
            // Random delay between emails (5-30 seconds)
            if (i > 0) {
                const delay = getRandomDelay();
                console.log(`Waiting ${delay/1000} seconds before sending next email...`);
                await new Promise(resolve => setTimeout(resolve, delay));
            }
            
            // Personalize message and HTML template
            let personalizedMessage = message;
            let personalizedHtml = htmlTemplate;
            
            if (emailData.name) {
                personalizedMessage = `Dear ${emailData.name},\n\n${message}`;
            }
            
            // Replace placeholders in HTML template if provided
            if (personalizedHtml) {
                personalizedHtml = personalizedHtml
                    .replace(/\{\{name\}\}/g, emailData.name || 'Valued Customer')
                    .replace(/\{\{email\}\}/g, emailData.email)
                    .replace(/\{\{company\}\}/g, emailData.company || 'Your Company');
            } else {
                // Use default HTML template
                personalizedHtml = `
                    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
                        ${emailData.name ? `<p>Dear ${emailData.name},</p>` : ''}
                        <div style="background-color: #f9f9f9; padding: 20px; border-radius: 5px;">
                            <p style="margin: 0; line-height: 1.6;">${message.replace(/\n/g, '<br>')}</p>
                        </div>
                        <hr style="margin: 20px 0; border: none; border-top: 1px solid #ddd;">
                        <p style="color: #666; font-size: 12px;">
                            This email was sent from the Email Sender application.
                        </p>
                    </div>
                `;
            }
            
            const mailOptions = {
                from: process.env.EMAIL_USER,
                to: emailData.email,
                subject: subject,
                html: personalizedHtml,
                text: personalizedMessage
            };
            
            const info = await transporter.sendMail(mailOptions);
            
            emailStatus.sent++;
            emailStatus.pending--;
            emailStatus.results.push({
                email: emailData.email,
                name: emailData.name,
                status: 'sent',
                messageId: info.messageId,
                timestamp: new Date()
            });
            
            console.log(`✅ Email sent to ${emailData.email} (${emailStatus.sent}/${emailStatus.total})`);
            
        } catch (error) {
            emailStatus.failed++;
            emailStatus.pending--;
            emailStatus.results.push({
                email: emailData.email,
                name: emailData.name,
                status: 'failed',
                error: error.message,
                timestamp: new Date()
            });
            
            console.error(`❌ Failed to send email to ${emailData.email}:`, error.message);
        }
    }
    
    emailStatus.inProgress = false;
    console.log(`📊 Bulk email completed: ${emailStatus.sent} sent, ${emailStatus.failed} failed`);
};

// Route to upload spreadsheet and send bulk emails
app.post('/upload-and-send', upload.single('spreadsheet'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'Please upload a spreadsheet file' });
        }
        
        const { subject, message } = req.body;
        
        if (!subject || !message) {
            // Clean up uploaded file
            fs.unlinkSync(req.file.path);
            return res.status(400).json({
                error: 'Please provide both subject and message'
            });
        }
        
        // Parse the spreadsheet
        const emails = parseSpreadsheet(req.file.path);
        
        // Clean up uploaded file
        fs.unlinkSync(req.file.path);
        
        if (emails.length === 0) {
            return res.status(400).json({
                error: 'No valid email addresses found in the spreadsheet. Please check your file format.'
            });
        }
        
        // Start sending emails asynchronously
        sendBulkEmails(emails, subject, message).catch(console.error);
        
        res.json({
            success: true,
            message: `Bulk email process started! Found ${emails.length} valid email addresses.`,
            emailCount: emails.length,
            emails: emails.map(e => ({ email: e.email, name: e.name }))
        });
        
    } catch (error) {
        console.error('Error in bulk email process:', error);
        
        // Clean up uploaded file if it exists
        if (req.file && fs.existsSync(req.file.path)) {
            fs.unlinkSync(req.file.path);
        }
        
        res.status(500).json({
            error: error.message || 'Failed to process spreadsheet'
        });
    }
});

// Route to get email sending status
app.get('/email-status', (req, res) => {
    res.json(emailStatus);
});

// Route to stop email sending (if needed)
app.post('/stop-emails', (req, res) => {
    // Note: This is a simple implementation. In production, you'd want more sophisticated queue management
    emailStatus.inProgress = false;
    res.json({ message: 'Email sending process stopped' });
});

// Route to send single email with custom HTML template
app.post('/send-template', upload.single('spreadsheet'), async (req, res) => {
    try {
        const { to, subject, htmlTemplate, plainText } = req.body;
        
        // Validate input
        if (!to || !subject || !htmlTemplate) {
            return res.status(400).json({
                error: 'Please provide all required fields: to, subject, and htmlTemplate'
            });
        }
        
        // Email validation
        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(to)) {
            return res.status(400).json({
                error: 'Please provide a valid email address'
            });
        }
        
        const transporter = createTransporter();
        
        // Add email-safe CSS to prevent automatic link styling
        const processedHtml = addEmailSafeCSS(htmlTemplate);
        
        // Mail options - send HTML template with email-safe CSS
        const mailOptions = {
            from: process.env.EMAIL_USER,
            to: to,
            subject: subject,
            html: processedHtml,
            text: plainText || htmlTemplate.replace(/<[^>]*>/g, '') // Strip HTML for plain text
        };
        
        // Send email
        const info = await transporter.sendMail(mailOptions);
        
        console.log('Template email sent successfully:', info.messageId);
        
        res.json({
            success: true,
            message: 'Template email sent successfully!',
            messageId: info.messageId
        });
        
    } catch (error) {
        console.error('Error sending template email:', error);
        
        let errorMessage = 'Failed to send template email. Please try again.';
        
        // Handle specific error types
        if (error.code === 'EAUTH') {
            errorMessage = 'Authentication failed. Please check your email credentials.';
        } else if (error.code === 'ENOTFOUND') {
            errorMessage = 'Network error. Please check your internet connection.';
        } else if (error.responseCode === 535) {
            errorMessage = 'Invalid email credentials. Please check your username and password.';
        }
        
        res.status(500).json({
            error: errorMessage,
            details: process.env.NODE_ENV === 'development' ? error.message : undefined
        });
    }
});

// Route to send bulk emails with custom HTML template
app.post('/send-template-bulk', upload.single('spreadsheet'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'Please upload a spreadsheet file' });
        }
        
        const { subject, htmlTemplate, plainText } = req.body;
        
        if (!subject || !htmlTemplate) {
            // Clean up uploaded file
            fs.unlinkSync(req.file.path);
            return res.status(400).json({
                error: 'Please provide both subject and HTML template'
            });
        }
        
        // Parse the spreadsheet
        const emails = parseSpreadsheet(req.file.path);
        
        // Clean up uploaded file
        fs.unlinkSync(req.file.path);
        
        if (emails.length === 0) {
            return res.status(400).json({
                error: 'No valid email addresses found in the spreadsheet. Please check your file format.'
            });
        }
        
        // Start sending template emails asynchronously
        sendBulkTemplateEmails(emails, subject, htmlTemplate, plainText).catch(console.error);
        
        res.json({
            success: true,
            message: `Bulk template email process started! Found ${emails.length} valid email addresses.`,
            emailCount: emails.length,
            emails: emails.map(e => ({ email: e.email, name: e.name }))
        });
        
    } catch (error) {
        console.error('Error in bulk template email process:', error);
        
        // Clean up uploaded file if it exists
        if (req.file && fs.existsSync(req.file.path)) {
            fs.unlinkSync(req.file.path);
        }
        
        res.status(500).json({
            error: error.message || 'Failed to process spreadsheet for template emails'
        });
    }
});

// Enhanced bulk email function for templates
const sendBulkTemplateEmails = async (emails, subject, htmlTemplate, plainText = null) => {
    emailStatus.inProgress = true;
    emailStatus.total = emails.length;
    emailStatus.sent = 0;
    emailStatus.failed = 0;
    emailStatus.pending = emails.length;
    emailStatus.startTime = new Date();
    emailStatus.results = [];
    
    const transporter = createTransporter();
    
    for (let i = 0; i < emails.length; i++) {
        const emailData = emails[i];
        
        try {
            // Random delay between emails (5-30 seconds)
            if (i > 0) {
                const delay = getRandomDelay();
                console.log(`Waiting ${delay/1000} seconds before sending next email...`);
                await new Promise(resolve => setTimeout(resolve, delay));
            }
            
            // Add email-safe CSS to prevent automatic link styling
            let processedHtml = addEmailSafeCSS(htmlTemplate);
            
            // Process plain text - send directly without modification
            let processedPlainText = plainText;
            if (!processedPlainText) {
                // Strip HTML tags for plain text version only if no plain text provided
                processedPlainText = htmlTemplate.replace(/<[^>]*>/g, '');
            }
            
            const mailOptions = {
                from: process.env.EMAIL_USER,
                to: emailData.email,
                subject: subject,
                html: processedHtml,
                text: processedPlainText
            };
            
            const info = await transporter.sendMail(mailOptions);
            
            emailStatus.sent++;
            emailStatus.pending--;
            emailStatus.results.push({
                email: emailData.email,
                name: emailData.name,
                status: 'sent',
                messageId: info.messageId,
                timestamp: new Date()
            });
            
            console.log(`✅ Template email sent to ${emailData.email} (${emailStatus.sent}/${emailStatus.total})`);
            
        } catch (error) {
            emailStatus.failed++;
            emailStatus.pending--;
            emailStatus.results.push({
                email: emailData.email,
                name: emailData.name,
                status: 'failed',
                error: error.message,
                timestamp: new Date()
            });
            
            console.error(`❌ Failed to send template email to ${emailData.email}:`, error.message);
        }
    }
    
    emailStatus.inProgress = false;
    console.log(`📊 Bulk template email completed: ${emailStatus.sent} sent, ${emailStatus.failed} failed`);
};

// Error handling middleware
app.use((err, req, res, next) => {
    console.error(err.stack);
    res.status(500).json({
        error: 'Something went wrong!'
    });
});

// Start server
app.listen(PORT, () => {
    console.log(`🚀 Server running on http://localhost:${PORT}`);
    console.log('📧 Email sender application is ready!');
    
    // Check if environment variables are set
    if (!process.env.EMAIL_USER || !process.env.EMAIL_PASS) {
        console.warn('⚠️  Warning: EMAIL_USER and EMAIL_PASS environment variables are not set!');
        console.log('Please create a .env file with your email credentials.');
    }
});
