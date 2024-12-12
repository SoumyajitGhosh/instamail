import nodemailer from 'nodemailer';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';
import path from 'path';

dotenv.config();

const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        type: 'OAuth2',
        user: process.env.SMTP_USER,
        clientId: process.env.CLIENT_ID,
        clientSecret: process.env.CLIENT_SECRET,
        refreshToken: process.env.REFRESH_TOKEN,
        accessToken: process.env.ACCESS_TOKEN,
    },
});

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function sendEmail(to: string, subject: string, text: string, html?: string, attachments?: any[]): Promise<void> {
    try {
        const mailOptions = {
            from: process.env.FROM_EMAIL || 'no-reply@example.com',
            to,
            subject,
            text,
            html,
            attachments,
        };

        const info = await transporter.sendMail(mailOptions);
        console.log('Email sent:', info.response);
    } catch (error) {
        console.error('Error sending email:', error);
    }
}

(async () => {
    try {
        const absolutePath = path.resolve(__dirname, '../../../../Downloads/Resume-SoumyajitGhosh.pdf');
        console.log('Attachment absolute path:', absolutePath);

        await sendEmail(
            'sjittorres2@gmail.com',
            'Test Email Subject',
            'This is a test email sent from Node.js using Nodemailer.',
            '<p>This is a <b>test email</b> sent from Node.js using <i>Nodemailer</i>.</p>',
            [
                {
                    path: absolutePath,
                }
            ]
        );
    } catch (err) {
        console.error('Error:', err);
    }
})();
