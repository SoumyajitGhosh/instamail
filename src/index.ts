import nodemailer from 'nodemailer';
import dotenv from 'dotenv';
import xlsx from 'xlsx';
import axios from 'axios';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

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

async function readExcelFromURL(url: string): Promise<any[]> {
    try {
        const response = await axios.get(url, { responseType: 'arraybuffer' });
        const buffer = Buffer.from(response.data);
        const workbook = xlsx.read(buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(sheet);
        return data;
    } catch (error) {
        console.error('Error reading Excel file from URL:', error);
        return [];
    }
}

(async () => {
    try {
        if (!process.env.FILE_URL) {
            throw new Error("Environment variable 'FILE_URL' is not defined.");
        }

        const excelData = await readExcelFromURL(process.env.FILE_URL);
        console.log('Excel Data:', excelData);

        const recipients = excelData.filter((row: any) => row.Email && row.Name && row.Company);

        if (!process.env.CV_PATH) {
            throw new Error("Environment variable 'CV_PATH' is not defined.");
        }

        const __filename = fileURLToPath(import.meta.url);
        const __dirname = path.dirname(__filename);
        const resumePath = path.resolve(__dirname, process.env.CV_PATH);

        for (const recipient of recipients) {
            const { Email: email, Name: name, Company: company } = recipient;

            const htmlContent = `<p>Dear ${name},</p>

                <p>I hope this email finds you well. My name is <strong>Soumyajit Ghosh</strong>, and I’m a software developer with over three years of experience in designing and building scalable web applications. I am reaching out to express my interest in contributing to the work being done at <strong>${company}</strong>.</p>

                <p>With expertise in <strong>React, Node.js, and PostgreSQL</strong>, I have successfully delivered solutions that improve efficiency and user experience. I am confident that my skills and experience can align with your team’s goals, and I am excited about the opportunity to contribute to your projects.</p>

                <p>I’ve attached my resume for your reference and would love the chance to connect and learn more about how I can add value to your team.</p>

                <p>Thank you for your time and consideration.</p>

                <p>Best regards,</p>

                <p>
                    <strong>Soumyajit Ghosh</strong><br>
                    📞 +91-9674447085<br>
                    📧 <a href="mailto:soumyajitghosh.official@gmail.com">soumyajitghosh.official@gmail.com</a><br>
                    🌐 <a href="https://www.linkedin.com/in/soumyajit-ghosh-17b719163/" target="_blank">LinkedIn</a> | 
                    <a href="https://portfolio.soumyajitghosh.life" target="_blank">Portfolio</a>
                </p>`;

            await sendEmail(
                email,
                `Software Developer with 3+ Years of Experience – Interested in Opportunities at ${company}`,
                '',
                htmlContent,
                [
                    {
                        path: resumePath,
                    }
                ]
            );

            await new Promise(resolve => setTimeout(resolve, 10000));
        }

        console.log('All emails sent successfully.');
    } catch (err) {
        console.error('Error:', err);
    }
})();
