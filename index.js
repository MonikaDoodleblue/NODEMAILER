const express = require('express');
const multer = require('multer');
const fs = require('fs');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
require("dotenv").config();
const port = process.env.PORT;
const app = express();
const upload = multer();

app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    // Read the .xlsx file
    const workbook = xlsx.read(req.file.buffer);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];     //Sheet1
    const emails = xlsx.utils.sheet_to_json(sheet, { header: 1 })     //data's in .xlsx file
      .flat()
      .filter(email => /\S+@\S+\.\S+/.test(email));

    if (emails.length == 0) {
      res.send({
        Status: 404,
        Message: "Not Found"
      });
    }

    // Store the email IDs in .text file
    const emailIds = emails.map((email) => email.toLowerCase());
    await fs.promises.writeFile('emails.txt', emailIds.join('\n'));
    console.log("emailIds----------", emailIds);
    res.send({
      Status: 200,
      Message: "Mail Id's Stored and Message Sent Successfully"
    });

    async function sendEmails() {
      // Read the email Id's from the file
      await fs.promises.readFile('emails.txt');
      // Create a transporter object
      const transporter = nodemailer.createTransport({
        host: process.env.HOST,
        port: process.env.PORTNO,
        secure: process.env.SECURE,
        auth: {
          user: process.env.USER,
          pass: process.env.PASS,
        }
      });

      // Prepare the email
      const mailOptions = {
        from: 'moniics03@gmail.com',
        to: emailIds,
        subject: 'Subject',
        text: 'Hello, this is a message sent to multiple email addresses.',
        attachments: [
          {
            filename: 'data.txt',
            path: 'C:/Users/Monika/OneDrive/Documents/nodeMailer/data.txt'
          }
        ]
      };

      // Send the email
      await transporter.sendMail(mailOptions);
      console.log('Emails sent successfully');
    }
    sendEmails();

  } catch (error) {
    res.send({
      Status: 500,
      Message: "Internal Server Error"
    });
  }
});

app.listen(port, () => {
  console.log(`Listening on port ${port}`);
});