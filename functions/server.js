const express = require("express");
const serverless = require("serverless-http");
const multer = require("multer");
const session = require("express-session");
const cors = require("cors");
const axios = require("axios");
require("dotenv").config();

const app = express();

// Middleware
app.use(express.json());
app.use(cors({
    origin: [
        'https://localhost:3000',
        'https://teams.microsoft.com',
        'http://localhost:3000',
        'https://*.teams.microsoft.com'
    ],
    credentials: true,
    methods: ['GET', 'POST', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization']
}));
app.use(session({
    secret: process.env.SESSION_SECRET || "your-secret-key",
    resave: false,
    saveUninitialized: true,
    cookie: {
        secure: true,
        httpOnly: true,
        sameSite: 'none'
    }
}));

// Routes
app.get("/record", (req, res) => {
    res.send("This is the recording page. Implement your recording UI here.");
});

app.post("/sendMessage", multer().single("voiceNote"), async (req, res) => {
    const accessToken = req.session.accessToken;
    const chatId = req.body.chatId;

    if (!accessToken) {
        return res.status(401).json({ error: "User not authenticated. Please log in at /login" });
    }

    if (!chatId || !req.file) {
        return res.status(400).json({ error: "Chat ID and voice message file are required" });
    }

    try {
        const uploadUrl = `https://graph.microsoft.com/v1.0/users/${req.session.userId}/chats/${chatId}/messages`;
        const messageBody = {
            content: "Voice message",
            attachments: [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": "voice-message.webm",
                    "contentBytes": req.file.buffer.toString('base64'),
                    "contentType": "audio/webm"
                }
            ]
        };

        const response = await axios.post(uploadUrl, messageBody, {
            headers: {
                Authorization: `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            }
        });

        res.json({ message: "Voice message sent successfully", data: response.data });
    } catch (error) {
        res.status(500).json({ error: "Failed to send voice message", details: error.message });
    }
});

// Export the serverless function
module.exports.handler = serverless(app);
