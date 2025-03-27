const express = require("express");
const axios = require("axios");
const multer = require("multer");
const path = require("path");
const session = require("express-session");
const cors = require("cors");
const https = require("https");
const http = require("http");
const fs = require("fs");
require("dotenv").config();

const app = express();
const port = 3000;

const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const TENANT_ID = process.env.TENANT_ID;
const REDIRECT_URI = "http://localhost:3000/auth/callback";
const sessionSecret = process.env.SESSION_SECRET;
const GRAPH_API_URL = "https://graph.microsoft.com/v1.0";

// Configure multer for handling file uploads
const upload = multer({
    storage: multer.memoryStorage(),
    limits: {
        fileSize: 10 * 1024 * 1024 // 10MB limit
    }
});

// Enable CORS for all routes
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

// Session Configuration
app.use(
    session({
        secret: "your-secret-key",
        resave: false,
        saveUninitialized: true,
        cookie: {
            secure: true,
            httpOnly: true,
            sameSite: 'none'
        }
    })
);

// Middleware to parse JSON bodies
app.use(express.json());

// Enable static files (for serving index.html)
app.use(express.static("public"));

console.log("🚀 Server is starting...");

// 📌 Route: Health check endpoint
app.get("/health", (req, res) => {
    res.json({ status: "healthy" });
});

// 📌 Route: Root endpoint
app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, "public", "index.html"));
});

// 📌 Route: Send voice message to Teams chat
app.post("/sendMessage", upload.single("voiceNote"), async (req, res) => {
    console.log("📢 Received voice message to send");
    const accessToken = req.session.accessToken;
    const chatId = req.body.chatId;

    if (!accessToken) {
        console.log("⚠️ User not authenticated!");
        return res.status(401).json({ error: "User not authenticated. Please log in at /login" });
    }

    if (!chatId) {
        return res.status(400).json({ error: "Chat ID is required" });
    }

    if (!req.file) {
        return res.status(400).json({ error: "No voice message file provided" });
    }

    try {
        // First, upload the file to Teams
        const uploadUrl = `${GRAPH_API_URL}/users/${req.session.userId}/chats/${chatId}/messages`;
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

        console.log("✅ Successfully sent voice message!");
        res.json({ message: "Voice message sent successfully", data: response.data });
    } catch (error) {
        console.error("❌ Error sending voice message:", error.response?.data || error.message);
        res.status(500).json({ error: "Failed to send voice message" });
    }
});

// 📌 Route: Login endpoint
app.get("/login", async (req, res) => {
    console.log("🔑 Starting login process...");
    const clientId = "f984ebaf-4c50-4de8-8687-80672674ab06";
    const redirectUri = "https://localhost:3000/auth/callback";
    const scope = "Chat.ReadWrite ChatMessage.Send User.Read";

    const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=code&redirect_uri=${encodeURIComponent(redirectUri)}&scope=${encodeURIComponent(scope)}&response_mode=query`;

    console.log("🔗 Redirecting to Microsoft login...");
    res.redirect(authUrl);
});

// 📌 Route: Auth callback endpoint
app.get("/auth/callback", async (req, res) => {
    console.log("🔄 Received auth callback");
    const { code } = req.query;

    if (!code) {
        console.error("❌ No code received in callback");
        return res.status(400).send("No code received");
    }

    try {
        const clientId = "f984ebaf-4c50-4de8-8687-80672674ab06";
        const clientSecret = "your-client-secret";
        const redirectUri = "https://localhost:3000/auth/callback";

        const tokenResponse = await axios.post(
            "https://login.microsoftonline.com/common/oauth2/v2.0/token",
            new URLSearchParams({
                client_id: clientId,
                client_secret: clientSecret,
                code: code,
                redirect_uri: redirectUri,
                grant_type: "authorization_code",
            }),
            {
                headers: {
                    "Content-Type": "application/x-www-form-urlencoded",
                },
            }
        );

        const { access_token } = tokenResponse.data;
        req.session.accessToken = access_token;

        console.log("✅ Successfully obtained access token");
        res.redirect("/");
    } catch (error) {
        console.error("❌ Error in auth callback:", error.response?.data || error.message);
        res.status(500).send("Authentication failed");
    }
});

// 📌 Route: Fetch Teams chats using stored token
app.get("/getChats", async (req, res) => {
    console.log("📨 Received request to fetch Teams chats");
    const accessToken = req.session.accessToken;

    if (!accessToken) {
        console.log("⚠️ User not authenticated!");
        return res.status(401).json({ error: "User not authenticated. Please log in at /login" });
    }

    try {
        console.log("🔄 Fetching chats from Microsoft Graph API...");
        let allChats = [];
        let nextLink = `${GRAPH_API_URL}/me/chats?$top=100&$expand=lastMessagePreview,members&$orderby=lastMessagePreview/createdDateTime desc`;

        while (nextLink) {
            console.log("Fetching from URL:", nextLink);
            const response = await axios.get(nextLink, {
                headers: { 
                    Authorization: `Bearer ${accessToken}`,
                    'Content-Type': 'application/json'
                }
            });

            const chats = response.data.value || [];
            allChats = [...allChats, ...chats];
            
            nextLink = response.data['@odata.nextLink'];
            
            console.log(`✅ Fetched ${chats.length} chats. Total so far: ${allChats.length}`);
            
            await new Promise(resolve => setTimeout(resolve, 100));
        }

        allChats.sort((a, b) => {
            const dateA = new Date(a.lastMessagePreview?.createdDateTime || 0);
            const dateB = new Date(b.lastMessagePreview?.createdDateTime || 0);
            return dateB - dateA;
        });

        console.log(`✅ Successfully fetched all ${allChats.length} chats!`);
        res.json({ value: allChats });
    } catch (error) {
        console.error("❌ Error fetching chats:", error.response?.data || error.message);
        res.status(500).json({ error: "Failed to fetch chat data" });
    }
});

// 📌 Route: Handle compose extension command
app.post("/composeExtension/command", async (req, res) => {
    console.log("📨 Received compose extension command");
    console.log("Request body:", JSON.stringify(req.body, null, 2));
    
    const accessToken = req.session.accessToken;

    if (!accessToken) {
        console.log("⚠️ User not authenticated!");
        return res.status(401).json({ error: "User not authenticated" });
    }

    try {
        res.json({
            task: {
                type: "continue",
                value: {
                    card: {
                        type: "AdaptiveCard",
                        version: "1.2",
                        body: [
                            {
                                type: "TextBlock",
                                text: "Record Voice Note",
                                weight: "bolder",
                                size: "large"
                            },
                            {
                                type: "TextBlock",
                                text: "Click the button below to start recording your voice note"
                            }
                        ],
                        actions: [
                            {
                                type: "Action.Submit",
                                title: "Start Recording",
                                data: {
                                    action: "startRecording"
                                }
                            }
                        ]
                    }
                }
            }
        });
    } catch (error) {
        console.error("❌ Error handling compose extension command:", error);
        res.status(500).json({ 
            error: "Failed to process command",
            details: error.message 
        });
    }
});

// 📌 Route: Handle task module submission
app.post("/composeExtension/submit", async (req, res) => {
    console.log("📨 Received task module submission");
    const { action, data } = req.body;

    if (action === "startRecording") {
        if (!data.chatId) {
            return res.status(400).json({ error: "Please select a recipient" });
        }

        // Store the selected chat ID in the session
        req.session.selectedChatId = data.chatId;

        res.json({
            task: {
                type: "continue",
                value: {
                    card: {
                        type: "AdaptiveCard",
                        version: "1.2",
                        body: [
                            {
                                type: "TextBlock",
                                text: "Recording Voice Note",
                                weight: "bolder",
                                size: "large"
                            },
                            {
                                type: "TextBlock",
                                text: "Recording in progress...",
                                color: "attention"
                            }
                        ],
                        actions: [
                            {
                                type: "Action.Submit",
                                title: "Stop Recording",
                                data: {
                                    action: "stopRecording"
                                }
                            }
                        ]
                    }
                }
            }
        });
    } else if (action === "stopRecording") {
        res.json({
            task: {
                type: "continue",
                value: {
                    card: {
                        type: "AdaptiveCard",
                        version: "1.2",
                        body: [
                            {
                                type: "TextBlock",
                                text: "Voice Note Recorded",
                                weight: "bolder",
                                size: "large"
                            },
                            {
                                type: "TextBlock",
                                text: "Your voice note has been recorded. Click Send to share it."
                            }
                        ],
                        actions: [
                            {
                                type: "Action.Submit",
                                title: "Send Voice Note",
                                data: {
                                    action: "sendVoiceNote"
                                }
                            }
                        ]
                    }
                }
            }
        });
    } else if (action === "sendVoiceNote") {
        const chatId = req.session.selectedChatId;
        if (!chatId) {
            return res.status(400).json({ error: "No recipient selected" });
        }

        res.json({
            task: {
                type: "message",
                value: "Voice note sent successfully!"
            }
        });
    } else {
        res.status(400).json({ error: "Invalid action" });
    }
});

// 📌 Route: Store Teams auth token
app.post("/storeToken", express.json(), (req, res) => {
    console.log("📨 Received request to store Teams auth token");
    const { token } = req.body;
    
    if (!token) {
        console.error("❌ No token provided");
        return res.status(400).json({ error: "No token provided" });
    }
    
    req.session.teamsToken = token;
    console.log("✅ Teams auth token stored successfully");
    res.json({ success: true });
});

// Error handling middleware
app.use((err, req, res, next) => {
    console.error("❌ Unhandled error:", err);
    res.status(500).json({ error: "Internal server error" });
});

// SSL configuration
let sslOptions;
try {
    sslOptions = {
        key: fs.readFileSync(path.join(__dirname, 'ssl', 'private.key')),
        cert: fs.readFileSync(path.join(__dirname, 'ssl', 'certificate.crt'))
    };
    console.log("✅ SSL certificates loaded successfully");
} catch (error) {
    console.error("❌ Error loading SSL certificates:", error.message);
    console.log("⚠️ Please ensure you have generated SSL certificates in the ssl directory");
    process.exit(1);
}

// Start both HTTP and HTTPS servers
const httpServer = http.createServer(app);
const httpsServer = https.createServer(sslOptions, app);

const startServer = (server, port) => {
    return new Promise((resolve, reject) => {
        server.on('error', (err) => {
            if (err.code === 'EADDRINUSE') {
                console.log(`⚠️ Port ${port} is in use, trying ${port + 1}`);
                server.listen(port + 1, () => {
                    console.log(`🚀 Server running on port ${port + 1}`);
                    resolve(port + 1);
                });
            } else {
                reject(err);
            }
        });

        server.listen(port, () => {
            console.log(`🚀 Server running on port ${port}`);
            resolve(port);
        });
    });
};

// Start servers
Promise.all([
    startServer(httpServer, port),
    startServer(httpsServer, port)
]).then(([httpPort, httpsPort]) => {
    console.log(`🚀 HTTP Server running at: http://localhost:${httpPort}`);
    console.log(`🚀 HTTPS Server running at: https://localhost:${httpsPort}`);
    console.log("🔒 HTTPS is enabled");
}).catch(err => {
    console.error("❌ Failed to start servers:", err);
    process.exit(1);
});