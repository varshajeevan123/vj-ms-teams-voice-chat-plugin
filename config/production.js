module.exports = {
    port: process.env.PORT || 3000,
    ssl: {
        enabled: true,
        keyPath: process.env.SSL_KEY_PATH || 'ssl/private.key',
        certPath: process.env.SSL_CERT_PATH || 'ssl/certificate.crt'
    },
    cors: {
        origin: [
            'https://teams.microsoft.com',
            'https://*.teams.microsoft.com',
            process.env.APP_URL
        ],
        credentials: true,
        methods: ['GET', 'POST', 'OPTIONS'],
        allowedHeaders: ['Content-Type', 'Authorization']
    },
    session: {
        secret: process.env.SESSION_SECRET,
        resave: false,
        saveUninitialized: true,
        cookie: {
            secure: true,
            httpOnly: true,
            sameSite: 'none'
        }
    },
    graphApi: {
        url: 'https://graph.microsoft.com/v1.0',
        scopes: [
            'Chat.Read',
            'Chat.ReadWrite',
            'ChatMessage.Send',
            'User.Read'
        ]
    }
}; 