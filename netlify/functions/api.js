const express = require('express');
const serverless = require('serverless-http');
const cors = require('cors');
const { graphApiHandler } = require('./graphApi');
const { authHandler } = require('./auth');

const app = express();
app.use(cors());
app.use(express.json());

// Routes
app.post('/api/composeExtension/command', graphApiHandler);
app.post('/api/composeExtension/taskModule', graphApiHandler);
app.get('/auth/callback', authHandler);
app.get('/auth/teams/callback', authHandler);

module.exports.handler = serverless(app); 