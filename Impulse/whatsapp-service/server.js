const express = require('express');
const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const qrcode = require('qrcode');
const cors = require('cors');
const fs = require('fs');

const app = express();
app.use(cors());
app.use(express.json());

process.on('uncaughtException', (err) => {
    console.error('Uncaught Exception:', err);
    console.log('Exiting process to allow auto-recovery...');
    process.exit(1);
});

// Locate browser executable dynamically across typical Windows/Server locations
function getBrowserExecutablePath() {
    const candidates = [
        'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
        'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
        'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
        'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
        process.env.PUPPETEER_EXECUTABLE_PATH
    ].filter(Boolean);

    for (const p of candidates) {
        if (fs.existsSync(p)) {
            console.log(`Using browser at: ${p}`);
            return p;
        }
    }
    console.log('No specific browser path found, falling back to Puppeteer default Chromium.');
    return undefined;
}

const browserExe = getBrowserExecutablePath();
const puppeteerConfig = {
    headless: true,
    args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-gpu',
        '--disable-dev-shm-usage',
        '--no-first-run',
        '--no-zygote'
    ]
};
if (browserExe) {
    puppeteerConfig.executablePath = browserExe;
}

const client = new Client({
    authStrategy: new LocalAuth(),
    puppeteer: puppeteerConfig
});

let qrCodeDataUrl = null;
let connectionStatus = 'INITIALIZING'; // INITIALIZING, QR_READY, CONNECTED, DISCONNECTED
let lastErrorMessage = '';

client.on('qr', async (qr) => {
    console.log('QR code received, generating data URL...');
    try {
        qrCodeDataUrl = await qrcode.toDataURL(qr);
        connectionStatus = 'QR_READY';
        lastErrorMessage = '';
    } catch (qrErr) {
        console.error('Error generating QR Data URL:', qrErr);
    }
});

client.on('ready', () => {
    console.log('WhatsApp Client is ready!');
    qrCodeDataUrl = null;
    connectionStatus = 'CONNECTED';
    lastErrorMessage = '';
});

client.on('authenticated', () => {
    console.log('WhatsApp Authenticated!');
    connectionStatus = 'CONNECTED';
    lastErrorMessage = '';
});

client.on('auth_failure', msg => {
    console.error('WhatsApp Authentication failure', msg);
    connectionStatus = 'DISCONNECTED';
    lastErrorMessage = typeof msg === 'string' ? msg : 'Authentication failed';
});

client.on('disconnected', (reason) => {
    console.log('WhatsApp Client was disconnected', reason);
    qrCodeDataUrl = null;
    connectionStatus = 'DISCONNECTED';
    lastErrorMessage = reason || 'Client disconnected';
    
    // Exit process so the auto-restart loop can wipe the corrupted session and cleanly restart
    console.log("Exiting process to allow clean restart...");
    setTimeout(() => process.exit(1), 1000);
});

// API Endpoints
app.get('/status', (req, res) => {
    res.json({ status: connectionStatus, error: lastErrorMessage });
});

app.get('/qr', (req, res) => {
    if (connectionStatus === 'QR_READY' && qrCodeDataUrl) {
        res.json({ qr: qrCodeDataUrl, status: connectionStatus, error: lastErrorMessage });
    } else {
        res.json({ qr: null, status: connectionStatus, error: lastErrorMessage });
    }
});

app.post('/disconnect', async (req, res) => {
    try {
        console.log("Disconnect request received. Logging out and clearing session...");
        connectionStatus = 'DISCONNECTED';
        try {
            await client.logout();
        } catch (e) {
            console.log("Client logout note:", e.message);
        }
        res.json({ success: true, message: 'WhatsApp session disconnected and logged out.' });
        setTimeout(() => process.exit(0), 1000);
    } catch (err) {
        res.status(500).json({ success: false, error: err.message });
    }
});

app.post('/send', async (req, res) => {
    if (connectionStatus !== 'CONNECTED') {
        return res.status(400).json({ success: false, error: 'WhatsApp is not connected' });
    }

    const { phoneNumber, message, attachments } = req.body;
    
    if (!phoneNumber || !message) {
        return res.status(400).json({ success: false, error: 'Phone number and message are required' });
    }

    // Format phone number for whatsapp-web.js
    let formattedNumber = phoneNumber.replace(/[\s+]/g, '');
    if (formattedNumber.startsWith('00')) {
        formattedNumber = formattedNumber.substring(2);
    }
    const chatId = `${formattedNumber}@c.us`;

    try {
        // First send the text message
        await client.sendMessage(chatId, message);
        console.log(`Text message sent to ${phoneNumber}`);

        // Then send attachments if any
        console.log(`Checking attachments for ${phoneNumber}... Length: ${attachments ? attachments.length : 0}`);
        if (attachments && Array.isArray(attachments)) {
            for (const filePath of attachments) {
                try {
                    console.log(`Attempting to send attachment: ${filePath}`);
                    const isAudio = filePath.toLowerCase().endsWith('.webm') || filePath.toLowerCase().endsWith('.ogg') || filePath.toLowerCase().endsWith('.mp3');
                    
                    if (isAudio && filePath.toLowerCase().endsWith('.webm')) {
                        console.log(`Sending webm as audio/mp4 bypass: ${filePath}`);
                        const base64 = fs.readFileSync(filePath, { encoding: 'base64' });
                        const media = new MessageMedia('audio/mp4', base64, filePath.split(/[\\/]/).pop());
                        await client.sendMessage(chatId, media);
                    }
                    else if (isAudio) {
                        const media = MessageMedia.fromFilePath(filePath);
                        console.log(`Sending as voice note: ${filePath}`);
                        await client.sendMessage(chatId, media, { sendAudioAsVoice: true });
                    } else {
                        const media = MessageMedia.fromFilePath(filePath);
                        await client.sendMessage(chatId, media);
                    }
                    console.log(`Attachment sent to ${phoneNumber}: ${filePath}`);
                } catch (mediaError) {
                    console.error(`Failed to send attachment ${filePath}:`, mediaError);
                }
            }
        }

        res.json({ success: true });
    } catch (error) {
        console.error('Error sending message:', error);
        res.status(500).json({ success: false, error: error.toString() });
    }
});

const PORT = 3001;
// Listen immediately so health checks and status endpoints are reachable right away
app.listen(PORT, () => {
    console.log(`WhatsApp Microservice running on port ${PORT}`);
    
    // Initialize WhatsApp client in the background after Express is active
    console.log('Initializing WhatsApp Client...');
    client.initialize().catch(err => {
        console.error('Failed to initialize WhatsApp client:', err);
        connectionStatus = 'DISCONNECTED';
        lastErrorMessage = err.message || 'Initialization failed';
    });
});
