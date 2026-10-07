const express = require('express');
const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const qrcode = require('qrcode');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

process.on('uncaughtException', (err) => {
    console.error('Uncaught Exception:', err);
    console.log('Exiting process to allow auto-recovery...');
    process.exit(1);
});

const client = new Client({
    authStrategy: new LocalAuth(),
    puppeteer: {
        headless: true,
        executablePath: 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    }
});

let qrCodeDataUrl = null;
let connectionStatus = 'INITIALIZING'; // INITIALIZING, QR_READY, CONNECTED, DISCONNECTED

client.on('qr', async (qr) => {
    console.log('QR code received, generating data URL...');
    qrCodeDataUrl = await qrcode.toDataURL(qr);
    connectionStatus = 'QR_READY';
});

client.on('ready', () => {
    console.log('WhatsApp Client is ready!');
    qrCodeDataUrl = null;
    connectionStatus = 'CONNECTED';
});

client.on('authenticated', () => {
    console.log('WhatsApp Authenticated!');
    connectionStatus = 'CONNECTED';
});

client.on('auth_failure', msg => {
    console.error('WhatsApp Authentication failure', msg);
    connectionStatus = 'DISCONNECTED';
});

client.on('disconnected', (reason) => {
    console.log('WhatsApp Client was disconnected', reason);
    qrCodeDataUrl = null;
    connectionStatus = 'DISCONNECTED';
    
    // Exit process so the auto-restart loop can wipe the corrupted session and cleanly restart
    console.log("Exiting process to allow clean restart...");
    setTimeout(() => process.exit(1), 1000);
});

client.initialize();

// API Endpoints
app.get('/status', (req, res) => {
    res.json({ status: connectionStatus });
});

app.get('/qr', (req, res) => {
    if (connectionStatus === 'QR_READY' && qrCodeDataUrl) {
        res.json({ qr: qrCodeDataUrl, status: connectionStatus });
    } else {
        res.json({ qr: null, status: connectionStatus });
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
            const fs = require('fs');
            for (const filePath of attachments) {
                try {
                    console.log(`Attempting to send attachment: ${filePath}`);
                    const isAudio = filePath.toLowerCase().endsWith('.webm') || filePath.toLowerCase().endsWith('.ogg') || filePath.toLowerCase().endsWith('.mp3');
                    
                    if (isAudio && filePath.toLowerCase().endsWith('.webm')) {
                        console.log(`Sending webm as audio/mp4 bypass: ${filePath}`);
                        const base64 = fs.readFileSync(filePath, { encoding: 'base64' });
                        const media = new MessageMedia('audio/mp4', base64, filePath.split(/[\\/]/).pop());
                        await client.sendMessage(chatId, media); // Without sendAudioAsVoice to avoid ffmpeg crash, but it sends as playable audio
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
app.listen(PORT, () => {
    console.log(`WhatsApp Microservice running on port ${PORT}`);
});
