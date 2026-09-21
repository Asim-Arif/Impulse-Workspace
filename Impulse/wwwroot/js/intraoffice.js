window.scrollToBottom = (elementId) => {
    const element = document.getElementById(elementId);
    if (element) {
        element.scrollTop = element.scrollHeight;
    }
};

window.focusElement = (elementId) => {
    const element = document.getElementById(elementId);
    if (element) {
        element.focus();
    }
};

window.downloadFile = (url, fileName) => {
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
};

window.showNotification = (title, body) => {
    if ('Notification' in window && Notification.permission === 'granted') {
        new Notification(title, { body: body, icon: '/favicon.ico' });
    }
};

window.playNotificationSound = () => {
    try {
        const AudioCtx = window.AudioContext || window.webkitAudioContext;
        if (!AudioCtx) return;
        const ctx = new AudioCtx();
        if (ctx.state === 'suspended') {
            ctx.resume();
        }
        const now = ctx.currentTime;

        // Pleasant chime - Tone 1 (E5 - 659.25 Hz)
        const osc1 = ctx.createOscillator();
        const gain1 = ctx.createGain();
        osc1.type = 'sine';
        osc1.frequency.setValueAtTime(659.25, now);
        gain1.gain.setValueAtTime(0.18, now);
        gain1.gain.exponentialRampToValueAtTime(0.001, now + 0.3);
        osc1.connect(gain1);
        gain1.connect(ctx.destination);
        osc1.start(now);
        osc1.stop(now + 0.3);

        // Pleasant chime - Tone 2 (B5 - 987.77 Hz)
        const osc2 = ctx.createOscillator();
        const gain2 = ctx.createGain();
        osc2.type = 'sine';
        osc2.frequency.setValueAtTime(987.77, now + 0.1);
        gain2.gain.setValueAtTime(0.22, now + 0.1);
        gain2.gain.exponentialRampToValueAtTime(0.001, now + 0.55);
        osc2.connect(gain2);
        gain2.connect(ctx.destination);
        osc2.start(now + 0.1);
        osc2.stop(now + 0.55);
    } catch (e) {
        console.warn('Audio chime error:', e);
    }
};

window.playMeetingAlertSound = () => {
    try {
        const AudioCtx = window.AudioContext || window.webkitAudioContext;
        if (!AudioCtx) return;
        const ctx = new AudioCtx();
        if (ctx.state === 'suspended') {
            ctx.resume();
        }
        const now = ctx.currentTime;

        // Distinctive 4-tone ascending meeting reminder chime (E5 -> G#5 -> B5 -> E6)
        const notes = [
            { freq: 659.25, time: 0, dur: 0.18, gain: 0.28 },
            { freq: 830.61, time: 0.16, dur: 0.18, gain: 0.30 },
            { freq: 987.77, time: 0.32, dur: 0.20, gain: 0.32 },
            { freq: 1318.51, time: 0.48, dur: 0.45, gain: 0.35 }
        ];

        notes.forEach(n => {
            const osc = ctx.createOscillator();
            const gain = ctx.createGain();
            osc.type = 'sine';
            osc.frequency.setValueAtTime(n.freq, now + n.time);
            gain.gain.setValueAtTime(n.gain, now + n.time);
            gain.gain.exponentialRampToValueAtTime(0.0001, now + n.time + n.dur);
            osc.connect(gain);
            gain.connect(ctx.destination);
            osc.start(now + n.time);
            osc.stop(now + n.time + n.dur);
        });
    } catch (e) {
        console.warn('Meeting audio alert error:', e);
    }
};

window.requestNotificationPermission = async () => {
    if ('Notification' in window) {
        const permission = await Notification.requestPermission();
        return permission === 'granted';
    }
    return false;
};

window.setTheme = (theme) => {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('theme', theme);
};

window.setFont = (font) => {
    document.documentElement.setAttribute('data-font', font);
    localStorage.setItem('font', font);
};

window.setAccent = (accent) => {
    document.documentElement.setAttribute('data-accent', accent);
    localStorage.setItem('accent', accent);
};

window.setMode = (mode) => {
    document.documentElement.setAttribute('data-mode', mode);
    localStorage.setItem('mode', mode);
};
// Utility functions (consolidated)


window.setStyle = (style) => {
    document.documentElement.setAttribute('data-style', style);
    const customBgColor = localStorage.getItem('customBgColor');
    if (customBgColor) document.documentElement.style.setProperty('--theme-bg', customBgColor);
    const customCaptionColor = localStorage.getItem('customCaptionColor');
    if (customCaptionColor) document.documentElement.style.setProperty('--caption-color', customCaptionColor);
    const customHeadingBgColor = localStorage.getItem('customHeadingBgColor');
    if (customHeadingBgColor) document.documentElement.style.setProperty('--heading-bg', customHeadingBgColor);
    const customHeadingTextColor = localStorage.getItem('customHeadingTextColor');
    if (customHeadingTextColor) document.documentElement.style.setProperty('--heading-color', customHeadingTextColor);
    localStorage.setItem('style', style);
};

window.setDisplayScale = (scale) => {
    document.documentElement.setAttribute('data-scale', scale);
    localStorage.setItem('display-scale', scale);
    const select = document.getElementById('display-scale-select');
    if (select) select.value = scale;
};

window.getTheme = () => {
    return localStorage.getItem('theme') || 'pastel';
};

document.addEventListener('DOMContentLoaded', () => {
    const scale = localStorage.getItem('display-scale') || 'normal';
    document.documentElement.setAttribute('data-scale', scale);
    const scaleSelect = document.getElementById('display-scale-select');
    if (scaleSelect) scaleSelect.value = scale;

    const theme = localStorage.getItem('theme') || 'pastel';
    const font = localStorage.getItem('font') || 'inter';
    const accent = localStorage.getItem('accent') || 'purple';
    const mode = localStorage.getItem('mode') || 'dark';
    const style = localStorage.getItem('style') || 'soft';

    document.documentElement.setAttribute('data-theme', theme);
    document.documentElement.setAttribute('data-font', font);
    document.documentElement.setAttribute('data-accent', accent);
    document.documentElement.setAttribute('data-mode', mode);
    document.documentElement.setAttribute('data-style', style);
    const customBgColor = localStorage.getItem('customBgColor');
    if (customBgColor) document.documentElement.style.setProperty('--theme-bg', customBgColor);
    const customCaptionColor = localStorage.getItem('customCaptionColor');
    if (customCaptionColor) document.documentElement.style.setProperty('--caption-color', customCaptionColor);
    const customHeadingBgColor = localStorage.getItem('customHeadingBgColor');
    if (customHeadingBgColor) document.documentElement.style.setProperty('--heading-bg', customHeadingBgColor);
    const customHeadingTextColor = localStorage.getItem('customHeadingTextColor');
    if (customHeadingTextColor) document.documentElement.style.setProperty('--heading-color', customHeadingTextColor);

    // Login Timer Logic
    let loginSeconds = 124; // start at 00:02:04 based on the screenshot
    setInterval(() => {
        loginSeconds++;
        const h = Math.floor(loginSeconds / 3600).toString().padStart(2, '0');
        const m = Math.floor((loginSeconds % 3600) / 60).toString().padStart(2, '0');
        const s = (loginSeconds % 60).toString().padStart(2, '0');
        const timerElement = document.getElementById('login-timer');
        if(timerElement) {
            timerElement.textContent = `${h}:${m}:${s}`;
        }
    }, 1000);

    // Real-Time System Health Simulation
    let healthValue = 65;
    setInterval(() => {
        // Randomly fluctuate health by -5 to +5
        const change = Math.floor(Math.random() * 11) - 5;
        healthValue = Math.max(10, Math.min(99, healthValue + change));

        const circle = document.getElementById('system-health-circle');
        const valueSpan = document.getElementById('system-health-value');
        const statusSpan = document.getElementById('system-health-status');

        if (circle && valueSpan && statusSpan) {
            valueSpan.textContent = `${healthValue}%`;

            if (healthValue >= 85) {
                // Critical
                circle.style.borderColor = '#e17055';
                circle.style.borderTopColor = 'rgba(225, 112, 85, 0.2)';
                statusSpan.textContent = 'Critical';
                statusSpan.style.color = '#e17055';
            } else if (healthValue >= 50) {
                // Warning
                circle.style.borderColor = '#ffb300';
                circle.style.borderTopColor = 'rgba(255, 179, 0, 0.2)';
                statusSpan.textContent = 'Warning';
                statusSpan.style.color = '#ffb300';
            } else {
                // Good
                circle.style.borderColor = '#1f8a70';
                circle.style.borderTopColor = 'rgba(31, 138, 112, 0.2)';
                statusSpan.textContent = 'Good';
                statusSpan.style.color = '#1f8a70';
            }
        }
    }, 2500);

    // Real-Time Network Speed Simulation
    let baseSpeed = 3.5;
    setInterval(() => {
        const speedSpan = document.getElementById('network-speed-value');
        if (speedSpan) {
            // Random fluctuation between -0.8 and +0.8
            const change = (Math.random() * 1.6) - 0.8;
            let newSpeed = baseSpeed + change;
            
            // Random occasional spike or drop
            if (Math.random() > 0.9) {
                newSpeed = Math.random() > 0.5 ? 8.5 + (Math.random() * 4) : 0.2 + (Math.random() * 0.8);
            }

            // Ensure speed doesn't go below 0.1
            newSpeed = Math.max(0.1, newSpeed);

            // Determine unit (MB/s or KB/s)
            let displaySpeed = '';
            if (newSpeed < 1.0) {
                displaySpeed = `${Math.floor(newSpeed * 1000)} KB/s`;
            } else {
                displaySpeed = `${newSpeed.toFixed(1)} MB/s`;
            }

            // Add a slight opacity fade to simulate activity blink
            speedSpan.style.opacity = '0.5';
            setTimeout(() => {
                speedSpan.textContent = displaySpeed;
                speedSpan.style.opacity = '1';
            }, 100);
            
            // Slowly drift the base speed over time
            if (Math.random() > 0.7) {
                baseSpeed += (Math.random() * 0.4 - 0.2);
                baseSpeed = Math.max(1.5, Math.min(6.5, baseSpeed));
            }
        }
    }, 1500);
});

window.stickyNotesWidgetRef = null;
window.registerStickyNotesWidget = function(ref) { window.stickyNotesWidgetRef = ref; console.log('Sticky Notes Widget Registered!'); };
window.triggerAddNewNote = function() { if (window.stickyNotesWidgetRef) { window.stickyNotesWidgetRef.invokeMethodAsync('AddNewNoteJS').catch(err => console.error('Error adding note:', err)); } else { fetch('/api/stickynotes', { method: 'POST' }).then(r => { if (r.ok) location.reload(); else alert('Failed to create note.'); }).catch(e => alert('Error creating note.')); } };

// STICKY NOTES VANILLA JS LOGIC
let activeNoteId = null;
let dragStartX = 0;
let dragStartY = 0;
let initialNoteX = 0;
let initialNoteY = 0;

window.beginDragJS = function(e, id) {
    if (e.target.tagName.toLowerCase() === 'select' || e.target.tagName.toLowerCase() === 'button' || e.target.closest('button')) return;
    activeNoteId = id;
    dragStartX = e.clientX;
    dragStartY = e.clientY;
    
    let noteEl = document.getElementById('note-' + id);
    initialNoteX = parseInt(noteEl.style.left) || 0;
    initialNoteY = parseInt(noteEl.style.top) || 0;
    
    document.addEventListener('mousemove', dragJS);
    document.addEventListener('mouseup', endDragJS);
};

function dragJS(e) {
    if (!activeNoteId) return;
    let noteEl = document.getElementById('note-' + activeNoteId);
    if (!noteEl) return;
    
    let deltaX = e.clientX - dragStartX;
    let deltaY = e.clientY - dragStartY;
    
    noteEl.style.left = (initialNoteX + deltaX) + 'px';
    noteEl.style.top = (initialNoteY + deltaY) + 'px';
}

function endDragJS(e) {
    if (activeNoteId) {
        window.saveNoteJS(activeNoteId);
        activeNoteId = null;
    }
    document.removeEventListener('mousemove', dragJS);
    document.removeEventListener('mouseup', endDragJS);
}

window.toggleMinimizeJS = function(id) {
    let content = document.getElementById('note-content-' + id);
    let noteEl = document.getElementById('note-' + id);
    let icon = document.getElementById('minimize-icon-' + id);
    if (content.classList.contains('hidden')) {
        content.classList.remove('hidden');
        noteEl.classList.remove('minimized');
        icon.classList.remove('bi-arrows-expand');
        icon.classList.add('bi-dash');
    } else {
        content.classList.add('hidden');
        noteEl.classList.add('minimized');
        icon.classList.remove('bi-dash');
        icon.classList.add('bi-arrows-expand');
    }
};

window.deleteNoteJS = function(id) {
    fetch('/api/stickynotes/' + id, { method: 'DELETE' }).then(r => {
        if (r.ok) {
            let el = document.getElementById('note-' + id);
            if (el) el.remove();
        }
    });
};

window.saveNoteJS = function(id) {
    let noteEl = document.getElementById('note-' + id);
    if (!noteEl) return;
    
    let data = {
        Id: id,
        Content: document.getElementById('note-text-' + id).value || '',
        Color: document.getElementById('note-color-' + id).value || 'yellow',
        XPos: parseInt(noteEl.style.left) || 0,
        YPos: parseInt(noteEl.style.top) || 0,
        ReminderTime: document.getElementById('note-reminder-' + id).value || null
    };
    
    // Convert background color immediately
    const colors = { 'yellow': '#fff9c4', 'blue': '#bbdefb', 'pink': '#f8bbd0', 'green': '#c8e6c9' };
    noteEl.style.backgroundColor = colors[data.Color];
    
    fetch('/api/stickynotes/' + id, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
    });
};


window.setBackgroundColor = (color) => {
    document.documentElement.style.setProperty('--theme-bg', color);
    localStorage.setItem('customBgColor', color);
};

window.setCaptionColor = (color) => {
    document.documentElement.style.setProperty('--caption-color', color);
    localStorage.setItem('customCaptionColor', color);
};

window.setHeadingBgColor = (color) => {
    document.documentElement.style.setProperty('--heading-bg', color);
    localStorage.setItem('customHeadingBgColor', color);
};

window.setHeadingTextColor = (color) => {
    document.documentElement.style.setProperty('--heading-color', color);
    localStorage.setItem('customHeadingTextColor', color);
};


window.fetchWhatsAppQR = async () => {
    const statusText = document.getElementById('whatsapp-status-text');
    const qrImage = document.getElementById('whatsapp-qr-image');
    const loader = document.getElementById('whatsapp-qr-loader');
    const dropdown = document.getElementById('whatsappDropdown');
    
    if (!statusText || !dropdown) return;
    
    statusText.innerText = 'Connecting to service...';
    qrImage.style.display = 'none';
    loader.style.display = 'inline-block';
    
    try {
        const response = await fetch('http://localhost:3001/qr');
        const data = await response.json();
        
        loader.style.display = 'none';
        
        if (data.status === 'CONNECTED') {
            statusText.innerText = 'Authenticated & Connected!';
            statusText.className = 'mb-2 text-success fw-bold';
            dropdown.innerText = 'WhatsApp Connected';
            dropdown.className = 'top-badge top-badge-success dropdown-toggle';
        } else if (data.qr) {
            statusText.innerText = 'Scan QR Code with WhatsApp';
            statusText.className = 'mb-2 text-dark fw-bold';
            qrImage.src = data.qr;
            qrImage.style.display = 'block';
            dropdown.innerText = 'WhatsApp Pending Scan';
            dropdown.className = 'top-badge top-badge-warning dropdown-toggle';
        } else {
            statusText.innerText = 'Status: ' + data.status;
            statusText.className = 'mb-2 text-muted fw-bold';
        }
    } catch (e) {
        loader.style.display = 'none';
        statusText.innerText = 'Service offline or unreachable';
        statusText.className = 'mb-2 text-danger fw-bold';
    }
};


document.addEventListener('DOMContentLoaded', () => {
    const whatsappDropdownElement = document.getElementById('whatsappDropdown')?.parentElement;
    if (whatsappDropdownElement) {
        whatsappDropdownElement.addEventListener('show.bs.dropdown', () => {
            if (window.fetchWhatsAppQR) {
                window.fetchWhatsAppQR();
            }
        });
    }
});


document.addEventListener('show.bs.dropdown', function (event) {
    if (event.target && event.target.id === 'whatsappDropdown') {
        if (typeof window.fetchWhatsAppQR === 'function') {
            window.fetchWhatsAppQR();
        }
    }
});

// AI Chatbot Voice & Audio Helpers
let currentSpeechRecognition = null;

window.speakText = (text) => {
    try {
        if (!('speechSynthesis' in window)) return;
        window.speechSynthesis.cancel(); // stop any ongoing speech
        if (!text) return;
        
        // Strip markdown characters for natural speech readout
        const cleanText = text
            .replace(/```[\s\S]*?```/g, 'Code snippet omitted.')
            .replace(/`([^`]+)`/g, '$1')
            .replace(/[*#_~>]/g, '')
            .replace(/\[([^\]]+)\]\([^)]+\)/g, '$1')
            .trim();

        const utterance = new SpeechSynthesisUtterance(cleanText);
        utterance.rate = 1.05;
        utterance.pitch = 1.0;
        
        // Select an English voice if available
        const voices = window.speechSynthesis.getVoices();
        const preferredVoice = voices.find(v => v.lang.startsWith('en') && (v.name.includes('Natural') || v.name.includes('Google') || v.name.includes('David') || v.name.includes('Zira') || v.name.includes('Samantha')));
        if (preferredVoice) {
            utterance.voice = preferredVoice;
        }

        window.speechSynthesis.speak(utterance);
    } catch (e) {
        console.warn('Speech synthesis error:', e);
    }
};

window.stopSpeaking = () => {
    try {
        if ('speechSynthesis' in window) {
            window.speechSynthesis.cancel();
        }
    } catch (e) {}
};

window.startVoiceDictation = (dotNetHelper, callbackMethod) => {
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) {
        alert('Voice dictation is not supported in this browser. Please use Chrome, Edge, or Safari.');
        return;
    }

    if (currentSpeechRecognition) {
        try { currentSpeechRecognition.stop(); } catch (e) {}
        currentSpeechRecognition = null;
    }

    try {
        const recognition = new SpeechRecognition();
        recognition.lang = 'en-US';
        recognition.interimResults = false;
        recognition.maxAlternatives = 1;

        recognition.onresult = (event) => {
            const transcript = event.results[0][0].transcript;
            if (transcript && dotNetHelper) {
                dotNetHelper.invokeMethodAsync(callbackMethod, transcript);
            }
        };

        recognition.onerror = (event) => {
            console.warn('Voice recognition error:', event.error);
            if (dotNetHelper) {
                dotNetHelper.invokeMethodAsync('OnVoiceError', event.error);
            }
        };

        recognition.onend = () => {
            if (dotNetHelper) {
                dotNetHelper.invokeMethodAsync('OnVoiceEnd');
            }
            currentSpeechRecognition = null;
        };

        currentSpeechRecognition = recognition;
        recognition.start();
    } catch (err) {
        console.error('Failed to start speech recognition:', err);
    }
};

window.stopVoiceDictation = () => {
    if (currentSpeechRecognition) {
        try {
            currentSpeechRecognition.stop();
        } catch (e) {}
        currentSpeechRecognition = null;
    }
};

window.copyToClipboard = async (text) => {
    try {
        await navigator.clipboard.writeText(text);
        return true;
    } catch (e) {
        // Fallback
        const textarea = document.createElement('textarea');
        textarea.value = text;
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        document.body.removeChild(textarea);
        return true;
    }
};

window.scrollChatToBottom = (elementId) => {
    const el = document.getElementById(elementId);
    if (el) {
        setTimeout(() => {
            el.scrollTop = el.scrollHeight;
        }, 50);
    }
};

// ==========================================
// MOVEABLE FLOATING WIDGETS (Sticky Notes & Ask AI)
// ==========================================

window.initStickyNoteButtonDrag = function() {
    const btn = document.getElementById('btn-add-sticky-note');
    if (!btn) return;

    const storageKey = 'sticky_note_btn_position';

    function restorePos() {
        let x, y;
        const saved = localStorage.getItem(storageKey);
        const rect = btn.getBoundingClientRect();
        const width = rect.width || 148;
        const height = rect.height || 36;

        if (saved) {
            try {
                const parsed = JSON.parse(saved);
                x = parsed.x;
                y = parsed.y;
            } catch (e) {}
        }

        if (x === undefined || y === undefined) {
            x = window.innerWidth - width - 40;
            y = window.innerHeight - height - 90;
        }

        x = Math.max(10, Math.min(window.innerWidth - width - 10, x));
        y = Math.max(10, Math.min(window.innerHeight - height - 10, y));

        btn.style.left = x + 'px';
        btn.style.top = y + 'px';
        btn.style.right = 'auto';
        btn.style.bottom = 'auto';
        btn.style.position = 'fixed';
    }
    btn._restorePos = restorePos;

    requestAnimationFrame(restorePos);

    if (btn._dragInitialized) {
        return;
    }
    btn._dragInitialized = true;

    let isDragging = false;
    let startX = 0;
    let startY = 0;
    let initialX = 0;
    let initialY = 0;

    function onStart(e) {
        if (e.type === 'mousedown' && e.button !== 0) return;
        const clientX = e.type.startsWith('touch') ? e.touches[0].clientX : e.clientX;
        const clientY = e.type.startsWith('touch') ? e.touches[0].clientY : e.clientY;

        startX = clientX;
        startY = clientY;

        const rect = btn.getBoundingClientRect();
        initialX = rect.left;
        initialY = rect.top;
        isDragging = false;

        document.addEventListener('mousemove', onMove, { passive: false });
        document.addEventListener('mouseup', onEnd);
        document.addEventListener('touchmove', onMove, { passive: false });
        document.addEventListener('touchend', onEnd);
    }

    function onMove(e) {
        const clientX = e.type.startsWith('touch') ? e.touches[0].clientX : e.clientX;
        const clientY = e.type.startsWith('touch') ? e.touches[0].clientY : e.clientY;

        const dx = clientX - startX;
        const dy = clientY - startY;

        if (!isDragging) {
            if (Math.abs(dx) >= 4 || Math.abs(dy) >= 4) {
                isDragging = true;
                btn.classList.add('is-dragging');
                document.body.style.userSelect = 'none';
            }
        }

        if (isDragging) {
            if (e.cancelable) e.preventDefault();
            const rect = btn.getBoundingClientRect();
            let nx = initialX + dx;
            let ny = initialY + dy;

            nx = Math.max(10, Math.min(window.innerWidth - rect.width - 10, nx));
            ny = Math.max(10, Math.min(window.innerHeight - rect.height - 10, ny));

            btn.style.left = nx + 'px';
            btn.style.top = ny + 'px';
        }
    }

    function onEnd() {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onEnd);
        document.removeEventListener('touchmove', onMove);
        document.removeEventListener('touchend', onEnd);
        document.body.style.userSelect = '';

        if (isDragging) {
            btn.classList.remove('is-dragging');
            const rect = btn.getBoundingClientRect();
            localStorage.setItem(storageKey, JSON.stringify({ x: rect.left, y: rect.top }));

            // Intercept and swallow click event so a new note is NOT created upon dragging
            const captureClick = (ev) => {
                ev.stopImmediatePropagation();
                ev.preventDefault();
                window.removeEventListener('click', captureClick, true);
            };
            window.addEventListener('click', captureClick, true);
            setTimeout(() => window.removeEventListener('click', captureClick, true), 120);
        }
        isDragging = false;
    }

    btn.addEventListener('mousedown', onStart);
    btn.addEventListener('touchstart', onStart, { passive: true });
};

// Ask AI launcher and window positioning
window.initAiChatbotLauncherPosition = function() {
    const root = document.getElementById('ai-chatbot-root');
    const launcher = document.getElementById('ai-chatbot-launcher');
    if (!root || !launcher) return;

    const storageKey = 'ask_ai_icon_pos';

    function restorePos() {
        let x, y;
        const saved = localStorage.getItem(storageKey);
        const rect = launcher.getBoundingClientRect();
        const width = rect.width || 120;
        const height = rect.height || 44;

        if (saved) {
            try {
                const parsed = JSON.parse(saved);
                x = parsed.x;
                y = parsed.y;
            } catch (e) {}
        }

        if (x === undefined || y === undefined) {
            x = window.innerWidth - width - 24;
            y = window.innerHeight - height - 24;
        }

        x = Math.max(10, Math.min(window.innerWidth - width - 10, x));
        y = Math.max(10, Math.min(window.innerHeight - height - 10, y));

        root.style.left = x + 'px';
        root.style.top = y + 'px';
        root.style.right = 'auto';
        root.style.bottom = 'auto';
        root.style.position = 'fixed';
    }

    requestAnimationFrame(restorePos);

    if (launcher._dragInitialized) {
        return;
    }
    launcher._dragInitialized = true;
    launcher._restorePos = restorePos;

    let isDragging = false;
    let startX = 0;
    let startY = 0;
    let initialX = 0;
    let initialY = 0;

    function onStart(e) {
        if (e.type === 'mousedown' && e.button !== 0) return;
        const clientX = e.type.startsWith('touch') ? e.touches[0].clientX : e.clientX;
        const clientY = e.type.startsWith('touch') ? e.touches[0].clientY : e.clientY;

        startX = clientX;
        startY = clientY;

        const rect = root.getBoundingClientRect();
        initialX = rect.left;
        initialY = rect.top;
        isDragging = false;

        document.addEventListener('mousemove', onMove, { passive: false });
        document.addEventListener('mouseup', onEnd);
        document.addEventListener('touchmove', onMove, { passive: false });
        document.addEventListener('touchend', onEnd);
    }

    function onMove(e) {
        const clientX = e.type.startsWith('touch') ? e.touches[0].clientX : e.clientX;
        const clientY = e.type.startsWith('touch') ? e.touches[0].clientY : e.clientY;

        const dx = clientX - startX;
        const dy = clientY - startY;

        if (!isDragging) {
            if (Math.abs(dx) >= 4 || Math.abs(dy) >= 4) {
                isDragging = true;
                launcher.classList.add('is-dragging');
                document.body.style.userSelect = 'none';
            }
        }

        if (isDragging) {
            if (e.cancelable) e.preventDefault();
            const rect = launcher.getBoundingClientRect();
            let nx = initialX + dx;
            let ny = initialY + dy;

            nx = Math.max(10, Math.min(window.innerWidth - rect.width - 10, nx));
            ny = Math.max(10, Math.min(window.innerHeight - rect.height - 10, ny));

            root.style.left = nx + 'px';
            root.style.top = ny + 'px';
        }
    }

    function onEnd() {
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onEnd);
        document.removeEventListener('touchmove', onMove);
        document.removeEventListener('touchend', onEnd);
        document.body.style.userSelect = '';

        if (isDragging) {
            launcher.classList.remove('is-dragging');
            const rect = root.getBoundingClientRect();
            localStorage.setItem(storageKey, JSON.stringify({ x: rect.left, y: rect.top }));

            // Intercept and swallow click event so chat does NOT open when user was dragging
            const captureClick = (ev) => {
                ev.stopImmediatePropagation();
                ev.preventDefault();
                window.removeEventListener('click', captureClick, true);
            };
            window.addEventListener('click', captureClick, true);
            setTimeout(() => window.removeEventListener('click', captureClick, true), 120);
        }
        isDragging = false;
    }

    launcher.addEventListener('mousedown', onStart);
    launcher.addEventListener('touchstart', onStart, { passive: true });
};

window.initAiChatbotWindowPosition = function() {
    const root = document.getElementById('ai-chatbot-root');
    const win = document.getElementById('ai-chatbot-window');
    if (!root || !win) return;

    let saved = null;
    try {
        saved = JSON.parse(localStorage.getItem('ask_ai_icon_pos'));
    } catch (e) {}

    const winWidth = Math.min(410, window.innerWidth - 32);
    const winHeight = Math.min(590, window.innerHeight - 40);

    let left, top;
    if (saved) {
        if (saved.x > window.innerWidth / 2) {
            left = saved.x + 130 - winWidth;
        } else {
            left = saved.x;
        }

        if (saved.y > window.innerHeight / 2) {
            top = saved.y + 44 - winHeight;
        } else {
            top = saved.y;
        }
    } else {
        left = window.innerWidth - winWidth - 24;
        top = window.innerHeight - winHeight - 24;
    }

    left = Math.max(16, Math.min(window.innerWidth - winWidth - 16, left));
    top = Math.max(16, Math.min(window.innerHeight - winHeight - 16, top));

    root.style.left = left + 'px';
    root.style.top = top + 'px';
    root.style.right = 'auto';
    root.style.bottom = 'auto';
    root.style.position = 'fixed';

    // Also make the window header draggable so the user can reposition the open window!
    const header = document.getElementById('ai-chat-header');
    if (!header || header._dragInitialized) return;
    header._dragInitialized = true;

    let isWinDragging = false;
    let wStartX = 0, wStartY = 0, wInitialX = 0, wInitialY = 0;

    function onHStart(e) {
        if (e.type === 'mousedown' && e.button !== 0) return;
        if (e.target.closest('button, input, select, textarea')) return;

        const clientX = e.type.startsWith('touch') ? e.touches[0].clientX : e.clientX;
        const clientY = e.type.startsWith('touch') ? e.touches[0].clientY : e.clientY;

        wStartX = clientX;
        wStartY = clientY;

        const rect = root.getBoundingClientRect();
        wInitialX = rect.left;
        wInitialY = rect.top;
        isWinDragging = false;

        document.addEventListener('mousemove', onHMove, { passive: false });
        document.addEventListener('mouseup', onHEnd);
        document.addEventListener('touchmove', onHMove, { passive: false });
        document.addEventListener('touchend', onHEnd);
    }

    function onHMove(e) {
        const clientX = e.type.startsWith('touch') ? e.touches[0].clientX : e.clientX;
        const clientY = e.type.startsWith('touch') ? e.touches[0].clientY : e.clientY;

        const dx = clientX - wStartX;
        const dy = clientY - wStartY;

        if (!isWinDragging) {
            if (Math.abs(dx) >= 4 || Math.abs(dy) >= 4) {
                isWinDragging = true;
                header.style.cursor = 'grabbing';
                document.body.style.userSelect = 'none';
            }
        }

        if (isWinDragging) {
            if (e.cancelable) e.preventDefault();
            const rect = win.getBoundingClientRect();
            let nx = wInitialX + dx;
            let ny = wInitialY + dy;

            nx = Math.max(10, Math.min(window.innerWidth - rect.width - 10, nx));
            ny = Math.max(10, Math.min(window.innerHeight - rect.height - 10, ny));

            root.style.left = nx + 'px';
            root.style.top = ny + 'px';
        }
    }

    function onHEnd() {
        document.removeEventListener('mousemove', onHMove);
        document.removeEventListener('mouseup', onHEnd);
        document.removeEventListener('touchmove', onHMove);
        document.removeEventListener('touchend', onHEnd);
        document.body.style.userSelect = '';
        if (isWinDragging) {
            header.style.cursor = 'grab';
        }
        isWinDragging = false;
    }

    header.addEventListener('mousedown', onHStart);
    header.addEventListener('touchstart', onHStart, { passive: true });
};

// Window resize safety listener
window.addEventListener('resize', () => {
    const btn = document.getElementById('btn-add-sticky-note');
    if (btn) {
        const rect = btn.getBoundingClientRect();
        if (rect.right > window.innerWidth || rect.bottom > window.innerHeight) {
            let x = Math.max(10, Math.min(window.innerWidth - rect.width - 10, rect.left));
            let y = Math.max(10, Math.min(window.innerHeight - rect.height - 10, rect.top));
            btn.style.left = x + 'px';
            btn.style.top = y + 'px';
            localStorage.setItem('sticky_note_btn_position', JSON.stringify({ x, y }));
        }
    }
    const aiRoot = document.getElementById('ai-chatbot-root');
    if (aiRoot) {
        const rect = aiRoot.getBoundingClientRect();
        if (rect.right > window.innerWidth || rect.bottom > window.innerHeight) {
            let x = Math.max(10, Math.min(window.innerWidth - rect.width - 10, rect.left));
            let y = Math.max(10, Math.min(window.innerHeight - rect.height - 10, rect.top));
            aiRoot.style.left = x + 'px';
            aiRoot.style.top = y + 'px';
            localStorage.setItem('ask_ai_icon_pos', JSON.stringify({ x, y }));
        }
    }
});



