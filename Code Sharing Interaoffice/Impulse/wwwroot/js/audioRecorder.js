window.audioRecorder = {
    recorder: null,
    audioChunks: [],

    startRecording: async function () {
        try {
            if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
                alert("Microphone access is not supported in this browser or you are not using HTTPS. Please ensure you access the site via HTTPS.");
                return false;
            }
            
            const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
            this.audioChunks = [];
            
            let options = { mimeType: 'audio/webm;codecs=opus' };
            if (!MediaRecorder.isTypeSupported(options.mimeType)) {
                options = { mimeType: 'audio/webm' };
                if (!MediaRecorder.isTypeSupported(options.mimeType)) {
                    options = { mimeType: 'audio/mp4' }; 
                    if (!MediaRecorder.isTypeSupported(options.mimeType)) {
                        options = {}; // use default
                    }
                }
            }
            
            this.recorder = new MediaRecorder(stream, options);
            
            this.recorder.ondataavailable = (event) => {
                if (event.data && event.data.size > 0) {
                    this.audioChunks.push(event.data);
                }
            };
            
            this.recorder.start(250); // Provide 250ms timeslice
            return true;
        } catch (error) {
            console.error("Error starting audio recording:", error);
            return false;
        }
    },

    stopRecording: function () {
        return new Promise((resolve) => {
            try {
                if (!this.recorder || this.recorder.state === "inactive") {
                    resolve(null);
                }

                // Fallback timeout in case onstop never fires
                const timeoutId = setTimeout(() => {
                    resolve(null);
                }, 3000);

                this.recorder.onstop = () => {
                    clearTimeout(timeoutId);
                    
                    if (this.recorder.stream) {
                        this.recorder.stream.getTracks().forEach(track => track.stop());
                    }
                    
                    if (this.audioChunks.length === 0) {
                        resolve(null);
                        return;
                    }

                    const audioBlob = new Blob(this.audioChunks, { type: this.recorder.mimeType || 'audio/webm' });
                    const reader = new FileReader();
                    reader.readAsDataURL(audioBlob);
                    reader.onloadend = () => {
                        if (reader.result) {
                            const base64data = reader.result.toString();
                            if (base64data.includes(',')) {
                                const base64String = base64data.split(',')[1];
                                const mimeType = audioBlob.type || 'audio/webm';
                                resolve({ base64: base64String, mimeType: mimeType });
                                return;
                            }
                        }
                        resolve(null);
                    };
                    reader.onerror = () => resolve(null);
                };
                
                this.recorder.stop();
            } catch (err) {
                console.error("Stop recording error:", err);
                resolve(null);
            }
        });
    }
};

window.meetingSpeech = {
    recognition: null,
    isListening: false,
    dotnetHelper: null,

    isSupported: function () {
        return !!(window.SpeechRecognition || window.webkitSpeechRecognition);
    },

    startListening: function (dotnetHelper) {
        if (!this.isSupported()) {
            alert("Speech recognition is not supported in this browser. Please use Google Chrome or Microsoft Edge.");
            return false;
        }

        this.dotnetHelper = dotnetHelper;
        const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;

        if (!this.recognition) {
            this.recognition = new SpeechRecognition();
            this.recognition.continuous = true;
            this.recognition.interimResults = true;
            this.recognition.lang = 'en-US';

            this.recognition.onresult = (event) => {
                let interimTranscript = '';
                let finalTranscript = '';

                for (let i = event.resultIndex; i < event.results.length; ++i) {
                    const transcriptPiece = event.results[i][0].transcript;
                    if (event.results[i].isFinal) {
                        finalTranscript += transcriptPiece + ' ';
                    } else {
                        interimTranscript += transcriptPiece;
                    }
                }

                if (this.dotnetHelper) {
                    this.dotnetHelper.invokeMethodAsync('OnSpeechRecognized', finalTranscript, interimTranscript);
                }
            };

            this.recognition.onerror = (event) => {
                console.warn("Speech recognition error:", event.error);
                if (event.error === 'not-allowed') {
                    alert("Microphone access was denied. Please allow microphone permissions to use voice dictation.");
                }
            };

            this.recognition.onend = () => {
                if (this.isListening) {
                    // Auto-restart if user did not manually stop
                    try {
                        this.recognition.start();
                    } catch (e) {
                        console.log("Recognition auto-restart suppressed:", e);
                    }
                }
            };
        }

        try {
            this.isListening = true;
            this.recognition.start();
            return true;
        } catch (e) {
            console.error("Error starting speech recognition:", e);
            return false;
        }
    },

    stopListening: function () {
        this.isListening = false;
        if (this.recognition) {
            try {
                this.recognition.stop();
            } catch (e) {
                console.error("Error stopping speech recognition:", e);
            }
        }
        return true;
    }
};
