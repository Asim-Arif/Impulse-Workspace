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
                    return;
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
