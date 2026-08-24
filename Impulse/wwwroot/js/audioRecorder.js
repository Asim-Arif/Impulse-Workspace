window.audioRecorder = {
    recorder: null,
    stream: null,
    audioChunks: [],

    startRecording: async function () {
        try {
            if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
                console.error("getUserMedia not supported or not HTTPS");
                alert("Microphone access is not supported in this browser or you are not using HTTPS. Please ensure you access the site via HTTPS.");
                return false;
            }
            
            this.stream = await navigator.mediaDevices.getUserMedia({ audio: true });
            this.audioChunks = [];
            
            let mimeType = 'audio/webm;codecs=opus';
            if (!MediaRecorder.isTypeSupported(mimeType)) {
                mimeType = 'audio/webm';
                if (!MediaRecorder.isTypeSupported(mimeType)) {
                    mimeType = 'audio/mp4'; 
                    if (!MediaRecorder.isTypeSupported(mimeType)) {
                        mimeType = ''; // browser default
                    }
                }
            }
            
            const options = mimeType ? { mimeType: mimeType } : {};
            this.recorder = new MediaRecorder(this.stream, options);
            
            this.recorder.ondataavailable = (event) => {
                if (event.data && event.data.size > 0) {
                    this.audioChunks.push(event.data);
                }
            };
            
            this.recorder.start(100); // 100ms chunks
            console.log("Audio recording started with mimeType:", this.recorder.mimeType);
            return true;
        } catch (error) {
            console.error("Error starting audio recording:", error);
            alert("Error accessing microphone: " + error.message);
            return false;
        }
    },

    stopRecording: function () {
        return new Promise((resolve) => {
            try {
                if (!this.recorder) {
                    console.warn("No recorder instance found");
                    resolve(null);
                    return;
                }

                if (this.recorder.state === "recording") {
                    try {
                        this.recorder.requestData();
                    } catch (e) { }
                }

                const actualMime = this.recorder.mimeType || 'audio/webm';

                this.recorder.onstop = () => {
                    try {
                        if (this.stream) {
                            this.stream.getTracks().forEach(track => track.stop());
                        }

                        if (this.audioChunks.length === 0) {
                            console.warn("No audio chunks recorded");
                            resolve(null);
                            return;
                        }

                        const audioBlob = new Blob(this.audioChunks, { type: actualMime });
                        const reader = new FileReader();
                        reader.onloadend = () => {
                            if (reader.result) {
                                const dataUrl = reader.result.toString();
                                const base64String = dataUrl.includes(',') ? dataUrl.split(',')[1] : dataUrl;
                                console.log("Voice note recorded successfully! Size: " + audioBlob.size + " bytes");
                                resolve({ base64: base64String, mimeType: actualMime });
                            } else {
                                resolve(null);
                            }
                        };
                        reader.onerror = (err) => {
                            console.error("FileReader error:", err);
                            resolve(null);
                        };
                        reader.readAsDataURL(audioBlob);
                    } catch (innerErr) {
                        console.error("Error in onstop handler:", innerErr);
                        resolve(null);
                    }
                };

                if (this.recorder.state !== "inactive") {
                    this.recorder.stop();
                } else {
                    this.recorder.onstop();
                }
            } catch (err) {
                console.error("Stop recording error:", err);
                resolve(null);
            }
        });
    }
};
