window.openPdfFromBase64 = (base64) => {
    try {
        console.log("PDF received, converting base64...");
        const byteCharacters = atob(base64);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        const blob = new Blob([byteArray], { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);

        const newWin = window.open(url, '_blank');
        if (!newWin || newWin.closed || typeof newWin.closed === 'undefined') {
            console.warn("Popup blocked by browser. Triggering download fallback...");
            const anchor = document.createElement('a');
            anchor.href = url;
            anchor.target = '_blank';
            anchor.download = 'Report.pdf';
            document.body.appendChild(anchor);
            anchor.click();
            anchor.remove();
        } else {
            console.log("PDF opened successfully in new tab.");
        }
    } catch (err) {
        console.error("Error in openPdfFromBase64:", err);
    }
};

window.downloadFileFromBytes = (byteArray, contentType, fileName) => {
    try {
        const blob = new Blob([new Uint8Array(byteArray)], { type: contentType });
        const url = window.URL.createObjectURL(blob);
        const anchorElement = document.createElement('a');
        anchorElement.href = url;
        anchorElement.download = fileName ?? 'Report';
        document.body.appendChild(anchorElement);
        anchorElement.click();
        anchorElement.remove();
        window.URL.revokeObjectURL(url);
    } catch (err) {
        console.error("Error in downloadFileFromBytes:", err);
    }
};