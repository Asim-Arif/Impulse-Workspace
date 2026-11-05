// Impulse ERP - Report Viewer & PDF Helper with Excel, Word & Image Export
window.openReportViewer = (base64, request) => {
    try {
        console.log("PDF received, preparing Report Viewer with Excel, Word & Image export...");
        const byteCharacters = atob(base64);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
            byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        const pdfBlob = new Blob([byteArray], { type: 'application/pdf' });
        const pdfUrl = URL.createObjectURL(pdfBlob);

        // Fallback to simple viewer if request metadata is missing
        if (!request) {
            window.openPdfFromBase64(base64);
            return;
        }

        const reportName = request.ReportName || request.reportName || "Report.rpt";
        const cleanTitle = reportName.replace(/\.rpt$/i, '');
        const origin = window.location.origin;
        const serializedRequest = JSON.stringify(request).replace(/</g, '\\u003c');

        const viewerHtml = `<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${cleanTitle} - Impulse ERP Preview</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.2/css/all.min.css" rel="stylesheet">
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"><\/script>
    <style>
        * { box-sizing: border-box; }
        html, body {
            margin: 0;
            padding: 0;
            width: 100%;
            height: 100%;
            overflow: hidden;
            background: #1e293b;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        .viewer-toolbar {
            height: 48px;
            background: #1e293b;
            color: #ffffff;
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 0 16px;
            border-bottom: 1px solid #334155;
            box-shadow: 0 2px 8px rgba(0,0,0,0.25);
            z-index: 100;
        }
        .report-title-badge {
            font-size: 0.9rem;
            font-weight: 600;
            color: #f8fafc;
            display: flex;
            align-items: center;
            gap: 8px;
            max-width: 380px;
            overflow: hidden;
            text-overflow: ellipsis;
            white-space: nowrap;
        }
        .btn-export-excel {
            background-color: #107c41;
            color: #ffffff;
            border: none;
            padding: 6px 13px;
            border-radius: 6px;
            font-size: 0.82rem;
            font-weight: 600;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            cursor: pointer;
            transition: all 0.2s ease;
        }
        .btn-export-excel:hover:not(:disabled) {
            background-color: #0e6b37;
            transform: translateY(-1px);
            color: #ffffff;
        }
        .btn-export-excel:disabled {
            opacity: 0.7;
            cursor: wait;
        }
        .btn-export-word {
            background-color: #185abd;
            color: #ffffff;
            border: none;
            padding: 6px 13px;
            border-radius: 6px;
            font-size: 0.82rem;
            font-weight: 600;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            cursor: pointer;
            transition: all 0.2s ease;
        }
        .btn-export-word:hover:not(:disabled) {
            background-color: #124796;
            transform: translateY(-1px);
            color: #ffffff;
        }
        .btn-export-word:disabled {
            opacity: 0.7;
            cursor: wait;
        }
        .btn-export-image {
            background-color: #d97706;
            color: #ffffff;
            border: none;
            padding: 6px 13px;
            border-radius: 6px;
            font-size: 0.82rem;
            font-weight: 600;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            cursor: pointer;
            transition: all 0.2s ease;
        }
        .btn-export-image:hover:not(:disabled) {
            background-color: #b45309;
            transform: translateY(-1px);
            color: #ffffff;
        }
        .btn-export-image:disabled {
            opacity: 0.7;
            cursor: wait;
        }
        .btn-action-tool {
            background: rgba(255, 255, 255, 0.08);
            color: #e2e8f0;
            border: 1px solid rgba(255, 255, 255, 0.2);
            padding: 6px 11px;
            border-radius: 6px;
            font-size: 0.82rem;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            cursor: pointer;
            transition: all 0.2s ease;
        }
        .btn-action-tool:hover {
            background: rgba(255, 255, 255, 0.16);
            color: #ffffff;
        }
        .pdf-frame-container {
            width: 100%;
            height: calc(100% - 48px);
            background: #525659;
        }
        .pdf-frame {
            width: 100%;
            height: 100%;
            border: none;
            display: block;
        }
        .toast-banner {
            position: fixed;
            top: 56px;
            right: 20px;
            z-index: 1000;
            padding: 10px 18px;
            border-radius: 8px;
            font-size: 0.85rem;
            box-shadow: 0 8px 24px rgba(0,0,0,0.3);
            display: none;
            align-items: center;
            gap: 8px;
            font-weight: 500;
        }
        .toast-success { background: #059669; color: white; }
        .toast-error { background: #dc2626; color: white; }
    </style>
</head>
<body>
    <div class="viewer-toolbar">
        <div class="d-flex align-items-center gap-2">
            <i class="fas fa-file-pdf text-danger fs-5"></i>
            <span class="report-title-badge" title="${reportName}">${reportName}</span>
            <span class="badge bg-secondary d-none d-lg-inline-block" style="font-size: 0.68rem; letter-spacing: 0.5px;">PDF PREVIEW</span>
        </div>

        <div class="d-flex align-items-center gap-2">
            <button id="btnExcel" class="btn-export-excel" onclick="triggerExport('excel')" title="Export to Microsoft Excel (.xlsx)">
                <i class="fas fa-file-excel"></i>
                <span>Export to Excel</span>
            </button>
            <button id="btnWord" class="btn-export-word" onclick="triggerExport('word')" title="Export to Microsoft Word (.docx)">
                <i class="fas fa-file-word"></i>
                <span>Export to Word</span>
            </button>
            <button id="btnImage" class="btn-export-image" onclick="exportToImage()" title="Export to High-Resolution PNG Image">
                <i class="fas fa-file-image"></i>
                <span>Export to Image</span>
            </button>
            <button id="btnDownloadPdf" class="btn-action-tool d-none d-md-inline-flex" onclick="downloadPdf()" title="Download PDF">
                <i class="fas fa-download"></i>
                <span>PDF</span>
            </button>
            <button id="btnPrintPdf" class="btn-action-tool d-none d-md-inline-flex" onclick="printPdf()" title="Print Report">
                <i class="fas fa-print"></i>
                <span>Print</span>
            </button>
            <button class="btn-action-tool text-white-50" onclick="window.close()" title="Close Viewer">
                <i class="fas fa-times"></i>
            </button>
        </div>
    </div>

    <div class="pdf-frame-container">
        <iframe id="pdfFrame" class="pdf-frame" src="${pdfUrl}"></iframe>
    </div>

    <div id="toastBanner" class="toast-banner"></div>

    <script>
        const reportPayload = ${serializedRequest};
        const pdfBlobUrl = "${pdfUrl}";
        const baseFileName = "${cleanTitle}";
        const apiOrigin = "${origin}";

        if (window.pdfjsLib) {
            pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        }

        function showToast(msg, isError) {
            const toast = document.getElementById('toastBanner');
            toast.className = 'toast-banner ' + (isError ? 'toast-error' : 'toast-success');
            toast.innerHTML = (isError ? '<i class=\"fas fa-triangle-exclamation\"></i>' : '<i class=\"fas fa-check-circle\"></i>') + ' <span>' + msg + '</span>';
            toast.style.display = 'flex';
            setTimeout(() => { toast.style.display = 'none'; }, 4000);
        }

        async function triggerExport(format) {
            const isExcel = (format === 'excel' || format === 'xlsx');
            const btn = document.getElementById(isExcel ? 'btnExcel' : 'btnWord');
            const originalHtml = btn.innerHTML;
            btn.disabled = true;
            btn.innerHTML = '<i class=\"fas fa-spinner fa-spin\"></i> <span>Generating ' + (isExcel ? 'Excel' : 'Word') + '...</span>';

            try {
                const reqBody = {
                    ReportName: reportPayload.ReportName || reportPayload.reportName || (baseFileName + '.rpt'),
                    SelectionFormula: reportPayload.SelectionFormula !== undefined ? reportPayload.SelectionFormula : (reportPayload.selectionFormula || null),
                    GroupSelectionFormula: reportPayload.GroupSelectionFormula !== undefined ? reportPayload.GroupSelectionFormula : (reportPayload.groupSelectionFormula || null),
                    Parameters: reportPayload.Parameters || reportPayload.parameters || null,
                    FormulaValues: reportPayload.FormulaValues || reportPayload.formulaValues || null,
                    Format: format
                };
                const endpoint = apiOrigin + '/api/reports/export';
                const response = await fetch(endpoint, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(reqBody)
                });

                if (!response.ok) {
                    const errorText = await response.text();
                    throw new Error(errorText || ('Server responded with status ' + response.status));
                }

                const blob = await response.blob();
                const ext = isExcel ? '.xlsx' : '.doc';
                let downloadFileName = baseFileName + ext;
                const disposition = response.headers.get('content-disposition');
                if (disposition && disposition.indexOf('filename=') !== -1) {
                    const matches = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/.exec(disposition);
                    if (matches != null && matches[1]) {
                        downloadFileName = matches[1].replace(/['"]/g, '');
                    }
                }

                const downloadUrl = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = downloadUrl;
                a.download = downloadFileName;
                document.body.appendChild(a);
                a.click();
                a.remove();
                setTimeout(() => URL.revokeObjectURL(downloadUrl), 10000);

                showToast(downloadFileName + ' downloaded successfully!', false);
            } catch (err) {
                console.error('Export error:', err);
                showToast('Failed to export: ' + (err.message || err), true);
            } finally {
                btn.disabled = false;
                btn.innerHTML = originalHtml;
            }
        }

        async function exportToImage() {
            const btn = document.getElementById('btnImage');
            const originalHtml = btn.innerHTML;
            btn.disabled = true;
            btn.innerHTML = '<i class=\"fas fa-spinner fa-spin\"></i> <span>Rendering Image...</span>';

            try {
                if (typeof pdfjsLib === 'undefined') {
                    throw new Error('PDF rendering library is initializing. Please try again in a few seconds.');
                }

                pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

                const loadingTask = pdfjsLib.getDocument(pdfBlobUrl);
                const pdf = await loadingTask.promise;
                const totalPages = pdf.numPages;

                for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
                    const page = await pdf.getPage(pageNum);
                    // 2.0 scale provides high-DPI crisp print resolution
                    const viewport = page.getViewport({ scale: 2.0 });

                    const canvas = document.createElement('canvas');
                    const context = canvas.getContext('2d');
                    canvas.height = viewport.height;
                    canvas.width = viewport.width;

                    await page.render({ canvasContext: context, viewport: viewport }).promise;

                    const pageSuffix = totalPages > 1 ? ('_Page_' + pageNum) : '';
                    const fileName = baseFileName + pageSuffix + '.png';

                    const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
                    const downloadUrl = URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = downloadUrl;
                    a.download = fileName;
                    document.body.appendChild(a);
                    a.click();
                    a.remove();
                    setTimeout(() => URL.revokeObjectURL(downloadUrl), 10000);
                }

                showToast(baseFileName + '.png exported successfully!', false);
            } catch (err) {
                console.error('Image export error:', err);
                showToast('Failed to export image: ' + (err.message || err), true);
            } finally {
                btn.disabled = false;
                btn.innerHTML = originalHtml;
            }
        }

        function downloadPdf() {
            const a = document.createElement('a');
            a.href = pdfBlobUrl;
            a.download = baseFileName + '.pdf';
            document.body.appendChild(a);
            a.click();
            a.remove();
        }

        function printPdf() {
            try {
                const frame = document.getElementById('pdfFrame');
                if (frame && frame.contentWindow) {
                    frame.contentWindow.focus();
                    frame.contentWindow.print();
                } else {
                    window.print();
                }
            } catch (e) {
                console.warn('Direct print blocked, downloading PDF instead:', e);
                downloadPdf();
            }
        }
    </script>
</body>
</html>`;

        const htmlBlob = new Blob([viewerHtml], { type: 'text/html' });
        const viewerUrl = URL.createObjectURL(htmlBlob);
        const newWin = window.open(viewerUrl, '_blank');

        if (!newWin || newWin.closed || typeof newWin.closed === 'undefined') {
            console.warn("Popup blocked by browser. Triggering fallback download...");
            const anchor = document.createElement('a');
            anchor.href = pdfUrl;
            anchor.target = '_blank';
            anchor.download = cleanTitle + '.pdf';
            document.body.appendChild(anchor);
            anchor.click();
            anchor.remove();
        } else {
            console.log("Report Viewer opened successfully in new tab.");
        }
    } catch (err) {
        console.error("Error in openReportViewer:", err);
        // Fallback to direct PDF viewer
        window.openPdfFromBase64(base64);
    }
};

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
