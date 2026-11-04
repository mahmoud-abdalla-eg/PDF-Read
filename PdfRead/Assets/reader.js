// PdfRead Reader Bridge Script
window.PdfReadBridge = {
    postMessage: function(type, payload) {
        if (window.chrome && window.chrome.webview) {
            window.chrome.webview.postMessage({ type: type, payload: payload });
        }
    },
    notifyReady: function() {
        this.postMessage("VIEWER_READY", {});
    }
};

document.addEventListener("DOMContentLoaded", function() {
    window.PdfReadBridge.notifyReady();
});
