// Inject the payload.js script into the current tab after the popout has loaded
window.addEventListener('load', function (evt) {
	chrome.tabs.executeScript(null, {
		file: 'payload.js'
	}, _=>chrome.runtime.lastError);
});

// Listen to messages from the payload.js script and write to popout.html
chrome.runtime.onMessage.addListener(function (message) {
	var chat = document.getElementById('chat');
	if(chat) {
		chat.innerHTML = message;
	}
});
