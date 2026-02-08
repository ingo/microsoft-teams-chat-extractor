document.addEventListener('DOMContentLoaded', function () {
	var optionsDiv = document.getElementById('options');
	var statusDiv = document.getElementById('status');
	var progressText = document.getElementById('progressText');
	var chatDiv = document.getElementById('chat');
	var titleEl = document.getElementById('title');
	var toolbarDiv = document.getElementById('toolbar');
	var copyBtn = document.getElementById('copyBtn');
	var mdBtn = document.getElementById('mdBtn');
	var toast = document.getElementById('toast');
	var activeTabId = null;

	function showToast(text) {
		toast.textContent = text;
		toast.classList.add('show');
		setTimeout(function () { toast.classList.remove('show'); }, 2000);
	}

	// Convert the transcript HTML to Markdown
	function htmlToMarkdown(container) {
		var lines = [];
		container.querySelectorAll('.message').forEach(function (msg) {
			var hr = msg.querySelector('hr');
			var bold = msg.querySelector('b');
			var time = msg.querySelector('span');
			if (hr && bold) {
				lines.push('---');
				lines.push('**' + bold.textContent + '** ' + (time ? time.textContent : ''));
				lines.push('');
			}
			var divider = msg.querySelector('.divider');
			if (divider) {
				lines.push('---');
				lines.push(divider.textContent.trim());
				lines.push('---');
				lines.push('');
			}
			var section = msg.querySelector('section');
			if (section) {
				lines.push('> ' + section.innerText.trim().replace(/\n/g, '\n> '));
				lines.push('');
			}
		});
		return lines.join('\n');
	}

	// Copy transcript text to clipboard
	copyBtn.addEventListener('click', function () {
		var transcript = chatDiv.querySelector('#transcript');
		if (!transcript) return;
		var text = transcript.innerText;
		navigator.clipboard.writeText(text).then(function () {
			showToast('Copied!');
		});
	});

	// Export transcript as Markdown file
	mdBtn.addEventListener('click', function () {
		var transcript = chatDiv.querySelector('#transcript');
		if (!transcript) return;
		var md = htmlToMarkdown(transcript);
		var blob = new Blob([md], { type: 'text/markdown' });
		var url = URL.createObjectURL(blob);
		var a = document.createElement('a');
		a.href = url;
		a.download = 'teams-chat-' + new Date().toISOString().slice(0, 10) + '.md';
		a.click();
		URL.revokeObjectURL(url);
		showToast('Downloaded!');
	});

	// Listen for progress and result messages from the content script
	chrome.runtime.onMessage.addListener(function (message, sender) {
		if (!sender.tab || sender.tab.id !== activeTabId) return;

		if (message.type === 'progress') {
			progressText.textContent = message.count + ' messages collected so far\u2026';
		} else if (message.type === 'result') {
			statusDiv.style.display = 'none';
			toolbarDiv.style.display = 'flex';
			chatDiv.style.display = 'block';
			chatDiv.innerHTML = '<div id="transcript">' + (message.html || '<p>No messages found.</p>') + '</div>';
			if (message.count) {
				titleEl.textContent = 'Chat Transcript (' + message.count + ' messages)';
			}
		} else if (message.type === 'error') {
			statusDiv.style.display = 'none';
			chatDiv.style.display = 'block';
			chatDiv.innerHTML = '<p style="color:red;">' + message.error + '</p>';
		}
	});

	// Handle time-range button clicks
	optionsDiv.addEventListener('click', async function (e) {
		var btn = e.target.closest('button');
		if (!btn) return;

		var days = parseInt(btn.dataset.days, 10);
		var sort = document.querySelector('input[name="sort"]:checked').value;

		optionsDiv.style.display = 'none';
		statusDiv.style.display = 'block';
		chatDiv.style.display = 'none';

		if (days < 0) {
			progressText.textContent = 'Extracting loaded messages\u2026';
		} else if (days === 0) {
			progressText.textContent = 'Scrolling to load all messages\u2026';
		} else {
			progressText.textContent = 'Scrolling to load last ' + days + ' days\u2026';
		}

		var tabs = await chrome.tabs.query({ active: true, currentWindow: true });
		var tab = tabs[0];
		if (!tab) {
			statusDiv.style.display = 'none';
			chatDiv.style.display = 'block';
			chatDiv.innerHTML = '<p style="color:red;">No active tab found.</p>';
			return;
		}

		activeTabId = tab.id;

		try {
			await chrome.scripting.executeScript({
				target: { tabId: tab.id },
				files: ['payload.js']
			});

			await chrome.tabs.sendMessage(tab.id, { action: 'extract', days: days, sort: sort });
		} catch (err) {
			statusDiv.style.display = 'none';
			chatDiv.style.display = 'block';
			chatDiv.innerHTML = '<p style="color:red;">Could not extract chat. Make sure you are on a Microsoft Teams chat page (teams.microsoft.com).</p>';
			console.error('Extraction error:', err);
		}
	});
});
