# Microsoft Teams Chat Extractor

![Transcript Icon](./icons/chat-128.png)

A Chrome extension to extract chat transcripts from Microsoft Teams Web (v2).

Microsoft Teams does not support easily copying or exporting chat transcripts. This extension lets you pull them out with a couple of clicks.

## Features

- Extract currently loaded messages instantly, or auto-scroll back to load messages from the **last 24 hours, 7 days, 30 days, 3 months**, or **all history**
- Sort output **oldest-first** or **newest-first**
- **Copy to clipboard** with one click
- **Export to Markdown** file
- Handles @mentions, emoji, quoted replies, and date dividers
- Dark-themed popup UI

## Installation

1. Download or clone this repository
2. Open `chrome://extensions` in Chrome
3. Enable **Developer mode** (top-right toggle)
4. Click **Load unpacked** and select the extension folder

## Usage

1. Log into Microsoft Teams on the web at [teams.microsoft.com](https://teams.microsoft.com)
2. Open the chat you want to extract
3. Click the purple chat icon in the Chrome toolbar
4. Choose a time range and sort order
5. Wait for the extension to scroll back and collect messages (a spinner shows progress)
6. Use the **Copy** or **Markdown** buttons in the toolbar to export the transcript

## How it works

The extension injects a content script into the active Teams tab that:

1. Locates the `#chat-pane-list` container in the Teams v2 DOM
2. If a time range is selected, finds the scrollable ancestor and repeatedly scrolls to the top to trigger Teams' infinite scroll, collecting message nodes via a `MutationObserver`
3. Deduplicates and sorts messages by their internal ID
4. Extracts author, timestamp, and body content from each message
5. Sends the resulting HTML back to the popup for display

## Permissions

- **activeTab** -- access the current tab only when you click the extension icon
- **scripting** -- inject the extraction script into the Teams page

No data leaves your browser. The extension has no background service worker, makes no network requests, and stores nothing.

## License

[MIT](LICENSE)
