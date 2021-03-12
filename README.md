# Microsoft Teams Chat Extractor
 A Chrome extension to extract chats from Microsoft Teams Web

![Transcript Icon](./icons/chat-128.png)

Microsoft teams does not currently support easily copying transcripts from the interface. This is a much requested feature, as evidenced by https://microsoftteams.uservoice.com/forums/555103-public/suggestions/32184895-export-your-conversations-to-a-document (3500 votes, 22 pages of comments)

This is not a perfect solution, but it works pretty well.

1. Download the source code for this extension to your computer (user the Code > Download .zip menu option)
2. Install the extension into Chrome (see https://www.cnet.com/how-to/how-to-install-chrome-extensions-manually/). If you don't see the icon, see "Pinnning an Extension" below.
3. Log into Microsoft Teams on the web (https://teams.microsoft.com)
4. Go to the chat you wish to extract. Note that it only will extract the content that's loaded, so if you want a longer history, scroll upwards in your chat history until you've loaded all the content you like
5. Press the little purple chat button in the extension bar. A popup will appear with the content included
6. Copy and paste into whatever form you like.

## Pinning an Extension
1. Open the Extensions by clicking the puzzle icon next to your profile avatar.
2. A dropdown menu will appear, showing you all of your enabled extensions. Each extension will have a pushpin icon to the right of it.
3. To pin an extension to Chrome, click the pushpin icon so that the icon turns blue.
4. To unpin an extension, click the pushpin icon so that the icon turns to a gray outline.

## Chat Format
A Microsoft Teams chat has an interesting hierarchy of divs/spans.

~~~~
. div.ts-message-list-item
+-- div.message-body
|   +-- div.ts-msg-name
|   +-- span.message-datetime
|   +-- div.message-body-content
+-- .ts-collapsed-string > div[aria-hidden="true"] (If some messages are collapsed)
+-- div.message-body
|   +-- div.ts-msg-name
|   +-- span.message-datetime
|   +-- div.message-body-content
. div.message-list-divider
. div.ts-message-list-item
+-- div.message-body
|   +-- div.ts-msg-name
|   +-- span.message-datetime
|   +-- div.message-body-content
+-- div.message-body
|   +-- div.ts-msg-name
|   +-- span.message-datetime
|   +-- div.message-body-content
~~~~

Comments? Changes? Please file GitHub issues or PRs.
