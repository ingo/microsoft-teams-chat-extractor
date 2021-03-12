// send the page title as a chrome message
chrome.runtime.sendMessage(sendChat());

function sendChat() {

  var linksText = "";

  var threads = document.getElementsByClassName("ts-message-list-item");
  if (threads.length == 0) {
    window.alert("This does not appear to be a Microsoft Teams Chat window");
    return null;
  }

  for (let i = 0; i < threads.length; i++) {
    var thread = threads[i];

    var messages = thread.querySelectorAll("div.ts-message-thread-body");
    var hidden = thread.querySelector('.ts-collapsed-string > div[aria-hidden="true"]');
    // only display threaded view if in team chat
    if(messages.length > 1)
    {
      linksText += "<div class='threaded'>"
    }

    for (let index = 0; index < messages.length; index++) {

      var message = messages[index];
      var name = message.querySelector('div.ts-msg-name');
      var timestamp = message.querySelector('span.message-datetime');
      var content = message.querySelector('div.message-body-content');


      var reply = "";
      if(index > 0) {
        reply = " class='reply' "
      }
      linksText += "<p" + reply + ">"
      if(name) {
        linksText += "<b>" + name.innerText + "</b> ";
      }
      if(timestamp) {
        linksText += timestamp.innerText;
      }
      if(content) {
        linksText += "<br>" + content.innerText;
      }
      if(hidden && hidden.innerText != "Collapse all") {
        linksText += "<br><div class='collapsed'>" + hidden.innerText + "</div>";
      }
      linksText += "</p>"
    }

    // only display threaded view if in team chat
    if(messages.length > 1)
    {
      linksText += "</div>"
    }
  }

  return linksText;
}
