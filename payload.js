// send the page title as a chrome message
chrome.runtime.sendMessage(sendChat());

function sendChat() {
  var messages = document.getElementsByClassName("message-body");
  if (messages.length == 0) {
    window.alert("This does not appear to be a Microsoft Teams Chat window");
    return null;
  }

  var linksText = "";

  for (let index = 0; index < messages.length; index++) {

    var message = messages[index];
    var name = message.querySelector('.ts-msg-name');
    var timestamp = message.querySelector('.message-datetime');
    var content = message.querySelector('.message-body-content');

    linksText += "<p><b>" + name.innerText + "</b> " + timestamp.innerText + "<br>" + content.innerText + "</p>";
  }

  return linksText;
}
