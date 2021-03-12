// send the page title as a chrome message
chrome.runtime.sendMessage(sendChat());

function sendChat() {

  var links = document.createElement("span");

  var threads = document.getElementsByClassName("ts-message-list-item");
  if (threads.length == 0) {
    window.alert("This does not appear to be a Microsoft Teams Chat window");
    return null;
  }

  for (let i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var messages = thread.querySelectorAll("div.message-body");
    var divider = thread.querySelector('div.message-list-divider');
    var node = document.createElement("div");

    // only display threaded view if in team chat
    for (let index = 0; index < messages.length; index++) {

      var message = messages[index];
      var name = message.querySelector('div.ts-msg-name');
      var timestamp = message.querySelector('span.message-datetime');
      var content = message.querySelector('div.message-body-content');
      var hidden = message.querySelector('.ts-collapsed-string > div[aria-hidden="true"]');

      var p = document.createElement("p");
      if(index > 0) {
        p.classList.add("reply");
      }
      if(name) {
        p.innerHTML += "<b>" + name.innerText + "</b>";
      }
      if(timestamp) {
        p.innerHTML += " " + timestamp.innerText;
      }
      if(content) {
        p.innerHTML += "<br>" + content.innerText;
      }
      if(hidden && hidden.innerText != "Collapse all") {
        p.classList.add("collapsed");
        p.innerHTML += hidden.innerText;
      }
      node.appendChild(p)
    }

    node.classList.add("thread");
    links.appendChild(node);

    if(divider) {
      var div = document.createElement("h2");
      div.innerText = divider.innerText;
      div.className = "divider";
      links.appendChild(div);
    }

  }

  return links.innerHTML;
}
