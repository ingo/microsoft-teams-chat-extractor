(function () {
  // Guard against multiple injections
  if (window._teamsExtractorLoaded) return;
  window._teamsExtractorLoaded = true;

  // Send a message to the popup, silently ignoring errors if it's closed
  function sendMsg(data) {
    chrome.runtime.sendMessage(data).catch(function () {});
  }

  function sleep(ms) {
    return new Promise(function (resolve) { setTimeout(resolve, ms); });
  }

  // Find the scroll container for the chat list.
  // Teams v2 may use overflow:auto, overflow:scroll, or overflow:hidden
  // on a virtualised wrapper. We try multiple strategies.
  function findScrollContainer(el) {
    // Strategy 1: standard scrollable ancestor (auto / scroll)
    var current = el;
    while (current && current !== document.documentElement) {
      var style = getComputedStyle(current);
      var ov = style.overflowY;
      if ((ov === 'auto' || ov === 'scroll') && current.scrollHeight > current.clientHeight) {
        return current;
      }
      current = current.parentElement;
    }
    // Strategy 2: any ancestor that is already scrolled (catches hidden containers)
    current = el;
    while (current && current !== document.documentElement) {
      if (current.scrollTop > 0 && current.scrollHeight > current.clientHeight) {
        return current;
      }
      current = current.parentElement;
    }
    // Strategy 3: the element itself if it has overflow content
    if (el.scrollHeight > el.clientHeight) return el;
    return null;
  }

  // Return the earliest timestamp found among collected nodes
  function getOldestTimestamp(nodes) {
    var oldest = null;
    for (var i = 0; i < nodes.length; i++) {
      var timeEl = nodes[i].querySelector('[id^="timestamp-"]');
      if (timeEl) {
        var dt = new Date(timeEl.getAttribute('datetime'));
        if (!oldest || dt < oldest) oldest = dt;
      }
    }
    return oldest;
  }

  // Count unique message IDs among collected nodes
  function countUnique(nodes) {
    var ids = new Set();
    for (var i = 0; i < nodes.length; i++) {
      var msg = nodes[i].querySelector('[id^="message-body-"]');
      if (msg) ids.add(msg.id);
    }
    return ids.size;
  }

  // Deduplicate nodes by message-body ID, filter out GIFs, and sort chronologically
  function filterAndSort(nodes) {
    var map = new Map();
    nodes.forEach(function (n) {
      var msg = n.querySelector('[id^="message-body-"]');
      if (msg && !n.querySelector('[aria-label="Animated GIF"]')) {
        var id = parseInt(msg.id.replace('message-body-', ''), 10);
        if (!map.has(id)) map.set(id, n);
      }
    });
    return Array.from(map.entries())
      .sort(function (a, b) { return a[0] - b[0]; })
      .map(function (entry) { return entry[1]; });
  }

  // Replace emoji <img> tags with their alt-text
  function replaceEmojiImages(node) {
    node.querySelectorAll('img[itemtype*="Emoji"]').forEach(function (img) {
      var span = document.createElement('span');
      span.innerText = img.alt || '';
      img.parentNode.replaceChild(span, img);
    });
  }

  // Turn block-level @mention divs into inline spans
  function replaceMentions(node) {
    node.querySelectorAll('div[aria-label*="Mention"]').forEach(function (div) {
      var span = document.createElement('span');
      span.innerHTML = div.innerHTML;
      span.className = div.className;
      span.style.fontWeight = 'bold';
      div.parentNode.insertBefore(span, div);
      div.parentNode.removeChild(div);
    });
  }

  // Convert quoted-reply wrappers into <blockquote>
  function replaceQuotedReplies(node) {
    node.querySelectorAll('div[data-track-module-name="messageQuotedReply"]').forEach(function (div) {
      var blockquote = document.createElement('blockquote');
      blockquote.innerHTML = div.innerHTML;
      blockquote.className = div.className;
      div.parentNode.insertBefore(blockquote, div);
      div.parentNode.removeChild(div);
    });
  }

  // Build the final HTML transcript from an array of sorted message nodes
  function buildTranscript(nodes) {
    var lastAuthor = '';
    var lastDate = null;
    var output = document.createElement('div');

    nodes.forEach(function (n) {
      // Clone so we don't mutate the live Teams DOM
      var clone = n.cloneNode(true);
      replaceEmojiImages(clone);
      replaceMentions(clone);
      replaceQuotedReplies(clone);

      var authorEl = clone.querySelector('[data-tid="message-author-name"]');
      var timeEl = clone.querySelector('[id^="timestamp-"]');
      var bodyEl = clone.querySelector('[id^="message-body-"] [id^="content-"]');

      if (!authorEl || !timeEl || !bodyEl) return;

      var author = authorEl.innerText.trim();
      var ts = new Date(timeEl.getAttribute('datetime'));
      var tsStr = ts.toLocaleString();
      var date = tsStr.split(',')[0];
      var newDay = lastDate && ts.toDateString() !== lastDate.toDateString();

      var messageDiv = document.createElement('div');
      messageDiv.className = 'message';

      if (author !== lastAuthor) {
        messageDiv.appendChild(document.createElement('hr'));
        var b = document.createElement('b');
        b.textContent = author;
        messageDiv.appendChild(b);
        var timeSpan = document.createElement('span');
        timeSpan.textContent = ' [' + tsStr + ']:';
        messageDiv.appendChild(timeSpan);
        messageDiv.appendChild(document.createElement('br'));
      } else if (newDay) {
        var divider = document.createElement('div');
        divider.className = 'divider';
        divider.innerHTML = '<hr/>' + date + '<hr/>';
        messageDiv.appendChild(divider);
      }

      lastAuthor = author;
      lastDate = ts;

      var section = document.createElement('section');
      section.innerHTML = bodyEl.innerHTML;
      messageDiv.appendChild(section);

      output.appendChild(messageDiv);
    });

    return output.innerHTML;
  }

  // ---- Main extraction routine ----
  // days = -1 : extract only currently-loaded messages (no scrolling)
  // days =  0 : scroll all the way to the beginning
  // days >  0 : scroll back until we pass the cutoff date
  async function scrollAndExtract(days, sort) {
    try {
      var list = document.getElementById('chat-pane-list');
      if (!list) {
        sendMsg({ type: 'error', error: 'No chat pane found. Make sure you have a chat open in Teams.' });
        return;
      }

      var needsScroll = days >= 0;
      var cutoffDate = days > 0 ? new Date(Date.now() - days * 86400000) : null;

      var collected = [];
      var collect = function () { collected.push.apply(collected, Array.from(list.children)); };
      collect();

      if (needsScroll) {
        var scrollContainer = findScrollContainer(list);
        if (scrollContainer) {
          var obs = new MutationObserver(function () { collect(); });
          obs.observe(list, { childList: true, subtree: true, characterData: true });

          var noChangeCount = 0;
          var prevOldest = null;
          while (true) {
            // Scroll up using scrollTop and keyboard fallback
            scrollContainer.scrollTop = 0;
            list.dispatchEvent(new KeyboardEvent('keydown', {
              key: 'Home', code: 'Home', bubbles: true, cancelable: true
            }));

            await sleep(1500);
            collect();

            // If we have a date target, check whether we've scrolled past it
            if (cutoffDate) {
              var oldest = getOldestTimestamp(collected);
              if (oldest && oldest <= cutoffDate) break;
            }

            // Detect when no new messages are loading by tracking the
            // oldest timestamp. Virtual lists keep constant scrollHeight
            // so we can't rely on that.
            var currentOldest = getOldestTimestamp(collected);
            var changed = !prevOldest || !currentOldest
              || currentOldest.getTime() !== prevOldest.getTime();
            if (!changed) {
              noChangeCount++;
              if (noChangeCount >= 3) break;
            } else {
              noChangeCount = 0;
            }
            prevOldest = currentOldest;

            sendMsg({ type: 'progress', count: countUnique(collected) });
          }

          obs.disconnect();
        }
      }

      // Deduplicate and sort
      var nodes = filterAndSort(collected);

      // Trim to the requested date range
      if (cutoffDate) {
        nodes = nodes.filter(function (n) {
          var timeEl = n.querySelector('[id^="timestamp-"]');
          if (!timeEl) return true;
          return new Date(timeEl.getAttribute('datetime')) >= cutoffDate;
        });
      }

      if (nodes.length === 0) {
        sendMsg({ type: 'error', error: 'No messages could be extracted.' });
        return;
      }

      if (sort === 'newest') nodes.reverse();

      var html = buildTranscript(nodes);
      sendMsg({ type: 'result', html: html, count: nodes.length });
    } catch (e) {
      sendMsg({ type: 'error', error: 'Extraction failed: ' + e.message });
    }
  }

  // Listen for extraction requests from the popup
  chrome.runtime.onMessage.addListener(function (message, sender, sendResponse) {
    if (message.action === 'extract') {
      scrollAndExtract(message.days, message.sort);
      sendResponse({ status: 'started' });
    }
    return false;
  });
})();
