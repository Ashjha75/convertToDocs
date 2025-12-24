javascript:(async function () {

  /* ---------- helpers ---------- */
  const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const simulateClick = (element) => {
    ['mouseover', 'mousedown', 'click', 'mouseup'].forEach(eventType => {
      element.dispatchEvent(new MouseEvent(eventType, {
        bubbles: true,
        cancelable: true,
        view: window,
        buttons: 1
      }));
    });
  };

  const save = (data, file) => {
    const blob = new Blob([data], { type: "text/plain" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = file;
    a.click();
  };

  const isNoise = (t) => (
    /external sender/i.test(t) ||
    /be careful with links/i.test(t) ||
    /caution/i.test(t) ||
    /warning/i.test(t) ||
    /inbound shield/i.test(t) ||
    /trustifi/i.test(t) ||
    /microsoft teams/i.test(t) ||
    /meeting id/i.test(t) ||
    /passcode/i.test(t) ||
    /join the meeting/i.test(t) ||
    /^(from:|to:|cc:|subject:|sent:|when:|location:)/i.test(t)
  );

  // Replace emails with a placeholder to preserve context but remove PII
  const stripEmails = (t) =>
    t.replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, "[EMAIL]");

  /* ---------- expansion logic ---------- */
  
  async function expandAll() {
    console.log("Starting expansion...");
    
    // 1. Click "Show more messages" buttons (for long threads)
    const findLoadMoreButtons = () => {
        const candidates = [
            ...document.querySelectorAll('button'), 
            ...document.querySelectorAll('[role="button"]'),
            ...document.querySelectorAll('.ms-Button')
        ];
        return candidates.filter(el => {
            const text = el.innerText ? el.innerText.toLowerCase() : "";
            return (text.includes('show') || text.includes('see') || text.includes('load')) && 
                   (text.includes('messages') || text.includes('items') || text.includes('history') || text.includes('more'));
        });
    };

    let loadMoreBtns = findLoadMoreButtons();
    let safety = 0;
    while (loadMoreBtns.length > 0 && safety < 15) {
        safety++;
        console.log(`Found ${loadMoreBtns.length} 'Load More' buttons.`);
        for (const btn of loadMoreBtns) {
            if (btn.offsetParent !== null) { 
                try {
                    console.log("Clicking 'Load More' button...");
                    btn.scrollIntoView({block: "center"});
                    await delay(500);
                    simulateClick(btn);
                    btn.click(); 
                    await delay(3000); 
                } catch (e) {
                    console.log("Error clicking button", e);
                }
            }
        }
        loadMoreBtns = findLoadMoreButtons();
    }

    // 2. Expand individual collapsed messages
    // Target specific structure provided by user: <div aria-expanded="false"> containing timestamp
    let expanded = true;
    let loops = 0;
    const MAX_LOOPS = 50; 

    while (expanded && loops < MAX_LOOPS) {
        expanded = false;
        loops++;
        
        // Find all collapsed containers that look like emails
        const collapsedItems = [...document.querySelectorAll('[aria-expanded="false"]')]
            .filter(el => el.querySelector('[data-testid="SentReceivedSavedTime"]'));

        console.log(`Found ${collapsedItems.length} collapsed emails (Loop ${loops})...`);

        for (const item of collapsedItems) {
            console.log("Expanding collapsed email...");
            try {
                item.scrollIntoView({block: "center"});
                await delay(200);
                
                // Click the container itself as it has the aria-expanded attribute
                simulateClick(item);
                item.click();
                
                // Also try clicking the first child div which is often the click target
                if (item.firstElementChild) {
                    simulateClick(item.firstElementChild);
                    item.firstElementChild.click();
                }

                expanded = true;
                await delay(1000); // Wait for UI expansion
            } catch (e) {
                console.log("Error clicking item", e);
            }
        }
        
        if (expanded) {
            await delay(2000); // Wait for DOM to settle
        }
    }
    console.log("Expansion complete.");
  }

  /* ---------- main ---------- */
  
  try {
      await expandAll();
  } catch (e) {
      console.error("Expansion failed:", e);
      alert("Auto-expansion failed. Running extraction on visible emails only.");
  }

  // Find messages using multiple strategies
  const getMessages = () => {
      // Strategy 1: Standard ARIA label
      let msgs = [...document.querySelectorAll('[aria-label="Email message"]')];
      
      // Strategy 2: Fallback to timestamp containers if Strategy 1 fails
      if (msgs.length === 0) {
          console.log("Standard selector failed. Trying fallback...");
          const timestamps = [...document.querySelectorAll('[data-testid="SentReceivedSavedTime"]')];
          // Find the closest container that looks like a message row
          msgs = timestamps.map(ts => {
              // Go up 4-5 levels to find the container. 
              // Based on snippet: timestamp -> div -> div -> div -> div(BS0OK) -> div(aVla3)
              return ts.closest('[role="listitem"]') || ts.closest('.aVla3') || ts.parentElement.parentElement.parentElement.parentElement;
          }).filter(x => x); // remove nulls
          
          // Deduplicate
          msgs = [...new Set(msgs)];
      }
      return msgs;
  };

  const messages = getMessages();
  console.log(`Found ${messages.length} messages to extract.`);
  
  let output = "";
  let count = 1;

  messages.forEach(msg => {

    // date & time (reliable)
    const timeEl = msg.querySelector('[data-testid="SentReceivedSavedTime"]');
    const dateTime = timeEl ? timeEl.innerText.trim() : "DATE NOT FOUND";

    // body container
    // Try standard role="document" first, then fallback to the message container itself
    let body = msg.querySelector('[role="document"]');
    
    // If no specific body container found, use the message element itself but exclude the header info if possible
    // For now, if we can't find role="document", we just take the whole text and rely on isNoise to clean it up
    if (!body) {
        body = msg;
    }

    const walker = document.createTreeWalker(body, NodeFilter.SHOW_TEXT);
    let lines = [];
    let node;

    while (node = walker.nextNode()) {
      let text = node.nodeValue.replace(/\s+/g, " ").trim();
      if (!text) continue;
      if (isNoise(text)) continue;

      text = stripEmails(text);
      if (text) lines.push(text);
    }

    if (!lines.length) return;

    output += "\n--------------------------------------------------\n";
    output += `EMAIL ${count++}\n`;
    output += `DATE: ${dateTime}\n`;
    output += "--------------------------------------------------\n";
    output += lines.join("\n") + "\n";
  });

  if (!output.trim()) {
    alert("No emails extracted. Ensure conversation is fully expanded.");
    return;
  }

  save(output, "email_thread.txt");

})();
