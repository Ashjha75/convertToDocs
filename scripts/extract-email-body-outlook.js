javascript:(async function () {

  /* ---------- helpers ---------- */
  const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

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
            ...document.querySelectorAll('[role="button"]')
        ];
        return candidates.filter(el => {
            const text = el.innerText ? el.innerText.toLowerCase() : "";
            return (text.includes('show') || text.includes('see')) && 
                   (text.includes('messages') || text.includes('items') || text.includes('history'));
        });
    };

    let loadMoreBtns = findLoadMoreButtons();
    let safety = 0;
    while (loadMoreBtns.length > 0 && safety < 10) {
        safety++;
        console.log(`Found ${loadMoreBtns.length} 'Load More' buttons.`);
        for (const btn of loadMoreBtns) {
            if (btn.offsetParent !== null) { 
                try {
                    btn.click();
                    await delay(2500); // Increased wait time for network
                } catch (e) {
                    console.log("Error clicking button", e);
                }
            }
        }
        loadMoreBtns = findLoadMoreButtons();
    }

    // 2. Expand individual collapsed messages
    let expanded = true;
    let loops = 0;
    const MAX_LOOPS = 50; // Increased limit for long threads

    while (expanded && loops < MAX_LOOPS) {
        expanded = false;
        loops++;
        
        const messages = document.querySelectorAll('[aria-label="Email message"]');
        console.log(`Checking ${messages.length} messages for collapsed state (Loop ${loops})...`);

        for (const msg of messages) {
            // Check if body is present
            const body = msg.querySelector('[role="document"]');
            
            if (!body) {
                console.log("Expanding a collapsed message...");
                
                // Strategy: Find the specific expander element if possible
                const expander = msg.querySelector('[aria-expanded="false"]');
                
                try {
                    if (expander) {
                        expander.click();
                    } else {
                        // Fallback: Click the header (usually the first child div)
                        // Clicking the whole container 'msg' sometimes doesn't work in new Outlook
                        const header = msg.querySelector('[role="heading"]') || msg.firstElementChild;
                        if (header) header.click();
                        else msg.click();
                    }
                    
                    expanded = true;
                    await delay(800); // Wait for UI expansion
                } catch (e) {
                    console.log("Error clicking message", e);
                }
            }
        }
        
        if (expanded) {
            await delay(1000); 
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

  const messages = [...document.querySelectorAll('[aria-label="Email message"]')];
  let output = "";
  let count = 1;

  messages.forEach(msg => {

    // date & time (reliable)
    const timeEl = msg.querySelector('[data-testid="SentReceivedSavedTime"]');
    const dateTime = timeEl ? timeEl.innerText.trim() : "DATE NOT FOUND";

    // body container
    const body = msg.querySelector('[role="document"]');
    if (!body) {
        // If still no body after expansion, we skip or mark it
        output += `\n[SKIPPED EMAIL ${count++} - CONTENT NOT LOADED]\n`;
        return;
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
