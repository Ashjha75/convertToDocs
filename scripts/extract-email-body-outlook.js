javascript:(async function () {

  /* ---------- helpers ---------- */
  const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  // Simpler click, less overhead
  const clickElement = (el) => {
      el.click();
      // specific for some react apps
      el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
      el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
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

  const stripEmails = (t) =>
    t.replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, "[EMAIL]");

  /* ---------- expansion logic ---------- */
  
  async function expandAll() {
    console.log("Starting expansion...");
    
    // 1. Click "Show more messages" buttons (ONCE per batch)
    // We won't loop infinitely here to avoid freezing
    const loadMoreBtns = [...document.querySelectorAll('button')].filter(el => {
        const text = el.innerText ? el.innerText.toLowerCase() : "";
        return (text.includes('show') || text.includes('see')) && 
               (text.includes('messages') || text.includes('items'));
    });

    if (loadMoreBtns.length > 0) {
        console.log(`Found ${loadMoreBtns.length} 'Load More' buttons.`);
        for (const btn of loadMoreBtns) {
            try {
                btn.click();
                await delay(2000); 
            } catch (e) {}
        }
    }

    // 2. Expand individual collapsed messages
    // We will do exactly 2 passes to be safe against freezing
    for (let pass = 1; pass <= 2; pass++) {
        // Find all collapsed containers that look like emails
        // Based on user snippet: <div aria-expanded="false"> ... <div data-testid="SentReceivedSavedTime">
        const collapsedItems = [...document.querySelectorAll('[aria-expanded="false"]')]
            .filter(el => el.querySelector('[data-testid="SentReceivedSavedTime"]'));

        if (collapsedItems.length === 0) break;

        console.log(`Pass ${pass}: Found ${collapsedItems.length} collapsed emails.`);

        for (const item of collapsedItems) {
            try {
                // Scroll just enough
                item.scrollIntoView({block: "center", behavior: "instant"}); 
                
                // Click the container
                clickElement(item);
                
                // Also try the first child (header) if it exists
                if (item.firstElementChild) {
                    clickElement(item.firstElementChild);
                }

                // Small delay to let UI react, but not too long
                await delay(300); 
            } catch (e) {
                console.log("Error clicking item", e);
            }
        }
        // Wait for batch expansion
        await delay(2000);
    }
    console.log("Expansion complete.");
  }

  /* ---------- main ---------- */
  
  try {
      await expandAll();
  } catch (e) {
      console.error("Expansion failed:", e);
  }

  // Extraction Logic
  // We look for the timestamp, then go up to find the container
  const timestamps = [...document.querySelectorAll('[data-testid="SentReceivedSavedTime"]')];
  console.log(`Found ${timestamps.length} emails to extract.`);
  
  let output = "";
  let count = 1;

  // Sort timestamps by position in DOM to ensure chronological order (usually)
  // They should already be in order if querySelectorAll is used
  
  for (const timeEl of timestamps) {
      const dateTime = timeEl.innerText.trim();
      
      // Find the message container. 
      // In the snippet: timestamp -> div -> div -> div -> div(BS0OK/Expanded) -> div(aVla3)
      // We want the content. 
      // If expanded, the structure changes. We need to find the common parent that holds the body.
      // Usually, we can just look for the closest "list item" or just grab the parent container text.
      
      // Let's try to find the specific body container `role="document"` or `aria-label="Email message body"`
      // But since we are iterating by timestamp, we need to find the body associated with THIS timestamp.
      
      // Go up to the message container
      const msgContainer = timeEl.closest('[aria-label="Email message"]') || 
                           timeEl.closest('[role="listitem"]') || 
                           timeEl.closest('.aVla3') || // from snippet
                           timeEl.parentElement.parentElement.parentElement.parentElement;

      if (!msgContainer) continue;

      // Try to find the body text
      // In expanded view, there is usually a div with role="document" or similar
      let body = msgContainer.querySelector('[role="document"]') || 
                 msgContainer.querySelector('.allowTextSelection') || // from snippet (header part, but body might be sibling)
                 msgContainer;

      // If we only found the header part, we might need to look for siblings
      // But let's just extract text from the whole container and filter noise
      
      const walker = document.createTreeWalker(msgContainer, NodeFilter.SHOW_TEXT);
      let lines = [];
      let node;

      while (node = walker.nextNode()) {
        let text = node.nodeValue.replace(/\s+/g, " ").trim();
        if (!text) continue;
        if (isNoise(text)) continue;
        // Skip the timestamp itself to avoid duplication
        if (text === dateTime) continue;

        text = stripEmails(text);
        if (text) lines.push(text);
      }

      if (lines.length) {
        output += "\n--------------------------------------------------\n";
        output += `EMAIL ${count++}\n`;
        output += `DATE: ${dateTime}\n`;
        output += "--------------------------------------------------\n";
        output += lines.join("\n") + "\n";
      }
  }

  if (!output.trim()) {
    alert("No emails extracted.");
    return;
  }

  save(output, "email_thread.txt");

})();
