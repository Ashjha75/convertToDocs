javascript:(async function () {

  /* ---------- helpers ---------- */
  const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const clickElement = (el) => {
      el.click();
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

  const isNoise = (t) => {
    const lower = t.toLowerCase();
    return (
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
      /^(from:|to:|cc:|subject:|sent:|when:|location:)/i.test(t) ||
      // UI Noise
      lower === "reply" ||
      lower === "reply all" ||
      lower === "forward" ||
      lower === "apps" ||
      lower === "more actions" ||
      lower.includes("view with a light background") ||
      lower.includes("save all to onedrive") ||
      lower.includes("download all") ||
      lower.includes("retention:") ||
      lower.includes("expires:") ||
      // Icon artifacts (single non-word characters)
      (t.length === 1 && !/[a-zA-Z0-9]/.test(t))
    );
  };

  const stripEmails = (t) =>
    t.replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, "[EMAIL]");

  /* ---------- expansion logic ---------- */
  async function expandAll() {
    console.log("Starting expansion...");
    
    // 1. Click "Show more messages" buttons
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
    // We do 2 passes to ensure everything is open
    for (let pass = 1; pass <= 2; pass++) {
        const collapsedItems = [...document.querySelectorAll('[aria-expanded="false"]')]
            .filter(el => el.querySelector('[data-testid="SentReceivedSavedTime"]'));

        if (collapsedItems.length === 0) {
            console.log("No collapsed items found.");
            break;
        }

        console.log(`Pass ${pass}: Found ${collapsedItems.length} collapsed emails. Expanding...`);

        for (const item of collapsedItems) {
            try {
                item.scrollIntoView({block: "center", behavior: "instant"}); 
                clickElement(item);
                if (item.firstElementChild) {
                    clickElement(item.firstElementChild);
                }
                await delay(300); 
            } catch (e) {
                console.log("Error clicking item", e);
            }
        }
        await delay(2000);
    }
    console.log("Expansion complete.");
  }

  /* ---------- extraction logic ---------- */
  function extractEmails() {
      console.log("Starting extraction...");
      const messages = [...document.querySelectorAll('[aria-label="Email message"]')];
      let output = "";
      let count = 1;

      messages.forEach(msg => {
        // date & time (reliable)
        const timeEl = msg.querySelector('[data-testid="SentReceivedSavedTime"]');
        const dateTime = timeEl ? timeEl.innerText.trim() : "DATE NOT FOUND";

        // body container
        let body = msg.querySelector('[role="document"]');
        
        // Fallback: Try to find the specific message content div if role="document" is missing
        if (!body) {
            body = msg.querySelector('.allowTextSelection');
        }

        // If still no body, we skip it
        if (!body) return;

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
  }

  /* ---------- main execution ---------- */
  try {
      await expandAll();
      // Small delay to ensure DOM is settled after last expansion
      await delay(1000);
      extractEmails();
  } catch (e) {
      console.error("Script failed:", e);
      alert("An error occurred: " + e.message);
  }

})();
