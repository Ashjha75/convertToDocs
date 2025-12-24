javascript:(async function () {
  console.log("Starting expansion script...");

  /* ---------- helpers ---------- */
  const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  const clickElement = (el) => {
      el.click();
      el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
      el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }));
  };

  /* ---------- expansion logic ---------- */
  async function expandAll() {
    
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
    console.log("Expansion complete. Now run the extraction script.");
    alert("Expansion complete! Now run the extraction script.");
  }

  await expandAll();

})();