javascript:(function () {

  /* ---------- helpers ---------- */
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
    t.replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, "");

  /* ---------- main ---------- */

  const messages = [...document.querySelectorAll('[aria-label="Email message"]')];
  let output = "";
  let count = 1;

  messages.forEach(msg => {

    // date & time (reliable)
    const timeEl = msg.querySelector('[data-testid="SentReceivedSavedTime"]');
    const dateTime = timeEl ? timeEl.innerText.trim() : "DATE NOT FOUND";

    // body container
    const body = msg.querySelector('[role="document"]');
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

})();
