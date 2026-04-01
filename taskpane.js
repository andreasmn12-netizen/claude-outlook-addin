/* global Office */

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    console.log("Claude Mail Assistant loaded.");
    const saved = localStorage.getItem("claude_api_key");
    if (saved) {
      document.getElementById("apiKeyInput").placeholder = "sk-ant-\u2026" + saved.slice(-4) + " (saved)";
    }
  }
});

function saveApiKey() {
  const input = document.getElementById("apiKeyInput");
  const key = input.value.trim();
  if (!key) return;
  localStorage.setItem("claude_api_key", key);
  input.value = "";
  input.placeholder = "sk-ant-\u2026" + key.slice(-4) + " (saved)";
  const msg = document.getElementById("apiKeyMsg");
  msg.textContent = "API key saved!";
  msg.classList.remove("hidden");
  setTimeout(() => msg.classList.add("hidden"), 3000);
}

async function improveEmail() {
  const btn = document.getElementById("improveBtn");
  const btnText = document.getElementById("btnText");
  const btnSpinner = document.getElementById("btnSpinner");
  const resultSection = document.getElementById("resultSection");
  const errorSection = document.getElementById("errorSection");

  resultSection.classList.add("hidden");
  errorSection.classList.add("hidden");

  const apiKey = localStorage.getItem("claude_api_key");
  if (!apiKey) {
    showError("No API key found. Paste your Anthropic API key in the field below and click Save.");
    return;
  }

  let emailText = "";
  try {
    emailText = await getEmailBody();
  } catch (err) {
    showError("Could not read the email body: " + err.message);
    return;
  }

  if (!emailText || emailText.trim().length < 5) {
    showError("The email appears to be empty. Write something first.");
    return;
  }

  btn.disabled = true;
  btnText.textContent = "Improving...";
  btnSpinner.classList.remove("hidden");

  const tone = document.getElementById("tone").value;
  const extra = document.getElementById("instructions").value.trim();
  const finalTone = extra ? tone + ". Additional instructions: " + extra : tone;

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "anthropic-dangerous-direct-browser-access": "true"
      },
      body: JSON.stringify({
        model: "claude-sonnet-4-6",
        max_tokens: 2048,
        system: "You are an expert email writer. Improve emails while preserving the sender intent. Return ONLY the improved email text with no explanations, preamble or commentary.",
        messages: [{ role: "user", content: "Improve this email. Make it " + finalTone + " and well-structured. Keep the same language as the original.\n\nOriginal email:\n" + emailText }]
      }),
    });

    if (!response.ok) {
      const errData = await response.json();
      showError("API error: " + (errData.error ? errData.error.message : response.status));
      return;
    }

    const data = await response.json();
    if (data.content && data.content[0]) {
      document.getElementById("resultText").value = data.content[0].text;
      resultSection.classList.remove("hidden");
      document.getElementById("replaceMsg").classList.add("hidden");
    } else {
      showError("Unexpected response from Claude.");
    }
  } catch (err) {
    showError("Request failed: " + err.message);
  } finally {
    btn.disabled = false;
    btnText.textContent = "Improve Email";
    btnSpinner.classList.add("hidden");
  }
}

async function replaceEmail() {
  const newText = document.getElementById("resultText").value;
  if (!newText) return;
  try {
    await setEmailBody(newText);
    const msg = document.getElementById("replaceMsg");
    msg.textContent = "Email replaced!";
    msg.classList.remove("hidden");
    setTimeout(() => msg.classList.add("hidden"), 3000);
  } catch (err) {
    showError("Could not replace email: " + err.message);
  }
}

function copyToClipboard() {
  const text = document.getElementById("resultText").value;
  if (!text) return;
  navigator.clipboard.writeText(text).then(() => {
    const btn = document.querySelector(".btn-icon");
    const original = btn.innerHTML;
    btn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"></polyline></svg> Copied!';
    setTimeout(() => (btn.innerHTML = original), 2000);
  });
}

// ---- Office.js helpers ----

function getEmailBody() {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item || !item.body) { reject(new Error("No mail item")); return; }
    item.body.getAsync(Office.CoercionType.Text, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
      else reject(new Error(result.error ? result.error.message : "Failed"));
    });
  });
}

function textToHtml(text) {
  return text.split(/\n{2,}/).map(p =>
    "<p>" + p.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\n/g, "<br>") + "</p>"
  ).join("");
}

function setEmailBody(text) {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item || !item.body) { reject(new Error("No mail item")); return; }
    item.body.setAsync(textToHtml(text), { coercionType: Office.CoercionType.Html }, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(result.error ? result.error.message : "Failed"));
    });
  });
}

function showError(message) {
  document.getElementById("errorMsg").textContent = message;
  document.getElementById("errorSection").classList.remove("hidden");
}
