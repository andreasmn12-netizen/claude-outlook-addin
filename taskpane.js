/* global Office */

const SYSTEM_IMPROVE = "You are an expert email writer. Improve the email while keeping the sender's intent and language. Return ONLY the improved email text. Always separate paragraphs and sections with a blank line. No explanations or preamble.";
const SYSTEM_REPLY   = "You are an expert email writer. Write a helpful, well-structured reply to the email. Return ONLY the reply text. Always separate paragraphs with a blank line. No explanations or preamble.";

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    const saved = localStorage.getItem("claude_api_key");
    updateKeyStatus(saved);
    // Show settings on first use if no key saved
    if (!saved) {
      document.getElementById("settingsPanel").classList.remove("hidden");
    }
  }
});

function updateKeyStatus(key) {
  const el = document.getElementById("keyStatus");
  if (key) {
    el.textContent = "Key saved (ends in ..." + key.slice(-4) + ")";
    el.style.color = "#2d7d46";
    document.getElementById("apiKeyInput").placeholder = "Enter a new key to replace...";
  } else {
    el.textContent = "No key saved yet";
    el.style.color = "#c0392b";
  }
}

function toggleSettings() {
  const panel = document.getElementById("settingsPanel");
  panel.classList.toggle("hidden");
}

function saveApiKey() {
  const input = document.getElementById("apiKeyInput");
  const key = input.value.trim();
  if (!key) return;
  localStorage.setItem("claude_api_key", key);
  input.value = "";
  updateKeyStatus(key);
  const msg = document.getElementById("apiKeyMsg");
  msg.classList.remove("hidden");
  setTimeout(() => {
    msg.classList.add("hidden");
    document.getElementById("settingsPanel").classList.add("hidden");
  }, 1500);
}

async function improveEmail() {
  const emailText = await fetchEmailBody("improveBtn", "improveSpinner", "improveText", "Improving...", "Improve Email");
  if (emailText === null) return;

  const tone = document.getElementById("tone").value;
  const extra = document.getElementById("instructions").value.trim();
  const finalTone = extra ? tone + ". Additional instructions: " + extra : tone;
  const userMsg = "Improve this email. Make it " + finalTone + " and well-structured. Keep the same language as the original.\n\nOriginal email:\n" + emailText;

  await callClaude("improveBtn", "improveSpinner", "improveText", "Improving...", "Improve Email", SYSTEM_IMPROVE, userMsg, "Improved version");
}

async function suggestReply() {
  const emailText = await fetchEmailBody("replyBtn", "replySpinner", "replyText", "Thinking...", "Suggest Reply");
  if (emailText === null) return;

  const tone = document.getElementById("tone").value;
  const extra = document.getElementById("instructions").value.trim();
  const finalTone = extra ? tone + ". Additional instructions: " + extra : tone;
  const userMsg = "Write a reply to this email. Make it " + finalTone + ". Keep the same language as the original email.\n\nEmail to reply to:\n" + emailText;

  await callClaude("replyBtn", "replySpinner", "replyText", "Thinking...", "Suggest Reply", SYSTEM_REPLY, userMsg, "Suggested reply");
}

async function fetchEmailBody(btnId, spinnerId, textId, loadingLabel, resetLabel) {
  document.getElementById("resultSection").classList.add("hidden");
  document.getElementById("errorSection").classList.add("hidden");

  const apiKey = localStorage.getItem("claude_api_key");
  if (!apiKey) {
    showError("No API key saved. Click the gear icon in the top right to add your key.");
    return null;
  }

  let emailText = "";
  try {
    emailText = await getEmailBody();
  } catch (err) {
    showError("Could not read the email body: " + err.message);
    return null;
  }

  if (!emailText || emailText.trim().length < 5) {
    showError("The email appears to be empty. Write something first.");
    return null;
  }

  return emailText;
}

async function callClaude(btnId, spinnerId, textId, loadingLabel, resetLabel, system, userMsg, resultLabel) {
  const btn = document.getElementById(btnId);
  const spinner = document.getElementById(spinnerId);
  const text = document.getElementById(textId);
  const apiKey = localStorage.getItem("claude_api_key");

  btn.disabled = true;
  text.textContent = loadingLabel;
  spinner.classList.remove("hidden");

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
        system: system,
        messages: [{ role: "user", content: userMsg }]
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
      document.getElementById("resultLabel").textContent = resultLabel;
      document.getElementById("resultSection").classList.remove("hidden");
      document.getElementById("replaceMsg").classList.add("hidden");
    } else {
      showError("Unexpected response from Claude.");
    }
  } catch (err) {
    showError("Request failed: " + err.message);
  } finally {
    btn.disabled = false;
    text.textContent = resetLabel;
    spinner.classList.add("hidden");
  }
}

async function replaceEmail() {
  const newText = document.getElementById("resultText").value;
  if (!newText) return;
  try {
    await setEmailBody(newText);
    const msg = document.getElementById("replaceMsg");
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
    btn.innerHTML = '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"></polyline></svg> Copied!';
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
