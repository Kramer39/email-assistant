/* AI Email Assistant — Outlook Add-in
   Proxy: https://email-assistant-izl3.onrender.com
*/

const API_ENDPOINT = "https://email-assistant-izl3.onrender.com/api/generate-email";

let lastSubject = "";
let lastBody    = "";

Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    console.log("AI Email Assistant loaded in Outlook");
  }
});

async function generateEmail(modifier) {
  const brief     = document.getElementById("brief").value.trim();
  const tone      = document.getElementById("tone").value;
  const emailType = document.getElementById("emailType").value;
  const recipient = document.getElementById("recipient").value.trim();

  if (!brief) {
    showError("Please describe what you want to say first.");
    return;
  }

  setLoading(true);

  const modifierNote  = modifier  ? `\nAdjustment: ${modifier} the email.` : "";
  const recipientNote = recipient ? `\nRecipient: ${recipient}` : "";

  const userPrompt = `Write a ${tone} ${emailType} email.
Brief: ${brief}${recipientNote}${modifierNote}

Return ONLY valid JSON with no markdown, no preamble:
{"subject": "...", "body": "..."}

The body must be a complete, ready-to-send email. Use \\n for line breaks.`;

  try {
    const response = await fetch(API_ENDPOINT, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ prompt: userPrompt }),
    });

    if (!response.ok) throw new Error(`Server error: ${response.status}`);

    const data   = await response.json();
    const text   = (data.content || []).filter(b => b.type === "text").map(b => b.text).join("");
    const clean  = text.replace(/```json|```/g, "").trim();
    const parsed = JSON.parse(clean);

    lastSubject = parsed.subject;
    lastBody    = parsed.body;

    document.getElementById("subjectOut").textContent      = parsed.subject;
    document.getElementById("bodyOut").textContent         = parsed.body;
    document.getElementById("resultSection").style.display = "block";
    hideError();
  } catch (err) {
    showError("Could not generate email. Check your connection and try again.\n" + err.message);
  } finally {
    setLoading(false);
  }
}

function insertIntoEmail() {
  const item = Office.context.mailbox.item;

  item.subject.setAsync(lastSubject, function (subjectResult) {
    if (subjectResult.status === Office.AsyncResultStatus.Failed) {
      showError("Could not set subject: " + subjectResult.error.message);
      return;
    }
    item.body.setAsync(lastBody, { coercionType: Office.CoercionType.Text }, function (bodyResult) {
      if (bodyResult.status === Office.AsyncResultStatus.Failed) {
        showError("Could not set body: " + bodyResult.error.message);
        return;
      }
      const btn = document.querySelector(".insert-btn");
      btn.textContent = "✓ Inserted!";
      setTimeout(() => { btn.textContent = "✓ Insert into email"; }, 2000);
    });
  });
}

function refine(modifier) {
  generateEmail(modifier);
}

function setLoading(isLoading) {
  document.getElementById("loading").style.display      = isLoading ? "block" : "none";
  document.getElementById("generateBtn").disabled       = isLoading;
  if (isLoading) document.getElementById("resultSection").style.display = "none";
}

function showError(msg) {
  const el = document.getElementById("errorMsg");
  el.textContent   = msg;
  el.style.display = "block";
}

function hideError() {
  document.getElementById("errorMsg").style.display = "none";
}
