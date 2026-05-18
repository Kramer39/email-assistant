// AI Email Assistant — API Proxy
// Deploy this on Render.com (free tier)
// Set environment variable: ANTHROPIC_API_KEY = your key from console.anthropic.com

const express = require("express");
const fetch   = require("node-fetch");
const cors    = require("cors");
const app     = express();

app.use(cors());
app.use(express.json());

app.get("/", (req, res) => {
  res.send("AI Email Assistant proxy is running.");
});

app.post("/api/generate-email", async (req, res) => {
  const apiKey = process.env.ANTHROPIC_API_KEY;

  if (!apiKey) {
    return res.status(500).json({ error: "ANTHROPIC_API_KEY not set" });
  }

  try {
    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type":      "application/json",
        "x-api-key":         apiKey,
        "anthropic-version": "2023-06-01"
      },
      body: JSON.stringify({
        model:      "claude-sonnet-4-20250514",
        max_tokens: 1000,
        system:     "You are an expert email writing assistant. Always reply only with valid JSON containing 'subject' and 'body' fields. No markdown, no preamble.",
        messages:   [{ role: "user", content: req.body.prompt }]
      })
    });

    const data = await response.json();
    res.json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Proxy running on port ${PORT}`));
