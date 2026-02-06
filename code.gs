const API_KEY = 'API_KEY'; 
const FOLDER_ID = 'FOLDER_ID';

function morningPodcastWorkflow() {
  const threads = GmailApp.search('label:newsletters newer_than:1d');
  
  if (threads.length === 0) {
    console.log("No newsletters found today.");
    return;
  }

  let combinedText = threads.map(t => t.getMessages().pop().getPlainBody()).join("\n\n");

  // --- GOOGLE DOC SUMMARY LOGIC (ORIGINAL PROMPT) ---
  const docPrompt = `You are a geopolitical, technology, and business analyst. I will give you distinct newsletters to summarize. Using only that text, do the following: Create a structured, bulleted script of the key points, grouped by major theme or section (for example: Geopolitics, AI & Big Tech, Corporate Legal Battles, Business & Fintech, Transportation & Public Safety Tech). Under each section, use concise bullet points that capture the main takeaways, key facts, and implications. Make sure to include a section somewhere for all the deals that were done related to M&A, VC and Private Equity investing. Rewrite it as a clear, spoken‑style script I can read aloud, using natural phrasing that would sound good in a podcast or presentation. Keep it under 20 minutes of reading time. Preserve all key facts, numbers, names, and relationships between events. Do not repeat points if already mentioned once in your output. Use a neutral, analytical tone (no advocacy) and plain language suitable for a smart but non‑expert audience. . Now produce the final bulleted script. 
                     Use clear headings for different topics. Focus on the most important news.
                     Newsletters: ${combinedText}`;
  
  // Updated to use the 2.5 Flash Lite identifier
  const docSummary = callGemini(docPrompt, "gemini-2.5-flash-lite");
  
  const dateStr = new Date().toLocaleDateString().replace(/\//g, '_');
  const doc = DocumentApp.create(`Newsletter_Summary_${dateStr}`);
  doc.getBody().setText(docSummary);
  
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const docFile = DriveApp.getFileById(doc.getId());
  
  folder.addFile(docFile);
  DriveApp.getRootFolder().removeFile(docFile);
  
  console.log("Success! Your Google Doc summary (Gemini 2.5 Flash Lite) is in your Drive folder.");
}

function callGemini(prompt, model) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${API_KEY}`;
  
  const payload = { 
    contents: [{ 
      parts: [{ text: prompt }] 
    }] 
  };
  
  const options = { 
    method: 'post', 
    contentType: 'application/json', 
    payload: JSON.stringify(payload),
    muteHttpExceptions: true 
  };

  const res = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(res.getContentText());
  
  if (json.error) {
    throw new Error("Text Error: " + json.error.message + " (Status: " + json.error.status + ")");
  }
  
  if (!json.candidates || !json.candidates[0].content) {
    throw new Error("No response from AI. Check safety settings or prompt length.");
  }
  
  return json.candidates[0].content.parts[0].text;
}