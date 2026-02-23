const GEMINI_MODEL = 'gemini-2.5-flash-lite';
const MAX_RETRIES = 3;

function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    apiKey: props.getProperty('GEMINI_API_KEY') || 'API_KEY',
    folderId: props.getProperty('FOLDER_ID') || 'FOLDER_ID',
  };
}

function morningPodcastWorkflow() {
  const config = getConfig();
  const threads = GmailApp.search('label:newsletters newer_than:1d');

  if (threads.length === 0) {
    console.log('No newsletters found today.');
    return;
  }

  // Track already-processed thread IDs to skip duplicate runs
  const props = PropertiesService.getScriptProperties();
  const lastRunIds = JSON.parse(props.getProperty('LAST_RUN_THREAD_IDS') || '[]');
  const lastRunSet = new Set(lastRunIds);

  const newThreads = threads.filter(t => !lastRunSet.has(t.getId()));
  if (newThreads.length === 0) {
    console.log('All newsletters already processed in a previous run.');
    return;
  }

  // Extract text, skipping threads with no messages and deduplicating content
  const seen = new Set();
  const texts = [];
  for (const thread of newThreads) {
    const messages = thread.getMessages();
    if (messages.length === 0) continue;
    const body = messages[messages.length - 1].getPlainBody();
    if (!body) continue;
    const fingerprint = body.slice(0, 200);
    if (seen.has(fingerprint)) continue;
    seen.add(fingerprint);
    texts.push(body);
  }

  if (texts.length === 0) {
    console.log('No newsletter content to process.');
    return;
  }

  const combinedText = texts.join('\n\n');
  console.log(`Processing ${texts.length} newsletters (${combinedText.length} chars).`);

  const docPrompt = `You are a geopolitical, technology, and business analyst. I will give you distinct newsletters to summarize. Using only that text, do the following: Create a structured, bulleted script of the key points, grouped by major theme or section (for example: Geopolitics, AI & Big Tech, Corporate Legal Battles, Business & Fintech, Transportation & Public Safety Tech). Under each section, use concise bullet points that capture the main takeaways, key facts, and implications. Make sure to include a section somewhere for all the deals that were done related to M&A, VC and Private Equity investing. Rewrite it as a clear, spoken-style script I can read aloud, using natural phrasing that would sound good in a podcast or presentation. Keep it under 20 minutes of reading time. Preserve all key facts, numbers, names, and relationships between events. Do not repeat points if already mentioned once in your output. Use a neutral, analytical tone (no advocacy) and plain language suitable for a smart but non-expert audience. Now produce the final bulleted script. Use clear headings for different topics. Focus on the most important news.
Newsletters: ${combinedText}`;

  const docSummary = callGemini(docPrompt, GEMINI_MODEL, config.apiKey);

  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const doc = DocumentApp.create(`Newsletter_Summary_${dateStr}`);
  doc.getBody().setText(docSummary);

  // Move doc to target folder in one step using moveTo
  const docFile = DriveApp.getFileById(doc.getId());
  const folder = DriveApp.getFolderById(config.folderId);
  docFile.moveTo(folder);

  const recipient = Session.getEffectiveUser().getEmail();
  const subject = `Your Morning Briefing is Ready - ${dateStr}`;
  const body = `Good morning!

Your morning briefing has been generated successfully.

- Read the Summary: ${doc.getUrl()}
- View Folder (Briefings): ${folder.getUrl()}

Have a great day!`;

  MailApp.sendEmail(recipient, subject, body);

  // Persist processed thread IDs so re-runs within the same day are skipped
  const processedIds = newThreads.map(t => t.getId());
  props.setProperty('LAST_RUN_THREAD_IDS', JSON.stringify(processedIds));

  console.log(`Success! Summarized ${texts.length} newsletters into ${doc.getUrl()}`);
}

function callGemini(prompt, model, apiKey) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

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

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      const res = UrlFetchApp.fetch(url, options);
      const code = res.getResponseCode();
      const json = JSON.parse(res.getContentText());

      if (json.error) {
        // Non-retryable client errors (400, 403, etc.)
        if (code >= 400 && code < 500 && code !== 429) {
          throw new Error('Gemini API error: ' + json.error.message);
        }
        // Retryable: 429 (rate limit) or 5xx
        console.log(`Gemini API attempt ${attempt}/${MAX_RETRIES} failed (${code}): ${json.error.message}`);
        if (attempt === MAX_RETRIES) throw new Error('Gemini API failed after retries: ' + json.error.message);
        Utilities.sleep(Math.pow(2, attempt) * 1000);
        continue;
      }

      if (!json.candidates || !json.candidates[0] || !json.candidates[0].content) {
        throw new Error('No response from Gemini. Check safety settings or prompt length.');
      }

      return json.candidates[0].content.parts[0].text;

    } catch (e) {
      if (e.message.startsWith('Gemini API error:') || e.message.startsWith('No response from Gemini')) throw e;
      console.log(`Gemini API attempt ${attempt}/${MAX_RETRIES} threw: ${e.message}`);
      if (attempt === MAX_RETRIES) throw e;
      Utilities.sleep(Math.pow(2, attempt) * 1000);
    }
  }
}
