# Security and privacy notes

- **No secrets in this repo.** API keys you enter in Settings are stored in your own Google account (Apps Script Script Properties) and are never sent back to the page. The page only learns whether a key exists (yes/no).
- **Your data stays in your Google account.** Tasks live in a Google Sheet in your Drive, calendar events in a calendar you own. There is no server run by the author.
- **Deploy with "Execute as: Me" and "Access: Only myself".** Anyone else who gets the link cannot open it. Do not change access to "Anyone" unless you understand what that exposes.
- **Never paste keys into Code.gs or Index.html**, and never commit a copy of your own deployed script or your "Data" Sheet.
- **AI calls go to the provider you choose** (Anthropic, OpenAI or xAI) using your own key. Task titles and any files you allow the chat to read are sent to them. You are billed by them. Nothing is auto-debited by this app.
- This is hobby software: it has had no security audit. Treat it that way.
- Found a problem? Open a private security advisory on the repository.
