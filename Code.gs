/**********************************************************************
 *  VAHANAM TRAVEL DESK BOT  —  Isha Yoga Center, Coimbatore
 *  Google Apps Script + Gemini (free tier) + a live Google Sheet for prices.
 *
 *  ============  FIRST-TIME SETUP  (do this once, in order)  ============
 *  1. Paste this whole file into your Apps Script project, replacing the old code. Save.
 *  2. Put your Gemini API key(s) in GEMINI_API_KEYS below (each starts with AIza...).
 *       You can paste several keys from different Google accounts for more daily headroom. Save.
 *  3. In the function dropdown at the top, choose  setupPriceSheet  and click Run.
 *       - Approve the permissions it asks for (Sheets + Gmail).
 *       - This creates a Google Sheet called "Vahanam Prices" in your Drive, fills in
 *         every fare, and remembers it automatically. Open View > Logs to see its link.
 *       - Run it again any time safely: if the sheet already exists it will NOT duplicate.
 *  4. Choose  runTravelDeskBot  and Run once. It creates 4 Gmail labels automatically:
 *       taxibot, taxibot-done, complaint, spam-review.
 *  5. In Gmail, colour the labels: complaint = RED, spam-review = GREY. Make a
 *       complaint-resolved = GREEN label to drag handled complaints into.
 *  6. Keep your internal-domain Gmail filter as-is. Delete any old "summarize" filter/label.
 *  7. Set up a time trigger (clock icon on the left) to run  runTravelDeskBot  every 5-10 min.
 *
 *  ============  CHANGING A PRICE LATER  ============
 *  Just open the "Vahanam Prices" Google Sheet and edit the number. Save. Done.
 *  The bot reads the live sheet on every reply — no need to touch this code ever again.
 *  To add a place: add a new row. Category must be one of: City, Local, Area, Outstation.
 *
 *  ============  THE BIG SWITCH  ============
 *  MODE = "draft"  -> bot writes a draft reply for you to check & send (use this first).
 *  MODE = "send"   -> bot sends replies automatically. Switch only once you fully trust it.
 *
 *  HOW IT SORTS MAIL (4 buckets):
 *   - normal     -> auto-reply, label taxibot-done (you never see it)
 *   - complaint  -> ONE warm reply asking for details, label complaint (RED)
 *   - sensitive  -> SILENT (accident/injury/harassment/legal), label complaint (RED) — you act
 *   - spam/scam  -> SILENT (fake CBI etc.), label spam-review (GREY) — you report & delete
 **********************************************************************/


/*  ============  EDIT THESE TWO  ============  */
// Paste one or MORE Gemini keys (one per Google account). The bot uses the first;
// if it's rate-limited, it automatically falls back to the next, and so on.
// Add as many as you like — just keep each in quotes, separated by commas.
const GEMINI_API_KEYS = [
  "PASTE_KEY_1_HERE",
  // "PASTE_KEY_2_HERE",
  // "PASTE_KEY_3_HERE",
];

const MODE = "draft";   // "draft" or "send"


/*  ============  Settings (leave as-is)  ============  */
const GEMINI_MODEL     = "gemini-2.5-flash";
const PROCESS_LABEL    = "taxibot";
const DONE_LABEL       = "taxibot-done";
const COMPLAINT_LABEL  = "complaint";      // make this RED in Gmail
const SPAM_LABEL       = "spam-review";    // make this GREY in Gmail
const MAX_THREADS      = 5;
const PRICE_SHEET_NAME = "Vahanam Prices";
const PRICE_PROP_KEY   = "VAHANAM_PRICE_SHEET_ID";   // where the sheet's id is remembered

// Seconds to wait between handling each email, so we stay under the per-minute API limit.
// Free tier allows roughly 10 requests/minute, so 6 seconds is safe. After you upgrade to
// a paid key you can lower this to 1 (or even 0) for faster handling.
const SECONDS_BETWEEN_EMAILS = 6;


/*  ============  ONE-TIME: build the price Google Sheet  ============  */
function setupPriceSheet() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty(PRICE_PROP_KEY);
  if (existingId) {
    try {
      const ss = SpreadsheetApp.openById(existingId);
      Logger.log('Price sheet already exists. Open it here:\n' + ss.getUrl());
      return;
    } catch (e) {
      Logger.log('Saved sheet was missing, creating a fresh one...');
    }
  }

  const ss = SpreadsheetApp.create(PRICE_SHEET_NAME);
  const sheet = ss.getSheets()[0].setName("Fares");

  const header = ["Place","Category","Base (Non-AC / Drop)","AC","Up & Down","AC extra","Innova","Notes"];
  const data = [
    ["R S Puram", "City", 1000, 1100, "", "", "", ""],
    ["Railway Station (Coimbatore Junction)", "City", 1000, 1100, "", "", 2200, "Innova also covers nearby city centre"],
    ["Gandhipuram", "City", 1000, 1100, "", "", "", ""],
    ["Omni Bus Stand", "City", 1000, 1100, "", "", "", ""],
    ["MTP Bus Stand", "City", 1000, 1100, "", "", "", ""],
    ["Race Course", "City", 1000, 1100, "", "", "", ""],
    ["Podanur Railway Station", "City", 1000, 1100, "", "", "", ""],
    ["Ramanathapuram", "City", 1000, 1100, "", "", "", ""],
    ["Residency", "City", 1000, 1100, "", "", "", ""],
    ["Lakshmi Mills", "City", 1000, "", "", "", "", "Flat fare"],
    ["Maruthamalai", "City", 1000, 1100, "", "", "", ""],
    ["Puliyakulam", "City", 1000, 1100, "", "", "", ""],
    ["Kuniyamuthur", "City", 1000, 1100, "", "", "", ""],
    ["Sowripalayam", "City", 1000, 1100, "", "", "", ""],
    ["Nava India", "City", 1100, "", "", "", "", "Flat fare"],
    ["PSG Hospital", "City", 1100, "", "", "", "", "Flat fare"],
    ["Peelamedu", "City", 1100, "", "", "", "", "Flat fare"],
    ["Thudiyalur", "City", 1100, "", "", "", "", "Flat fare"],
    ["Singanallur", "City", 1200, "", "", "", "", "Flat fare"],
    ["Echanari temple", "City", 1300, "", "", "", "", "Flat fare"],
    ["Airport (Coimbatore International / CJB)", "City", 1400, "", "", "", 2500, ""],
    ["Le Meridien", "City", 1400, "", "", "", "", "Airport-zone hotel"],
    ["Lemon Tree", "City", 1400, "", "", "", "", "Airport-zone hotel"],
    ["Royal Care Hospital", "City", 1500, "", "", "", "", ""],
    ["Anaikati", "City", 1500, "", "", "", "", ""],
    ["Sulur", "City", 1700, "", "", "", "", ""],
    ["Avinashi", "City", 2000, "", "", "", "", ""],
    ["Palladam", "City", 2000, "", "", "", "", ""],
    ["Mettupalayam", "City", 2200, "", "", "", "", ""],
    ["Pollachi", "City", 2200, "", "", "", "", ""],
    ["Tiruppur", "City", 2200, "", "", "", "", ""],
    ["Palakkad", "City", 2400, "", "", "", "", "Toll + permit + parking extra"],
    ["Adiyogi", "Local", 150, "", 250, "", "", "Waiting Rs.60/hr"],
    ["Ashram Inside", "Local", 150, "", 250, "", "", "Waiting Rs.60/hr"],
    ["Poondi", "Local", 200, "", 300, "", "", ""],
    ["Semmedu", "Local", 200, "", 300, "", "", ""],
    ["Iruttu Pallam", "Local", 300, "", 350, "", "", ""],
    ["Narasipuram", "Local", 250, "", 350, "", "", ""],
    ["Alanthurai", "Local", 350, "", 450, "", "", ""],
    ["Karunya (University)", "Local", 350, "", 450, "", "", ""],
    ["Mathuvarayapuram", "Local", 300, "", 400, "", "", ""],
    ["Dvara", "Local", 300, "", 400, "", "", ""],
    ["Celebrity Resort", "Local", 400, "", 500, "", "", ""],
    ["Pooluvapatti", "Local", 400, "", 500, "", "", ""],
    ["Isha Vidhya", "Local", 350, "", 450, "", "", ""],
    ["Kovai Kutralam", "Local", 400, "", 650, "", "", ""],
    ["Sadivayal", "Local", 400, "", 600, "", "", ""],
    ["Madampatti", "Area", 600, 700, 850, 100, "", "Waiting Rs.60/hr"],
    ["Thondamuthur", "Area", 600, 700, 850, 100, "", ""],
    ["Perur (Pateeswarar temple)", "Area", 700, 800, 950, 100, "", ""],
    ["Selvapuram", "Area", 800, 900, 1000, 150, "", ""],
    ["Vadavalli", "Area", 800, 900, 1000, 150, "", ""],
    ["Ooty", "Outstation", "", "", "", "", 6000, "Approx; + tolls/permits/parking; advance needed"],
    ["Coonoor", "Outstation", "", "", "", "", 5200, "Approx; + tolls/permits/parking; advance needed"]
  ];

  const all = [header].concat(data);
  sheet.getRange(1, 1, all.length, header.length).setValues(all);
  sheet.getRange(1, 1, 1, header.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, header.length);

  props.setProperty(PRICE_PROP_KEY, ss.getId());
  Logger.log('DONE. Your price sheet is ready. Open and edit prices here:\n' + ss.getUrl());
}


/*  ============  Read the live sheet into text for the bot  ============
    GENERIC: reads EVERY column using the header row as the label. So you can add
    any new column to the sheet (e.g. "Waiting charge", "Per km rate") and the bot
    will read it automatically — no code change needed. The first column is the place
    name; if there is a column titled "Category" the fares are grouped under it.       */
function getPriceText_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(PRICE_PROP_KEY);
  if (!id) return "";
  let values;
  try {
    values = SpreadsheetApp.openById(id).getSheetByName("Fares").getDataRange().getValues();
  } catch (e) {
    Logger.log('Could not read price sheet: ' + e);
    return "";
  }
  if (values.length < 2) return "";

  const headers = values[0].map(function (h) { return (h || "").toString().trim(); });
  const catIdx = headers.indexOf("Category");   // -1 if there is no Category column

  let lines = [];
  let lastCat = "";
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const place = (row[0] || "").toString().trim();
    if (!place) continue;

    if (catIdx > -1) {
      const cat = (row[catIdx] || "").toString().trim();
      if (cat && cat !== lastCat) { lines.push("\n### " + cat + " fares"); lastCat = cat; }
    }

    // Build "Header: value" for every non-empty cell (skip the name and category columns).
    let parts = [];
    for (let c = 1; c < headers.length; c++) {
      if (c === catIdx) continue;
      const label = headers[c];
      const val = row[c];
      if (label && val !== "" && val !== null && val !== undefined) {
        parts.push(label + ": " + val);
      }
    }
    lines.push(place + " — " + parts.join(", "));
  }
  return lines.join("\n");
}


/*  ============  MAIN — the timer runs this  ============  */
function runTravelDeskBot() {
  // GUARD A — never let two runs work at the same time (stops overlap double-sends).
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) { Logger.log('Another run is already active. Skipping this one.'); return; }

  try {
    const priceText = getPriceText_();
    if (!priceText) {
      Logger.log('No prices found. Run setupPriceSheet first. Stopping (no replies sent).');
      return;   // FAIL SAFE: never reply without real prices
    }

    const processLabel   = getOrCreateLabel(PROCESS_LABEL);
    const doneLabel      = getOrCreateLabel(DONE_LABEL);
    const complaintLabel = getOrCreateLabel(COMPLAINT_LABEL);
    const spamLabel      = getOrCreateLabel(SPAM_LABEL);

    const myEmail = (Session.getEffectiveUser().getEmail() || "").toLowerCase();   // this desk's own address

    const query   = 'label:' + PROCESS_LABEL + ' -label:' + DONE_LABEL + ' -label:' + COMPLAINT_LABEL + ' -label:' + SPAM_LABEL;
    const threads = GmailApp.search(query, 0, MAX_THREADS);
    Logger.log('Mode: ' + MODE + '. Found ' + threads.length + ' email(s).');

    for (let i = 0; i < threads.length; i++) {
      const thread = threads[i];
      try {
        const messages    = thread.getMessages();
        const lastMessage = messages[messages.length - 1];
        const from        = (lastMessage.getFrom() || "").toLowerCase();

        // GUARD B — if the most recent message is OURS, we already handled this thread.
        // (Gmail search lags after re-labelling, so a handled thread can reappear here.)
        // Don't reply again — just make sure it's marked done and move on.
        if (myEmail && from.indexOf(myEmail) > -1) {
          thread.removeLabel(processLabel);
          thread.addLabel(doneLabel);
          Logger.log('Already replied (skipping duplicate): ' + thread.getFirstMessageSubject());
          continue;
        }

        const guestText = lastMessage.getPlainBody().slice(0, 4000);
        const decision  = askGemini(guestText, priceText);
        if (!decision) { Logger.log('No decision (will retry): ' + thread.getFirstMessageSubject()); continue; }

        // GUARD C — re-label BEFORE sending, so even if anything reruns, this thread
        // is already out of the search and cannot be picked up and answered twice.
        thread.removeLabel(processLabel);
        if (decision.type === "complaint" || decision.type === "sensitive") {
          thread.addLabel(complaintLabel);
        } else if (decision.type === "spam") {
          thread.addLabel(spamLabel);
        } else {
          thread.addLabel(doneLabel);
        }

        if (decision.reply && decision.reply.trim() !== "") {
          if (MODE === "send") lastMessage.reply(decision.reply);
          else                 lastMessage.createDraftReply(decision.reply);
        }
        Logger.log(decision.type.toUpperCase() + ' (' + MODE + '): ' + thread.getFirstMessageSubject());

      } catch (err) {
        Logger.log('Error on a thread: ' + err);
      }

      Utilities.sleep(SECONDS_BETWEEN_EMAILS * 1000);   // pause so we don't trip the per-minute limit
    }
    Logger.log('Done.');

  } finally {
    lock.releaseLock();
  }
}


/*  ============  Ask Gemini: classify + write the reply (with key fallback)  ============  */
function askGemini(guestText, priceText) {
  const systemText = RULES_KB + "\n\n## CURRENT FARES (live from your price sheet — quote ONLY these)\n" + priceText + "\n\n" + DECISION_RULES;
  const payload = {
    systemInstruction: { parts: [{ text: systemText }] },
    contents: [{ role: 'user', parts: [{ text: guestText }] }],
    generationConfig: { temperature: 0.3, responseMimeType: "application/json" }
  };
  const options = { method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true };

  // Try each key in turn. If one is rate-limited (429) or busy (503), fall to the next.
  for (let k = 0; k < GEMINI_API_KEYS.length; k++) {
    const key = GEMINI_API_KEYS[k];
    if (!key || key.indexOf("PASTE") === 0) continue;   // skip blank/placeholder keys

    const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + key;
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();

    if (code === 429 || code === 503) {
      Logger.log('Key #' + (k + 1) + ' is rate-limited/busy. Trying the next key...');
      continue;
    }
    if (code !== 200) {
      Logger.log('Gemini error (key #' + (k + 1) + ') ' + code + ': ' + response.getContentText());
      continue;
    }

    const data = JSON.parse(response.getContentText());
    if (!(data && data.candidates && data.candidates[0] && data.candidates[0].content && data.candidates[0].content.parts)) {
      Logger.log('Unexpected Gemini response: ' + response.getContentText());
      continue;
    }
    const rawText = data.candidates[0].content.parts[0].text || "";

    try {
      const clean = rawText.replace(/```json/gi, '').replace(/```/g, '').trim();
      const parsed = JSON.parse(clean);
      let type = "normal";
      if (parsed.type === "complaint") type = "complaint";
      if (parsed.type === "sensitive") type = "sensitive";
      if (parsed.type === "spam")      type = "spam";
      let reply = (typeof parsed.reply === "string") ? parsed.reply : "";
      if (type === "sensitive" || type === "spam") reply = "";   // SAFETY BELT: never auto-reply
      return { type: type, reply: reply };
    } catch (e) {
      Logger.log('Could not parse decision, sending to complaint pile. Raw: ' + rawText);
      return { type: "complaint", reply: "" };
    }
  }

  Logger.log('All keys were unavailable. Leaving this email for the next run.');
  return null;   // every key failed -> retry next run
}


/*  ============  Helper  ============  */
function getOrCreateLabel(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}


/*  ============  Rules & tone (edit wording here; prices live in the Sheet)  ============  */
const RULES_KB = `# Vahanam Travel Desk — Rules & Tone (Isha Yoga Center, Coimbatore)

You are the Vahanam Travel Desk, replying to guests about TAXI & TRAVEL only.
All fares come from the CURRENT FARES list given to you below. NEVER invent, estimate,
or calculate a fare. Only quote a fare that appears in that list, exactly.

## IRON RULES
1. Never invent or calculate a fare. Quote only fares from the CURRENT FARES list.
2. Match small spelling differences to listed places ("Adhiyogi" -> Adiyogi,
   "Marudamalai" -> Maruthamalai, "rail station" -> Railway Station). But if a place/route
   is NOT in the list, or you are unsure what the guest means -> do NOT guess. Use the
   FUNNEL LINE. When in doubt, funnel.
3. Always greet with "Namaskaram" and end with "Pranam."
4. All trips start or end at Isha Yoga Center.
5. The guest books themselves: https://vahanam.co/cab, or Counter 22 (Welcome Point),
   or call/WhatsApp 9442615436. You quote and guide; you do not confirm a booking.
6. The numbers are frozen; the words are warm. Be kind and human, never change a price.
7. Outstation trips: say the fare is APPROXIMATE, that tolls/permits/parking are extra,
   and an advance may be needed.
8. Never promise a specific driver or car model.
9. Keep replies short, warm, clear. Answer what was asked; don't dump the whole tariff.
   BUT always mention any extra charge that applies to the trip the guest asked about —
   e.g. waiting charge, AC charge, a per-km rate, or tolls/permits — if it is listed for
   that place in the CURRENT FARES.
10. You handle ONLY taxi/travel. For anything else (stay/cottage, school, health,
    Samskriti, programs), warmly say you handle only travel and give the right contact
    from DEPARTMENTS. For emergencies, give the emergency number first.

## TONE
Warm, respectful, calm, concise — like a kind person at a temple help desk.
Start "Namaskaram", end "Pranam". Quote the price plainly, then point to the booking link.
Reassure anxious guests simply and truthfully. Write rupees plainly like Rs.1400.

## VEHICLES
- Sedan (Glanza, Etios, Swift Dzire, Ciaz — guest can't pick the model): up to 4 passengers,
  limited luggage. If the vehicle isn't mentioned, give the Sedan price and mention Innova too.
- Innova ("SUV"): up to 6-7 passengers, more luggage. Quote an Innova price ONLY for routes
  where one is listed in CURRENT FARES. Innova is NOT available for short local trips — offer
  the Sedan instead. Innova Crysta is not part of standard booking — ask them to call 9442615436.
- Local trips have a waiting charge of Rs.60 per hour if the driver waits.

## GROUPS (offer multiple Sedans — usually the economical choice)
1 Sedan = up to 4 people. 5-8 -> 2 Sedans, 9-12 -> 3 Sedans. Total = the Sedan fare for that
route x the number of Sedans (multiplying one fixed fare by a small whole number is allowed;
this is not inventing a fare). Always offer the multiple-Sedan option for groups over 4; never
push the costlier vehicle. Where an Innova fare exists for that route, mention it too.

## HOTELS -> ZONE (recognise what guests say)
- Airport-zone hotels (use the Airport Sedan fare): Le Meridien, Lemon Tree, Fairfield by
  Marriott, Radisson Blu, Vivanta, Zibe by GRT, Ginger (Avinashi Road), The Residency Towers,
  The Orbis, Gokulam Park, ibis Coimbatore (all on Avinashi Road near the airport).
- City-centre hotels (use the city fare): hotels in R S Puram, Gandhipuram, Race Course,
  Residency area.
- If you cannot place a hotel in a known zone -> funnel. Do not guess.

## OUTSTATION
For any outstation trip: say the fare is approximate, tolls/permits/parking extra, and an
advance is usually needed. Only quote an outstation fare if it appears in CURRENT FARES (e.g.
Ooty, Coonoor). For any other outstation route, or any Sedan outstation fare -> do NOT guess,
use the FUNNEL LINE.

## FUNNEL LINE (when a place/route/fare is not in CURRENT FARES, or you're unsure)
Say warmly, never with a made-up number:
"Namaskaram. For this, our travel desk will share the exact details. Please book through
https://vahanam.co/cab, or WhatsApp/call us on 9442615436, and we'll take care of it. Pranam."

## FAQ
- Payment: cash or UPI, paid directly to the driver at the end of the trip (only these two).
- Advance: needed for outstation/long trips; not for local/city trips.
- Cancel: WhatsApp/call 9442615436 as early as possible.
- AC: available; small extra charge (shown in CURRENT FARES where it applies).
- Book: https://vahanam.co/cab, or WhatsApp/call 9442615436, or Counter 22, Welcome Point.

## KEY INFO
Book: https://vahanam.co/cab | Helpline (call/WhatsApp): 9442615436
Email: travel.desk@sadhguru.org | In person: Counter 22, Welcome Point.
All trips start or end at Isha Yoga Center.

## DEPARTMENTS (for NON-travel questions — name the contact, stay brief)
- Isha Yoga Center (general/stay info): 8300083111
- Isha Cottage (room/stay): support.cottage@sadhguru.org or 8300083111
- Isha Home School: 04222515444 / 04222515445, office@ishahomeschool.org
- Isha Samskriti: 04222515480, isha.samskriti@sadhguru.org
- Isha Health Solutions (9:30am-4:30pm, 7 days): 8300045333, op.healthsolutions@sadhguru.org
- Emergency: 83001 00100 (give this first for emergencies)
- Travel Desk (us): 9442615436 (use 9442615434 only if the first is unreachable)`;

const DECISION_RULES = `--- YOUR JOB ---
You are the Vahanam Travel Desk replying to a guest's email. Do TWO things.

STEP 1 — DECIDE the type of the message:
- "spam"       = junk, scams, phishing, or threats from strangers — NOT a real guest about a
                 real trip. Examples: fake "police/CBI case against you, submit documents",
                 demands for money, prize/lottery, random promotions, anything unrelated to
                 Isha taxi travel sent to threaten or cheat. Stay SILENT. A person will report
                 and delete it. If torn between "spam" and a real serious issue, choose
                 "sensitive" (never bin a possibly-real problem).
- "sensitive"  = something serious where an automatic reply could be hurtful, unsafe, or risky:
                 an accident or injury, harassment or misconduct, a threat to safety, a genuine
                 legal matter about a trip, or a person in real distress. STAY SILENT — a human
                 will reach out personally.
- "complaint"  = the guest is unhappy or reporting an ordinary problem: rude driver, rash
                 driving, billing dispute, lost item, being cheated, disappointment. These
                 deserve a warm reply asking for details. When unsure between "normal" and
                 "complaint", choose "complaint".
- "normal"     = a routine request: fare/price questions, booking help, vehicle choice,
                 payment/cancellation/FAQ, or a non-travel question you simply redirect.

STEP 2 — WRITE the reply ("reply"):
- For "spam" and "sensitive": leave "reply" as an empty string "". Write nothing.
- For "normal": answer using ONLY the rules and the CURRENT FARES list above. Never invent a
  fare. If unsure of a place/fare, use the funnel line. Start "Namaskaram", end "Pranam".
  Plain text, warm, short. Rupees plainly like Rs.1400. No asterisks or markdown.
- For "complaint": do NOT quote fares or try to solve it. Write a short, warm, sincere reply
  that (a) apologises and takes it seriously, (b) asks for the booking details to look into it:
  date and time of the trip, and the driver's name or vehicle number. Start "Namaskaram", end
  "Pranam". Promise no specific outcome. Invent nothing.

OUTPUT FORMAT — return ONLY this JSON object and nothing else (no markdown, no backticks):
{"type": "normal" OR "complaint" OR "sensitive" OR "spam", "reply": "the message (empty for sensitive and spam)"}`;
