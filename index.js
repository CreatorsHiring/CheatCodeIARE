const express = require("express");
const path = require("path");
const session = require("express-session");
const dotenv = require("dotenv");
const { GoogleGenerativeAI, HarmCategory, HarmBlockThreshold } = require("@google/generative-ai");
const PptxGenJS = require("pptxgenjs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");

const { inject } = require("@vercel/analytics");

dotenv.config();
inject();


const app = express();
const port = 8080;

// ─── Middleware ───────────────────────────────────────────────────────────────
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(session({
    secret: process.env.SESSION_SECRET || "cheatcodeiare_secret_key",
    resave: false,
    saveUninitialized: true
}));

// ─── View Engine ─────────────────────────────────────────────────────────────
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

// ─── Static Files ─────────────────────────────────────────────────────────────
app.use(express.static(path.join(__dirname, "public")));
app.use('/fa', express.static(path.join(__dirname, 'node_modules/@fortawesome/fontawesome-free')));

// ─── Gemini Setup ─────────────────────────────────────────────────────────────
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);

// Fallback chain — tried in order when a model fails (429, 503, quota errors).
// Best/newest first, most stable last.
const MODEL_FALLBACK_CHAIN = [
    "gemini-2.5-flash",       // latest 2.5 flash (stable alias)
    "gemini-2.5-pro",         // latest 2.5 pro (stable alias)
    "gemini-2.0-flash",       // solid 2.0 flash
    "gemini-flash-latest",
    "gemini-pro-latest",
    "gemini-1.5-flash",       // very stable fallback
    "gemini-1.5-pro",         // older pro, widely available
];

const SAFETY_SETTINGS = [
    { category: HarmCategory.HARM_CATEGORY_HARASSMENT,        threshold: HarmBlockThreshold.BLOCK_NONE },
    { category: HarmCategory.HARM_CATEGORY_HATE_SPEECH,       threshold: HarmBlockThreshold.BLOCK_NONE },
    { category: HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT, threshold: HarmBlockThreshold.BLOCK_NONE },
    { category: HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT, threshold: HarmBlockThreshold.BLOCK_NONE },
];

// Returns true for errors that warrant trying the next model
function isRetryableError(err) {
    const msg = (err.message || "").toLowerCase();
    return (
        msg.includes("429") ||
        msg.includes("503") ||
        msg.includes("404") ||
        msg.includes("not found") ||
        msg.includes("quota") ||
        msg.includes("rate limit") ||
        msg.includes("overloaded") ||
        msg.includes("too many requests")
    );
}

/**
 * generateWithFallback(promptFn, useJsonMode)
 *
 * Tries each model in MODEL_FALLBACK_CHAIN in order.
 * promptFn(model) — receives the constructed model, must return a Promise.
 * useJsonMode — if true, adds responseMimeType: "application/json"
 *
 * Example:
 *   const result = await generateWithFallback(m => m.generateContent(prompt), true);
 */
async function generateWithFallback(promptFn, useJsonMode = false) {
    let lastError;
    for (const modelName of MODEL_FALLBACK_CHAIN) {
        try {
            const model = genAI.getGenerativeModel({
                model: modelName,
                generationConfig: useJsonMode ? { responseMimeType: "application/json" } : {},
                safetySettings: SAFETY_SETTINGS,
            });
            console.log(`[Gemini] Trying: ${modelName}`);
            const result = await promptFn(model);
            console.log(`[Gemini] Success: ${modelName}`);
            return result;
        } catch (err) {
            lastError = err;
            if (isRetryableError(err)) {
                console.warn(`[Gemini] ${modelName} failed (${(err.message || "").slice(0, 80)}). Trying next model...`);
                continue;
            }
            throw err; // non-retryable — don't waste time on other models
        }
    }
    throw new Error(`All Gemini models exhausted. Last error: ${lastError?.message}`);
}

// ─── Auth Middleware ──────────────────────────────────────────────────────────
const isAuthenticated = (req, res, next) => {
    if (req.session.user) return next();
    // Store the intended destination before redirecting to login
    req.session.returnTo = req.originalUrl;
    res.redirect("/login");
};

// ─── PPT Content Generator (shared helper) ───────────────────────────────────
async function generatePptContent(problemStatement) {
    const prompt = `
    You are an expert academic assistant helping a university student prepare a high-quality AAT (Alternative Assessment Tool) PowerPoint presentation.

    Your task is to generate detailed, well-structured slide content for the following topic.

    Topic: ${problemStatement}

    Output STRICT JSON only — no markdown, no code fences, no extra text before or after the JSON.

    JSON schema:
    {
      "title": "A clear, professional title that matches the topic exactly",
      "introduction": "A detailed 4-6 sentence introduction covering: what the topic is, why it matters, its real-world relevance, and what the presentation will cover.",
      "index": [
        "Topic 1 name",
        "Topic 2 name",
        "Topic 3 name",
        "Topic 4 name",
        "Topic 5 name",
        "Topic 6 name",
        "Topic 7 name"
      ],
      "slides": [
        {
          "heading": "Slide heading matching the index item",
          "bulletPoints": [
            "First detailed point — explain the concept clearly in one complete sentence.",
            "Second point — include specific technical details, examples, or data where relevant.",
            "Third point — elaborate on applications, advantages, or how it works.",
            "Fourth point — cover challenges, limitations, or comparisons.",
            "Fifth point — real-world use case or industry relevance."
          ]
        }
      ],
      "conclusion": "A strong 4-5 sentence conclusion that summarises the key takeaways, the significance of the topic, what was learned, and future scope or recommendations."
    }

    Rules:
    - Generate exactly 10 index items and exactly 10 corresponding slides in the same order.
    - Each slide must have exactly 4-5 bullet points.
    - Every bullet point must be a complete, detailed sentence — not a fragment or keyword The sentences must be Bullet Points.
    - Do NOT use LaTeX. Write formulas in plain text (e.g., E = mc^2).
    - Tone must be academic, technical, and formal — suitable for a university-level AAT.
    - Do NOT include filler phrases like "In conclusion" or "As we can see".
    - Do NOT repeat the same information across slides.
    `;

    const result = await generateWithFallback(m => m.generateContent(prompt), true);
    let responseText = result.response.text();

    if (!responseText) throw new Error("Gemini returned an empty response.");

    responseText = responseText.replace(/```json/g, "").replace(/```/g, "").trim();
    return JSON.parse(responseText);
}

// ─── PPTX File Builder (shared helper) ───────────────────────────────────────
async function buildPptxBuffer(content, formData, user) {
    const pres = new PptxGenJS();

    const BG_COLOR       = "EDEDEE";
    const IARE_BLUE      = "003366";
    const TEXT_MAIN      = "000000";
    const NEW_BLUE       = "004170";
    const THANK_YOU_COLOR = "51237F";

    pres.layout = "LAYOUT_16x9";

    pres.defineSlideMaster({
        title: "TITLE_MASTER",
        background: { color: BG_COLOR },
        objects: [
            { image: { x: 0, y: 0, w: "100%", h: 1.5, path: path.join(__dirname, "assets", "image1.png") } }
        ]
    });

    pres.defineSlideMaster({
        title: "CONTENT_MASTER",
        background: { color: BG_COLOR },
        objects: [
            { image: { x: "85%", y: 0.05, w: 1.2, h: 0.7, path: path.join(__dirname, "assets", "image2.png") } },
            { rect: { x: 0, y: "95%", w: "100%", h: 0.4, fill: { color: IARE_BLUE } } },
            { slideNumber: { x: "90%", y: "96%", fontSize: 10, color: "FFFFFF" } }
        ]
    });

    // Title Slide
    const slide1 = pres.addSlide({ masterName: "TITLE_MASTER" });
    slide1.addText("AAT - TECH TALK", { x: 0.5, y: 1.8, w: "90%", fontSize: 28, color: NEW_BLUE, bold: true, align: "center", fontFace: "Arial" });
    slide1.addText(
        [
            { text: "Topic - ", options: { fontSize: 28, color: NEW_BLUE, fontFace: "Arial", bold: true } },
            { text: formData.problemStatement, options: { fontSize: 24, color: NEW_BLUE } }
        ],
        { x: 1.0, y: 2.5, w: "80%", align: "left" }
    );
    const details = [
        { label: "Name",    value: user.name },
        { label: "Roll No", value: user.studentId },
        { label: "Branch",  value: formData.department },
        { label: "Subject", value: formData.subject }
    ];
    details.forEach((item, i) => {
        slide1.addText(
            [
                { text: `${item.label}: `, options: { fontSize: 18, color: NEW_BLUE, bold: true } },
                { text: `${item.value}`,   options: { fontSize: 18, color: "000000" } }
            ],
            { x: 1.0, y: 3.5 + (i * 0.5), w: "80%", align: "left" }
        );
    });

    // Layout constants — keeps all slides consistent and prevents overflow
    // Slide is 10" wide x 7.5" tall (16x9). Footer bar starts at 95% ≈ 7.13"
    // Header image is 0.7" tall. Usable content area: y=0.85" to y=6.9"
    const HEADING_Y  = 0.82;   // heading sits just below the header image
    const HEADING_H  = 0.55;   // fixed height for heading row
    const CONTENT_Y  = 1.5;    // content starts here — clear gap below heading
    const CONTENT_H  = 5.25;   // ends at ~6.75", well above the footer bar
    const CONTENT_X  = 0.5;
    const CONTENT_W  = 9.0;    // full usable width (slide is 10")

    // Index Slide
    const slideIndex = pres.addSlide({ masterName: "CONTENT_MASTER" });
    slideIndex.addText("INDEX", {
        x: CONTENT_X, y: HEADING_Y, w: CONTENT_W, h: HEADING_H,
        fontSize: 22, color: IARE_BLUE, bold: true, valign: "middle"
    });
    const indexItems = content.index.map((item, idx) => ({
        text: `${idx + 1}.  ${item}
`,
        options: { fontSize: 16, color: TEXT_MAIN, bullet: false }
    }));
    slideIndex.addText(indexItems, {
        x: CONTENT_X, y: CONTENT_Y, w: CONTENT_W, h: CONTENT_H,
        valign: "top", lineSpacingMultiple: 1.3
    });

    // Introduction Slide
    const slideIntro = pres.addSlide({ masterName: "CONTENT_MASTER" });
    slideIntro.addText("INTRODUCTION", {
        x: CONTENT_X, y: HEADING_Y, w: CONTENT_W, h: HEADING_H,
        fontSize: 22, color: IARE_BLUE, bold: true, valign: "middle"
    });
    slideIntro.addText(content.introduction, {
        x: CONTENT_X, y: CONTENT_Y, w: CONTENT_W, h: CONTENT_H,
        fontSize: 15, color: TEXT_MAIN, valign: "top",
        wrap: true, lineSpacingMultiple: 1.4
    });

    // Content Slides
    content.slides.forEach(slideData => {
        const s = pres.addSlide({ masterName: "CONTENT_MASTER" });
        s.addText(slideData.heading.toUpperCase(), {
            x: CONTENT_X, y: HEADING_Y, w: CONTENT_W, h: HEADING_H,
            fontSize: 22, color: IARE_BLUE, bold: true, valign: "middle"
        });
        const bulletItems = slideData.bulletPoints.map(bp => ({
            text: bp,
            options: {
                bullet: { type: "bullet", characterCode: "2022" },
                fontSize: 14,
                color: TEXT_MAIN,
                paraSpaceAfter: 6,
            }
        }));
        s.addText(bulletItems, {
            x: CONTENT_X, y: CONTENT_Y, w: CONTENT_W, h: CONTENT_H,
            valign: "top", lineSpacingMultiple: 1.3
        });
    });

    // Conclusion Slide
    const slideConc = pres.addSlide({ masterName: "CONTENT_MASTER" });
    slideConc.addText("CONCLUSION", {
        x: CONTENT_X, y: HEADING_Y, w: CONTENT_W, h: HEADING_H,
        fontSize: 22, color: IARE_BLUE, bold: true, valign: "middle"
    });
    slideConc.addText(content.conclusion, {
        x: CONTENT_X, y: CONTENT_Y, w: CONTENT_W, h: CONTENT_H,
        fontSize: 15, color: TEXT_MAIN, valign: "top",
        wrap: true, lineSpacingMultiple: 1.4
    });

    // Thank You Slide
    const slideLast = pres.addSlide({ masterName: "TITLE_MASTER" });
    slideLast.addText("THANK YOU", { x: 0.5, y: 1.5, w: "90%", h: "50%", fontSize: 48, color: THANK_YOU_COLOR, bold: true, align: "center", valign: "middle" });

    return pres.write("nodebuffer");
}

// ═════════════════════════════════════════════════════════════════════════════
// ROUTES
// ═════════════════════════════════════════════════════════════════════════════

// ─── Home ─────────────────────────────────────────────────────────────────────
app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ─── Auth ─────────────────────────────────────────────────────────────────────
app.get("/login", (req, res) => {
    res.render("login");
});

app.post("/login", (req, res) => {
    const { name, studentId } = req.body;
    if (name && studentId) {
        req.session.user = { name, studentId };
        // Redirect to the intended destination or default to homepage
        const returnTo = req.session.returnTo || "/";
        req.session.returnTo = null; // Clear the returnTo after using it
        res.redirect(returnTo);
    } else {
        res.redirect("/login");
    }
});

// ─── Signup ───────────────────────────────────────────────────────────────────
app.get("/signup", (req, res) => {
    res.render("signup");
});

app.post("/signup", (req, res) => {
    const { name, studentId, email, password } = req.body;
    if (name && studentId && email && password) {
        // Store user info in session and redirect to homepage
        req.session.user = { name, studentId, email };
        res.redirect("/");
    } else {
        res.redirect("/signup");
    }
});

// ─── PPT Form Page ────────────────────────────────────────────────────────────
app.get("/ppt", isAuthenticated, (req, res) => {
    res.render("ppt", { user: req.session.user });
});

// ─── Generate PPT → Store in session → Redirect to /edit ─────────────────────
app.post("/generate-ppt", isAuthenticated, async (req, res) => {
    try {
        const { department, subject, problemStatement } = req.body;

        const content = await generatePptContent(problemStatement);

        // Store everything in session so /edit and /download-ppt can use it
        req.session.pptData = {
            department,
            subject,
            problemStatement,
            content
        };

        res.redirect("/edit");

    } catch (error) {
        console.error("PPT Generation Error:", error);
        res.status(500).send("Error generating PPT content. " + error.message);
    }
});

// ─── Edit Page ────────────────────────────────────────────────────────────────
app.get("/edit", isAuthenticated, (req, res) => {
    const pptData = req.session.pptData;

    if (!pptData) {
        // Nothing generated yet — send back to form
        return res.redirect("/ppt");
    }

    res.render("edit", {
        user: req.session.user,
        pptData: pptData
    });
});

// ─── Chatbot: General Q&A ─────────────────────────────────────────────────────
// Called by the chat sidebar with a free-form user message.
// It sends back a plain text reply — no PPT modification.
app.post("/api/chat", isAuthenticated, async (req, res) => {
    try {
        const { message, history } = req.body;

        if (!message || !message.trim()) {
            return res.status(400).json({ success: false, reply: "Empty message." });
        }

        const pptData = req.session.pptData;

        // Build a system-style preamble so the bot is aware of the presentation context
        const systemContext = pptData
            ? `You are a helpful AI assistant integrated into a PPT editor app called CheatCodeIARE.
The user has generated a presentation titled "${pptData.content?.title || pptData.problemStatement}".
Answer their questions helpfully and concisely. 
If they want to MODIFY slides, tell them to use the "Edit Slides" prompt box for that.`
            : `You are a helpful AI assistant integrated into a PPT editor app called CheatCodeIARE.
Answer the user's questions helpfully and concisely.`;

        // Build conversation history for multi-turn context
        const turns = [];
        if (history && Array.isArray(history)) {
            history.forEach(h => {
                // Gemini SDK only accepts "user" or "model" — normalize "assistant"
                const role = h.role === "assistant" ? "model" : h.role;
                turns.push({ role, parts: [{ text: h.text }] });
            });
        }

        // Chat uses sendMessage — fallback by rebuilding chat on each model
        let reply = "";
        let chatSuccess = false;
        for (const modelName of MODEL_FALLBACK_CHAIN) {
            try {
                const m = genAI.getGenerativeModel({ model: modelName, safetySettings: SAFETY_SETTINGS });
                const chatSession = m.startChat({ history: turns, systemInstruction: systemContext });
                const result = await chatSession.sendMessage(message);
                reply = result.response.text();
                chatSuccess = true;
                console.log(`[Gemini] Chat success: ${modelName}`);
                break;
            } catch (err) {
                if (isRetryableError(err)) {
                    console.warn(`[Gemini] Chat ${modelName} failed. Trying next...`);
                    continue;
                }
                throw err;
            }
        }
        if (!chatSuccess) throw new Error("All Gemini models failed for chat.");

        res.json({ success: true, reply });

    } catch (error) {
        console.error("Chat API Error:", error);
        res.status(500).json({ success: false, reply: "Sorry, I ran into a server error. Please try again." });
    }
});

// ─── Edit PPT via AI: Modify the slide content stored in session ───────────────
// Called when the user wants to change slides via a natural language prompt.
app.post("/api/edit-ppt", isAuthenticated, async (req, res) => {
    try {
        const { prompt } = req.body;

        if (!prompt || !prompt.trim()) {
            return res.status(400).json({ success: false, error: "Empty prompt." });
        }

        const pptData = req.session.pptData;
        if (!pptData) {
            return res.status(400).json({ success: false, error: "No active presentation found. Please generate one first." });
        }

        const editPrompt = `
You are an AI that edits PowerPoint presentation content based on user instructions.

Current presentation JSON:
${JSON.stringify(pptData.content, null, 2)}

User instruction: "${prompt}"

Apply the requested changes and return the COMPLETE updated JSON using the EXACT same schema:
{
  "title": "String",
  "introduction": "String",
  "index": ["String", ...],
  "slides": [{ "heading": "String", "bulletPoints": ["String", ...] }, ...],
  "conclusion": "String"
}

Rules:
- Return ONLY valid JSON, no markdown, no extra text.
- Keep all fields present even if not changed.
- Do not add or remove the top-level keys.
        `;

        const result = await generateWithFallback(m => m.generateContent(editPrompt), true);
        let responseText = result.response.text();

        if (!responseText) throw new Error("Gemini returned empty response.");

        responseText = responseText.replace(/```json/g, "").replace(/```/g, "").trim();
        const updatedContent = JSON.parse(responseText);

        // Persist changes back into the session
        req.session.pptData.content = updatedContent;

        res.json({ success: true, content: updatedContent });

    } catch (error) {
        console.error("Edit PPT API Error:", error);
        res.status(500).json({ success: false, error: "Failed to update presentation: " + error.message });
    }
});

// ─── Download PPT ─────────────────────────────────────────────────────────────
// Builds and streams the PPTX file from whatever is currently in the session.
app.post("/download-ppt", isAuthenticated, async (req, res) => {
    try {
        const pptData = req.session.pptData;

        if (!pptData) {
            return res.status(400).send("No presentation data found. Please generate a PPT first.");
        }

        const buffer = await buildPptxBuffer(
            pptData.content,
            {
                department:       pptData.department,
                subject:          pptData.subject,
                problemStatement: pptData.problemStatement
            },
            req.session.user
        );

        const safeFilename = `${req.session.user.name.replace(/\s+/g, '_')}_AAT.pptx`;

        res.setHeader("Content-Disposition", `attachment; filename="${safeFilename}"`);
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
        res.send(buffer);

    } catch (error) {
        console.error("Download PPT Error:", error);
        res.status(500).send("Error building PPTX file. " + error.message);
    }
});

// ─── Report Page ──────────────────────────────────────────────────────────────
app.get("/report", isAuthenticated, (req, res) => {
    res.render("report", { user: req.session.user });
});

// ─── Worksheets Page (Coming Soon) ───────────────────────────────────────────
app.get("/worksheets", (req, res) => {
    res.render("worksheets");
});

// ─── (Keep existing report generation logic below) ────────────────────────────

const markdownToWordXML = (text) => {
    if (!text) return "";

    let xml = text
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");

    const lines = xml.split('\n');
    let finalXml = "";
    let inTable = false;

    lines.forEach(line => {
        const trimmedLine = line.trim();
        if (!trimmedLine) return;

        if (trimmedLine.startsWith("TABLE:")) {
            inTable = true;
            finalXml += '<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>';
            return;
        }

        if (inTable) {
            if (line.includes("|")) {
                finalXml += "<w:tr>";
                const cells = line.split("|");
                cells.forEach(cell => {
                    const cellContent = cell.trim();
                    finalXml += `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>${cellContent}</w:t></w:r></w:p></w:tc>`;
                });
                finalXml += "</w:tr>";
            } else {
                inTable = false;
                finalXml += "</w:tbl>";
                finalXml += `<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t xml:space="preserve">${line}</w:t></w:r></w:p>`;
            }
        } else {
            finalXml += `<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t xml:space="preserve">${line}</w:t></w:r></w:p>`;
        }
    });

    if (inTable) finalXml += "</w:tbl>";

    return finalXml;
};

const batchGenerateAnswers = async (questions) => {
    const answers = new Array(10).fill("");

    // Helper: send one batch of questions and fill answers array
    const processBatch = async (batch, batchNum) => {
        const numberedQuestions = batch
            .map(({ q }, idx) => `Q${idx + 1}: ${q}`)
            .join("\n\n");

        const prompt = `
You are an AI academic content generator for a student-focused AAT report system.
Generate detailed, exam-ready answers for ALL questions below.

⚠️ Output must be PLAIN TEXT ONLY — no HTML, no Markdown, no XML tags.
Use CAPITALIZED HEADINGS and hyphen bullets (-) only.
For tables use:
TABLE:
Col A | Col B
Row1A | Row1B

Return ONLY a raw JSON array — no markdown fences, no extra text:
[
  { "index": 0, "answer": "answer for Q1 here" },
  { "index": 1, "answer": "answer for Q2 here" }
]

Each answer must be detailed and exam-ready (around 100-120 words).
The "index" must start at 0 and match the question order.

QUESTIONS:
${numberedQuestions}
        `;

        try {
            console.log(`Processing batch ${batchNum} (${batch.length} questions)...`);
            const result = await generateWithFallback(m => m.generateContent(prompt), false);
            let text = result.response.text()
                .replace(/```json/g, "")
                .replace(/```/g, "")
                .trim();

            const parsed = JSON.parse(text);
            parsed.forEach((item) => {
                // Use item.index (as returned by the model) to find the original question index
                const batchItem = batch[item.index];
                if (batchItem !== undefined && item.answer) {
                    answers[batchItem.i] = item.answer;
                }
            });
        } catch (e) {
            console.error(`Batch ${batchNum} error:`, e.message);
            batch.forEach(({ i }) => { answers[i] = "AI Generation Failed."; });
        }
    };

    // Filter out blank questions keeping original indices
    const nonEmpty = questions
        .map((q, i) => ({ q, i }))
        .filter(({ q }) => q && q.trim());

    if (nonEmpty.length === 0) return answers;

    // Split into 2 batches of 5 — avoids response truncation, still only 2 API calls
    const mid = Math.ceil(nonEmpty.length / 2);
    const batch1 = nonEmpty.slice(0, mid);
    const batch2 = nonEmpty.slice(mid);

    await processBatch(batch1, 1);
    if (batch2.length > 0) await processBatch(batch2, 2);

    return answers;
};

app.post("/generate-report", isAuthenticated, async (req, res) => {
    try {
        const formData = req.body;
        const questions = [];
        for (let i = 1; i <= 10; i++) {
            questions.push(formData[`question${i}`] || "");
        }

        // 2. Generate Answers (Batched)
        const rawAnswers = await batchGenerateAnswers(questions);

        // 3. Prepare Data for DOCX with XML transformation
        const data = {
            name: formData.name,
            rollNo: formData.rollNo,
            program: formData.program,
            semester: formData.semester,
            class: formData.class,
            regulation: formData.regulation,
            courseTitle: formData.courseTitle,
            courseCode: formData.courseCode,
            aatNo: formData.aatNo,
        };

        for (let i = 0; i < 10; i++) {
            data[`question${i + 1}`] = questions[i];
            // Render answer as Raw XML
            data[`answer${i + 1}`] = markdownToWordXML(rawAnswers[i] || "AI Generation Failed.");
        }

        // 4. Load Template and Generate DOCX
        const content = fs.readFileSync(path.join(__dirname, "assets", "ReportTemplate.docx"), "binary");
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
        });

        doc.render(data);

        const buf = doc.getZip().generate({
            type: "nodebuffer",
            compression: "DEFLATE",
        });

        const safeReportName = (formData.name || "Student").replace(/[^a-zA-Z0-9_\-]/g, "_");
        res.setHeader("Content-Disposition", `attachment; filename="${safeReportName}_Report.docx"`);
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        res.send(buf);

    } catch (error) {
        console.error("Report Generation Error:", error);
        res.status(500).send("Error generating report. " + error.message);
    }
});

// ─── Start Server ─────────────────────────────────────────────────────────────
app.listen(port, () => {
    console.log(`Server listening on port ${port}`);
});