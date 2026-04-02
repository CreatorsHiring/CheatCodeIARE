const dotenv = require("dotenv");
const express = require("express");
const { Redis } = require("@upstash/redis");
const path = require("path");
const session = require("express-session");
const { GoogleGenerativeAI, HarmCategory, HarmBlockThreshold } = require("@google/generative-ai");
const PptxGenJS = require("pptxgenjs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");




const { inject } = require("@vercel/analytics");

dotenv.config();
inject();

const redis = new Redis({
  url: process.env.UPSTASH_REDIS_REST_URL,
  token: process.env.UPSTASH_REDIS_REST_TOKEN
});

const app = express();


// ─── Middleware ───────────────────────────────────────────────────────────────
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(session({
    secret: process.env.SESSION_SECRET || "cheatcodeiare_secret_key",
    resave: false,
    saveUninitialized: false
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
    "gemini-2.5-flash-lite",         // 2. Best Value (₹8 per 1M tokens) - Highly recommended
    "gemini-3.1-flash-lite-preview",  // 3. Newest & Fastest (₹20 per 1M tokens) - Backup
    "gemini-2.5-flash"
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
            // Run the actual API call through the queue — max 6 concurrent Gemini requests
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
    - Generate exactly 8 index items and exactly 8 corresponding slides in the same order.
    - Each slide must have exactly 4-5 bullet points.
    - Every bullet point must be a complete, detailed sentence — not a fragment or keyword.
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
                bullet: { code: "2022" },
                fontSize: 14,
                color: TEXT_MAIN,
                paraSpaceAfter: 8,
                indentLevel: 0,
            }
        }));
        s.addText(bulletItems, {
            x: CONTENT_X, y: CONTENT_Y, w: CONTENT_W, h: CONTENT_H,
            valign: "top", lineSpacingMultiple: 1.35,
            indentLevel: 0,
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

// ─── Redis Queue Constants ────────────────────────────────────────────────────
const MAX_ACTIVE_JOBS = 4;
const JOB_TTL_SECONDS = 600; // 10 min — auto-cleanup if something crashes

// ─── Generate PPT → Queue-based with Redis ────────────────────────────────────
app.post("/generate-ppt", isAuthenticated, async (req, res) => {
    const { department, subject, problemStatement } = req.body;
    const userId = req.session.user.studentId;

    try {
        // Store form data in session so the worker can use it
        req.session.formData = { department, subject, problemStatement };

        // Check if user already has a job running or waiting
        const existingStatus = await redis.get(`ppt_result:${userId}`);
        if (existingStatus) {
            const existing = typeof existingStatus === "string"
                ? JSON.parse(existingStatus)
                : existingStatus;
            if (existing.status === "waiting" || existing.status === "generating") {
                return res.json({ queued: true, position: existing.position || 0 });
            }
        }

        // Save form data + user identity to Redis (session not reliable across Vercel instances)
        await redis.set(
            `ppt_form:${userId}`,
            JSON.stringify({
                department, subject, problemStatement,
                userName:  req.session.user.name,
                studentId: req.session.user.studentId
            }),
            { ex: JOB_TTL_SECONDS }
        );

        // Remove user from queue if already in it (re-submit scenario)
        await redis.lrem("ppt_queue", 0, userId);

        // Add to queue
        await redis.rpush("ppt_queue", userId);

        // Get their position (1-based)
        const queue = await redis.lrange("ppt_queue", 0, -1);
        const position = queue.indexOf(userId) + 1;

        // Save initial waiting status with TTL
        await redis.set(
            `ppt_result:${userId}`,
            JSON.stringify({ status: "waiting", position }),
            { ex: JOB_TTL_SECONDS }
        );

        // Try to start processing immediately (non-blocking)
        processNextInQueue().catch(err => console.error("[Queue] processNextInQueue error:", err));

        // Return queued status to frontend — frontend will poll /ppt-status/:userId
        return res.json({ queued: true, position });

    } catch (error) {
        console.error("PPT Queue Error:", error);
        return res.status(500).json({ error: "Failed to queue PPT generation: " + error.message });
    }
});

// ─── Queue Worker: process next job if slot available ─────────────────────────
async function processNextInQueue() {
    // Acquire lock — only one Vercel instance runs this at a time
    const lock = await redis.set("queue_lock", "1", { nx: true, ex: 15 });
    if (!lock) {
        console.log("[Queue] Lock held, skipping.");
        return;
    }

    try {
        // Read + clamp active counter
        let activeJobs = parseInt(await redis.get("active_jobs") || "0", 10);
        if (isNaN(activeJobs) || activeJobs < 0) { activeJobs = 0; await redis.set("active_jobs", "0"); }

        console.log(`[Queue] active=${activeJobs}/${MAX_ACTIVE_JOBS}`);

        if (activeJobs >= MAX_ACTIVE_JOBS) return; // finally releases lock

        const userId = await redis.lpop("ppt_queue");
        if (!userId) { console.log("[Queue] Queue empty."); return; }

        await redis.lpush("ppt_processing", userId);
        await redis.incr("active_jobs");
        await redis.set(
            `ppt_result:${userId}`,
            JSON.stringify({ status: "generating", startedAt: Date.now() }),
            { ex: JOB_TTL_SECONDS }
        );

        // Update positions for remaining waiters
        const remaining = await redis.lrange("ppt_queue", 0, -1);
        for (let i = 0; i < remaining.length; i++) {
            await redis.set(
                `ppt_result:${remaining[i]}`,
                JSON.stringify({ status: "waiting", position: i + 1 }),
                { ex: JOB_TTL_SECONDS }
            );
        }

        console.log(`[Queue] Dispatching ${userId} | active=${activeJobs + 1}`);
        runPptJob(userId).catch(err => console.error(`[Queue] runPptJob error ${userId}:`, err));

    } catch (err) {
        console.error("[Queue] processNextInQueue error:", err);
    } finally {
        // ALWAYS release — no guard variables
        await redis.del("queue_lock");
    }
}

// ─── Run the actual PPT generation job ────────────────────────────────────────
async function runPptJob(userId) {
    try {
        const active = await redis.get(`user_active:${userId}`);

        if(!active){
        console.log("User left, cancelling job");
        }
        // Retrieve form data stored during /generate-ppt — stored in Redis too for safety
        const formRaw = await redis.get(`ppt_form:${userId}`);
        const formData = formRaw
            ? (typeof formRaw === "string" ? JSON.parse(formRaw) : formRaw)
            : null;

        if (!formData) {
            throw new Error("Form data not found in Redis for user: " + userId);
        }

        const content = await generatePptContent(formData.problemStatement);

        const result = {
            status: "done",
            userName:  formData.userName  || "Student",
            studentId: formData.studentId || userId,
            pptData: {
                department:       formData.department,
                subject:          formData.subject,
                problemStatement: formData.problemStatement,
                content
            }
        };

        // Store result — frontend will read this on next poll
        await redis.set(`ppt_result:${userId}`, JSON.stringify(result), { ex: JOB_TTL_SECONDS });
        console.log(`[Queue] Job done for user: ${userId}`);

    } catch (err) {
        console.error(`[Queue] Generation failed for ${userId}:`, err.message);
        await redis.set(
            `ppt_result:${userId}`,
            JSON.stringify({ status: "error", message: err.message }),
            { ex: 60 }
        );
    } finally {

    await redis.decr("active_jobs");

    const active = parseInt(await redis.get("active_jobs") || "0");

    if (active < 0) {
        await redis.set("active_jobs", 0);
    }

    //START NEXT JOB
    processNextInQueue().catch(err =>
        console.error("[Queue] Failed to trigger next job:", err)
    );
    await redis.lrem("ppt_processing", 0, userId);
}
}

async function recoverStuckJobs() {
    // Throttle: only run once per 60s globally
    const throttle = await redis.set("recovery_lock", "1", { nx: true, ex: 60 });
    if (!throttle) return;

    try {
        const processing = await redis.lrange("ppt_processing", 0, -1);
        const STUCK_MS = 8 * 60 * 1000; // 8 min — Gemini never takes this long

        for (const userId of processing) {
            const resultRaw = await redis.get(`ppt_result:${userId}`);

            if (!resultRaw) {
                // Key expired — job never finished, free the slot
                console.warn(`[Recovery] Key expired for ${userId}, freeing slot`);
                await redis.lrem("ppt_processing", 0, userId);
                const n = await redis.decr("active_jobs");
                if (n < 0) await redis.set("active_jobs", "0");
                continue;
            }

            const result = typeof resultRaw === "string" ? JSON.parse(resultRaw) : resultRaw;

            if (result.status === "generating") {
                const age = Date.now() - (result.startedAt || 0);
                if (!result.startedAt || age > STUCK_MS) {
                    console.warn(`[Recovery] Stuck job for ${userId} (age=${Math.round(age/1000)}s), re-queuing`);
                    await redis.lrem("ppt_processing", 0, userId);
                    const n = await redis.decr("active_jobs");
                    if (n < 0) await redis.set("active_jobs", "0");
                    const pos = await redis.lpos("ppt_queue", userId);
                    if (pos === null) await redis.rpush("ppt_queue", userId);
                    await redis.set(
                        `ppt_result:${userId}`,
                        JSON.stringify({ status: "waiting", position: 99 }),
                        { ex: JOB_TTL_SECONDS }
                    );
                }
            }
        }
    } catch (err) {
        console.error("[Recovery] error:", err);
    }
}

async function recoverStuckReports() {
    const throttle = await redis.set("report_recovery_lock", "1", { nx: true, ex: 60 });
    if (!throttle) return;

    try {
        let freedAny = false; // <-- ADD THIS

        const processing = await redis.lrange("report_processing", 0, -1);
        const STUCK_MS = 8 * 60 * 1000; // 8 minutes

        for (const userId of processing) {
            const resultRaw = await redis.get(`report_result:${userId}`);

            if (!resultRaw) {
                console.warn(`[ReportRecovery] Missing result for ${userId}, freeing slot`);
                await redis.lrem("report_processing", 0, userId);

                const n = await redis.decr("report_active");
                if (n < 0) await redis.set("report_active", "0");

                const pos = await redis.lpos("report_queue", userId);
                if (pos === null) await redis.rpush("report_queue", userId);

                await redis.set(
                    `report_result:${userId}`,
                    JSON.stringify({ status: "waiting", position: 99 }),
                    { ex: REPORT_TTL }
                );

                freedAny = true; // <-- ADD THIS
                continue;
            }

            const result = typeof resultRaw === "string" ? JSON.parse(resultRaw) : resultRaw;

            if (result.status === "generating") {
                const age = Date.now() - (result.startedAt || 0);

                if (!result.startedAt || age > STUCK_MS) {
                    console.warn(`[ReportRecovery] Stuck report for ${userId}, re-queuing`);

                    await redis.lrem("report_processing", 0, userId);

                    const n = await redis.decr("report_active");
                    if (n < 0) await redis.set("report_active", "0");

                    const pos = await redis.lpos("report_queue", userId);
                    if (pos === null) await redis.rpush("report_queue", userId);

                    await redis.set(
                        `report_result:${userId}`,
                        JSON.stringify({ status: "waiting", position: 99 }),
                        { ex: REPORT_TTL }
                    );

                    freedAny = true; // <-- ADD THIS
                }
            }
        }

        await refreshReportQueuePositions();

        //  IMPORTANT: restart queue if slots were freed
        if (freedAny) {
            console.log("[ReportRecovery] Freed stuck slots, restarting queue...");
            processNextReport().catch(err =>
                console.error("[ReportRecovery] Failed to restart queue:", err)
            );
        }

    } catch (err) {
        console.error("[ReportRecovery] error:", err);
    }
}

// ─── Heartbeat: tells server user is still on queue page ──────────────────────
app.post("/heartbeat/:userId", async (req,res)=>{

    const { userId } = req.params;

    try {

        // mark user as active for 15 seconds
        await redis.set(`user_active:${userId}`, "1", { ex: 60 });

        res.sendStatus(200);

    } catch(err){

        console.error("[Heartbeat] error:", err);
        res.sendStatus(500);

    }

});

// ─── Status Polling Endpoint ──────────────────────────────────────────────────
app.get("/ppt-status/:userId",  async (req, res) => {
    const { userId } = req.params;
    // Security: users can only check their own status
    // if (req.session.user.studentId !== userId) {
    //     return res.status(403).json({ error: "Forbidden" });
    // }

    try {
        // Throttled recovery — max once per 60s, not on every 3s poll
        recoverStuckJobs().catch(console.error);

        const raw = await redis.get(`ppt_result:${userId}`);
        if (!raw) return res.json({ status: "not_found" });

        const data = typeof raw === "string" ? JSON.parse(raw) : raw;

        if (data.status === "done") {
            // Populate session so /edit and /download work on this instance
            req.session.pptData = data.pptData;
            req.session.user    = {
                name:      data.userName  || req.session.user?.name || "Student",
                studentId: data.studentId || userId
            };
            // DO NOT delete ppt_result here — /edit/:userId and /download-ppt still need it
            return res.json({ status: "done" });
        }

        return res.json(data);

    } catch (err) {
        console.error("[Status] Error:", err);
        return res.status(500).json({ error: "Failed to get status." });
    }
});

// ─── Queue Waiting Page ──────────────────────────────────────────────────────
app.get("/queue", isAuthenticated, (req, res) => {
    res.render("queue", { user: req.session.user });
});

// ─── Edit Page ────────────────────────────────────────────────────────────────
app.get("/edit/:userId", async (req, res) => {

    const { userId } = req.params;

    const raw = await redis.get(`ppt_result:${userId}`);

    if (!raw) {
        return res.redirect("/ppt");
    }

    const data = typeof raw === "string" ? JSON.parse(raw) : raw;

    if (data.status !== "done") {
        return res.redirect("/ppt");
    }

    const realUser = {
        name:      data.userName  || req.session.user?.name || "Student",
        studentId: data.studentId || userId
    };
    req.session.user    = realUser;
    req.session.pptData = data.pptData;

    res.render("edit", { user: realUser, pptData: data.pptData });
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
app.post("/download-ppt", async (req, res) => {
    try {
        const { userId } = req.body;
        if (!userId) return res.status(400).send("Missing userId.");

        // Primary: read from session (fastest, no Redis round-trip)
        let pptData  = req.session.pptData;
        let userName = req.session.user?.name || null;

        // Fallback: read from Redis (different Vercel instance, session empty)
        if (!pptData) {
            console.log(`[Download] Session empty for ${userId}, reading Redis...`);
            const raw = await redis.get(`ppt_result:${userId}`);
            if (!raw) return res.status(404).send("Presentation not found. Please generate again.");
            const data = typeof raw === "string" ? JSON.parse(raw) : raw;
            if (data.status !== "done") return res.status(400).send("Presentation not ready yet.");
            pptData  = data.pptData;
            userName = data.userName || "Student";
            // Restore session so future requests on this instance are fast
            req.session.pptData = pptData;
            req.session.user    = { name: userName, studentId: data.studentId || userId };
        }

        if (!userName) userName = "Student";

        const buffer = await buildPptxBuffer(
            pptData.content,
            {
                department:       pptData.department,
                subject:          pptData.subject,
                problemStatement: pptData.problemStatement
            },
            { name: userName, studentId: userId }
        );

        const safeName = userName.replace(/[^a-zA-Z0-9_\-]/g, "_");
        res.setHeader("Content-Disposition", `attachment; filename="${safeName}_AAT.pptx"`);
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
        res.send(buffer);

    } catch (err) {
        console.error("[Download] Error:", err);
        res.status(500).send("Download failed: " + err.message);

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

// ─── Complex Engineering Problems ────────────────────────────────────────────
app.get("/complex-engineering", isAuthenticated, (req, res) => {
    res.render("complex-engineering", { user: req.session.user });
});

// ─── (Keep existing report generation logic below) ────────────────────────────

// ─── Escape raw text for Word XML ────────────────────────────────────────────
function escXml(str) {
    return String(str || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
}

// ─── Symbol substitution — replaces text codes with proper Unicode ───────────
function applySymbols(text) {
    return text
        .replace(/\bohm\b/gi,      '\u03A9')   // Ω
        .replace(/\bOhm\b/g,       '\u03A9')
        .replace(/Omega/g,         '\u03A9')
        .replace(/\bpi\b/gi,       '\u03C0')   // π
        .replace(/\balpha\b/gi,    '\u03B1')   // α
        .replace(/\bbeta\b/gi,     '\u03B2')   // β
        .replace(/\bgamma\b/gi,    '\u03B3')   // γ
        .replace(/\bdelta\b/gi,    '\u03B4')   // δ
        .replace(/\bDelta\b/g,     '\u0394')   // Δ
        .replace(/\bmu\b/gi,       '\u03BC')   // μ
        .replace(/\btheta\b/gi,    '\u03B8')   // θ
        .replace(/\bomega\b/gi,    '\u03C9')   // ω (lowercase)
        .replace(/\bsigma\b/gi,    '\u03C3')   // σ
        .replace(/\bphi\b/gi,      '\u03C6')   // φ
        .replace(/\blambda\b/gi,   '\u03BB')   // λ
        .replace(/\betaeta\b/gi,   '\u03B7')   // η
        .replace(/\^2\b/g,         '\u00B2')   // ²
        .replace(/\^3\b/g,         '\u00B3')   // ³
        .replace(/>=|=>|&gt;=/g,   '\u2265')   // ≥
        .replace(/<=|=<|&lt;=/g,   '\u2264')   // ≤
        .replace(/ -> /g,          ' \u2192 ') // →
        .replace(/sqrt\(([^)]+)\)/g, '\u221A($1)') // √(x)
        .replace(/\+-/g,           '\u00B1')   // ±
        .replace(/\binfinity\b/gi, '\u221E')   // ∞
        .replace(/\bdeg\b/gi,      '\u00B0')   // °
        .replace(/\bmu_0\b/gi,     '\u03BC\u2080') // μ₀
        .replace(/\bepsilon\b/gi,  '\u03B5');  // ε
}

// ─── Build a Word XML paragraph ───────────────────────────────────────────────
// opts: { bold, italic, bullet, spaceAfter, spaceBefore, fontSize, indent, centered }
function makeWordPara(text, opts = {}) {
    const {
        bold = false, italic = false, bullet = false,
        spaceAfter = 80, spaceBefore = 0, fontSize = 22,
        indent = 0, centered = false
    } = opts;

    const safeText = applySymbols(escXml(text));

    let pPr = `<w:pPr><w:spacing w:before="${spaceBefore}" w:after="${spaceAfter}"/>`;
    if (centered) pPr += `<w:jc w:val="center"/>`;
    if (indent > 0) pPr += `<w:ind w:left="${indent}"/>`;
    if (bullet) pPr += `<w:ind w:left="720" w:hanging="360"/>`;
    pPr += `</w:pPr>`;

    let rPr = `<w:rPr><w:sz w:val="${fontSize}"/><w:szCs w:val="${fontSize}"/>`;
    if (bold)   rPr += `<w:b/><w:bCs/>`;
    if (italic) rPr += `<w:i/><w:iCs/>`;
    rPr += `</w:rPr>`;

    const bulletChar = bullet
        ? `<w:r><w:rPr><w:sz w:val="${fontSize}"/><w:szCs w:val="${fontSize}"/></w:rPr><w:t xml:space="preserve">• </w:t></w:r>`
        : "";

    return `<w:p>${pPr}${bulletChar}<w:r>${rPr}<w:t xml:space="preserve">${safeText}</w:t></w:r></w:p>`;
}

// ─── Build a bordered Word table ──────────────────────────────────────────────
function makeWordTable(rows) {
    const borderProps = `
        <w:top    w:val="single" w:sz="6" w:space="0" w:color="000000"/>
        <w:left   w:val="single" w:sz="6" w:space="0" w:color="000000"/>
        <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>
        <w:right  w:val="single" w:sz="6" w:space="0" w:color="000000"/>
        <w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/>
        <w:insideV w:val="single" w:sz="6" w:space="0" w:color="000000"/>`;

    let tbl = `<w:tbl><w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
        <w:tblW w:w="0" w:type="auto"/>
        <w:tblBorders>${borderProps}</w:tblBorders>
        <w:tblCellMar>
            <w:top    w:w="100" w:type="dxa"/>
            <w:left   w:w="144" w:type="dxa"/>
            <w:bottom w:w="100" w:type="dxa"/>
            <w:right  w:w="144" w:type="dxa"/>
        </w:tblCellMar>
    </w:tblPr>`;

    rows.forEach((row, rowIdx) => {
        const cells    = row.split("|").map(s => s.trim()).filter(s => s.length > 0);
        const isHeader = rowIdx === 0;
        tbl += `<w:tr>`;
        cells.forEach(cell => {
            const cellText = applySymbols(escXml(cell));
            const rPr = isHeader
                ? `<w:rPr><w:b/><w:bCs/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>`
                : `<w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>`;
            const shading = isHeader
                ? `<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>`
                : "";
            tbl += `<w:tc>
                <w:tcPr><w:tcW w:w="0" w:type="auto"/>${shading}</w:tcPr>
                <w:p><w:pPr><w:spacing w:after="0"/></w:pPr>
                <w:r>${rPr}<w:t xml:space="preserve">${cellText}</w:t></w:r></w:p>
            </w:tc>`;
        });
        tbl += `</w:tr>`;
    });

    tbl += `</w:tbl><w:p><w:pPr><w:spacing w:after="120"/></w:pPr></w:p>`;
    return tbl;
}

// ─── Main converter: structured answer text → Word XML ───────────────────────
//
// Line types recognised:
//   HEADING: Title         → bold, uppercase, larger font, space above
//   BULLET: sentence       → bullet point, indented
//   GIVEN: V = 200 V, ...  → bold label "Given:" + indented italic values
//   STEP: 1. Find current  → bold step label, space above
//   CALC: Ia = IL - Ish    → centered italic, indented — the actual calculation line
//   FORMULA: F = ma        → italic, indented
//   RESULT: Ta = 230.5 Nm  → bold italic, indented — final boxed result line
//   TABLE: / END_TABLE     → bordered table
//   (plain text)           → normal paragraph
//
const markdownToWordXML = (text) => {
    if (!text) return "";

    const lines   = text.split("\n");
    let finalXml  = "";
    let tableRows = [];
    let inTable   = false;

    for (const rawLine of lines) {
        const line = rawLine.trim();
        if (!line) continue;

        const UP = line.toUpperCase();

        // ── Table handling ────────────────────────────────────────────
        if (UP === "TABLE:") {
            inTable = true; tableRows = []; continue;
        }
        if (inTable) {
            if (UP === "END_TABLE") {
                inTable = false;
                finalXml += makeWordTable(tableRows);
                tableRows = [];
            } else if (line.includes("|")) {
                tableRows.push(line);
            }
            continue;
        }

        // ── Structured line types ─────────────────────────────────────
        if (UP.startsWith("HEADING:")) {
            const t = line.slice(8).trim().toUpperCase();
            // Extra space before heading so sections breathe
            finalXml += makeWordPara(t, { bold: true, spaceBefore: 160, spaceAfter: 60, fontSize: 24 });

        } else if (UP.startsWith("BULLET:")) {
            const t = line.slice(7).trim();
            finalXml += makeWordPara(t, { bullet: true, spaceAfter: 60, fontSize: 21 });

        } else if (UP.startsWith("GIVEN:")) {
            // "Given:" label bold, then the values on the same para — indented
            const values = line.slice(6).trim();
            // Bold "Given:" prefix
            const safeV  = applySymbols(escXml(values));
            finalXml += `<w:p>
                <w:pPr><w:spacing w:before="120" w:after="60"/><w:ind w:left="360"/></w:pPr>
                <w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
                    <w:t xml:space="preserve">Given:  </w:t></w:r>
                <w:r><w:rPr><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
                    <w:t xml:space="preserve">${safeV}</w:t></w:r>
            </w:p>`;

        } else if (UP.startsWith("STEP:")) {
            // Step label — bold, space above so each step is clearly separated
            const t = line.slice(5).trim();
            finalXml += makeWordPara(t, { bold: true, spaceBefore: 160, spaceAfter: 40, fontSize: 22 });

        } else if (UP.startsWith("CALC:")) {
            // The actual equation — italic, centered, indented, space below
            const t = line.slice(5).trim();
            finalXml += makeWordPara(t, { italic: true, centered: true, spaceAfter: 80, fontSize: 22 });

        } else if (UP.startsWith("FORMULA:")) {
            // Named formula reference — italic, indented
            const t = line.slice(8).trim();
            finalXml += makeWordPara(t, { italic: true, indent: 720, spaceAfter: 80, fontSize: 22 });

        } else if (UP.startsWith("RESULT:")) {
            // Final answer — bold italic, indented, space above+below
            const t = line.slice(7).trim();
            finalXml += makeWordPara(t, { bold: true, italic: true, indent: 720, spaceBefore: 80, spaceAfter: 120, fontSize: 22 });

        } else {
            // Plain paragraph
            finalXml += makeWordPara(line, { spaceAfter: 80, fontSize: 22 });
        }
    }

    if (inTable && tableRows.length) finalXml += makeWordTable(tableRows);

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
You are an expert academic content generator writing detailed answers for a university AAT (Alternative Assessment Tool) report.
Generate thorough, exam-ready answers for ALL questions below.

━━━ STRICT LINE TYPES — use ONLY these, no Markdown, no HTML, no asterisks ━━━

  HEADING: <TOPIC TITLE IN UPPERCASE>
  BULLET: <one complete informative sentence>
  GIVEN: <list all known values on one line, e.g. V = 200 V, Ra = 0.1 Ohm, N = 750 rpm>
  STEP: <Step N: description of what this step calculates>
  CALC: <the actual equation and its evaluated result, e.g. Ia = IL - Ish = 100 - 5 = 95 A>
  RESULT: <final answer with units, e.g. Torque Ta = 230.43 Nm>
  FORMULA: <named formula used, e.g. Ta = (Eb x Ia) / (2 x pi x N / 60)>
  TABLE:
  Col1 | Col2 | Col3
  Row1A | Row1B | Row1C
  END_TABLE

━━━ SYMBOL RULES ━━━
- Write Ohm for Ω (the renderer will convert it automatically)
- Write pi for π, alpha for α, beta for β, omega for ω, theta for θ
- Write ^2 for ², ^3 for ³, +- for ±, sqrt(x) for √x, deg for °

━━━ CONTENT RULES ━━━
1. Every answer MUST start with a HEADING: line.
2. Theory/concept questions: AT LEAST 5 BULLET: lines of detailed sentences. No STEP/CALC needed.
3. Calculation questions: use GIVEN: → STEP: + CALC: pattern (one STEP per logical step). Include RESULT: at the end of each sub-part. No BULLET lines needed for pure calc questions.
4. Mixed questions (theory + calculation): use BULLET for theory, then GIVEN/STEP/CALC/RESULT for the calculation part under a sub-HEADING.
5. Comparison/difference questions: MUST include a TABLE (Feature | Type A | Type B).
6. Every formula used in a calculation MUST appear as a FORMULA: line before or after the CALC line.
7. Multiple HEADING: lines allowed to separate sub-topics or sub-parts (i, ii, iii).
8. Do NOT use ALL CAPS in BULLET, CALC, STEP, or RESULT lines — only HEADING text is uppercase.
9. Aim for 180-220 words per answer.

━━━ EXAMPLE — CALCULATION QUESTION ━━━
HEADING: DC SHUNT MOTOR — TORQUE AND COPPER LOSSES
GIVEN: V = 200 V, IL = 100 A, N = 750 rpm, Ra = 0.1 Ohm, Rsh = 40 Ohm, Pmech = 1500 W
HEADING: (i) TORQUE DEVELOPED BY ARMATURE
STEP: Step 1: Find shunt field current
FORMULA: Ish = V / Rsh
CALC: Ish = 200 / 40 = 5 A
STEP: Step 2: Find armature current
FORMULA: Ia = IL - Ish
CALC: Ia = 100 - 5 = 95 A
STEP: Step 3: Find back EMF
FORMULA: Eb = V - (Ia x Ra)
CALC: Eb = 200 - (95 x 0.1) = 200 - 9.5 = 190.5 V
STEP: Step 4: Find power developed
FORMULA: Pdev = Eb x Ia
CALC: Pdev = 190.5 x 95 = 18097.5 W
STEP: Step 5: Find torque
FORMULA: Ta = Pdev / (2 x pi x N / 60)
CALC: Ta = 18097.5 / (2 x pi x 750 / 60) = 18097.5 / 78.54 = 230.43 Nm
RESULT: Torque developed Ta = 230.43 Nm
HEADING: (ii) COPPER LOSSES
STEP: Step 1: Armature copper loss
FORMULA: Pcu_a = Ia^2 x Ra
CALC: Pcu_a = 95^2 x 0.1 = 9025 x 0.1 = 902.5 W
STEP: Step 2: Shunt field copper loss
FORMULA: Pcu_sh = V x Ish
CALC: Pcu_sh = 200 x 5 = 1000 W
STEP: Step 3: Total copper loss
FORMULA: Pcu = Pcu_a + Pcu_sh
CALC: Pcu = 902.5 + 1000 = 1902.5 W
RESULT: Total Copper Loss Pcu = 1902.5 W

━━━ EXAMPLE — THEORY QUESTION ━━━
HEADING: TRANSFORMER — WORKING PRINCIPLE
A transformer transfers electrical energy between circuits using electromagnetic induction.
BULLET: It works on the principle of mutual induction between two coils wound on a common magnetic core.
BULLET: The primary winding receives AC supply and produces a time-varying magnetic flux.
BULLET: This alternating flux links the secondary winding and induces an EMF by Faraday's law.
FORMULA: E = -N x (dPhi/dt)
BULLET: The turns ratio determines whether the transformer steps up or steps down the voltage.
BULLET: Ideal transformers have no copper or core losses and 100% efficiency.
BULLET: Practical applications include power transmission, isolation circuits, and impedance matching.

━━━ JSON RESPONSE FORMAT ━━━
Return ONLY a raw JSON array — no markdown fences, no extra text:
[
  { "index": 0, "answer": "full formatted answer for Q1 here" },
  { "index": 1, "answer": "full formatted answer for Q2 here" }
]
"index" starts at 0 and matches question order exactly.

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

// ─── Report Queue Constants ──────────────────────────────────────────────────
const MAX_REPORT_JOBS  = 4;
const REPORT_TTL       = 600; // 10 min

async function refreshReportQueuePositions() {
    const queue = await redis.lrange("report_queue", 0, -1);

    for (let i = 0; i < queue.length; i++) {
        await redis.set(
            `report_result:${queue[i]}`,
            JSON.stringify({ status: "waiting", position: i + 1 }),
            { ex: REPORT_TTL }
        );
    }
}

// ─── Enqueue Report ───────────────────────────────────────────────────────────
app.post("/generate-report", isAuthenticated, async (req, res) => {
    const userId = req.session.user.studentId;

    try {
        // Don't double-queue if already waiting/generating
        const existingRaw = await redis.get(`report_result:${userId}`);
        if (existingRaw) {
            const existing = typeof existingRaw === "string" ? JSON.parse(existingRaw) : existingRaw;
            if (existing.status === "waiting" || existing.status === "generating") {
                return res.json({ queued: true, position: existing.position || 1 });
            }
        }

        // Persist entire form to Redis — session unreliable across Vercel instances
        const formData = req.body;
        await redis.set(
            `report_form:${userId}`,
            JSON.stringify(formData),
            { ex: REPORT_TTL }
        );

        // Add to queue
        await redis.lrem("report_queue", 0, userId);
        await redis.rpush("report_queue", userId);

        const queue    = await redis.lrange("report_queue", 0, -1);
        const position = queue.indexOf(userId) + 1;

        await redis.set(
            `report_result:${userId}`,
            JSON.stringify({ status: "waiting", position }),
            { ex: REPORT_TTL }
        );

        // Kick the worker
        processNextReport().catch(err => console.error("[ReportQueue] boot error:", err));

        return res.json({ queued: true, position });

    } catch (err) {
        console.error("[ReportQueue] Enqueue error:", err);
        return res.status(500).json({ error: "Failed to queue report: " + err.message });
    }
});

// ─── Report Queue Worker ──────────────────────────────────────────────────────
async function processNextReport() {
    const lock = await redis.set("report_lock", "1", { nx: true, ex: 15 });
    if (!lock) { console.log("[ReportQueue] Lock held, skipping."); return; }

    try {
        let active = parseInt(await redis.get("report_active") || "0", 10);
        if (isNaN(active) || active < 0) { active = 0; await redis.set("report_active", "0"); }

        console.log(`[ReportQueue] active=${active}/${MAX_REPORT_JOBS}`);
        if (active >= MAX_REPORT_JOBS) return;

        const userId = await redis.lpop("report_queue");
        if (!userId) { console.log("[ReportQueue] Queue empty."); return; }

        await redis.lpush("report_processing", userId);
        await redis.incr("report_active");
        await redis.set(
            `report_result:${userId}`,
            JSON.stringify({ status: "generating", startedAt: Date.now() }),
            { ex: REPORT_TTL }
        );

        // Update positions for remaining waiters
        await refreshReportQueuePositions();

        console.log(`[ReportQueue] Dispatching ${userId} | active=${active + 1}`);
        runReportJob(userId).catch(err => console.error(`[ReportQueue] job error ${userId}:`, err));

    } catch (err) {
        console.error("[ReportQueue] processNextReport error:", err);
    } finally {
        await redis.del("report_lock");
    }
}

// ─── Run one report generation job ───────────────────────────────────────────
async function runReportJob(userId) {
    try {
        const formRaw = await redis.get(`report_form:${userId}`);
        if (!formRaw) throw new Error("Form data missing for user: " + userId);
        const formData = typeof formRaw === "string" ? JSON.parse(formRaw) : formRaw;

        const questions = [];
        for (let i = 1; i <= 10; i++) questions.push(formData[`question${i}`] || "");

        const rawAnswers = await batchGenerateAnswers(questions);

        const docData = {
            name:        formData.name,
            rollNo:      formData.rollNo,
            program:     formData.program,
            semester:    formData.semester,
            class:       formData.class,
            regulation:  formData.regulation,
            courseTitle: formData.courseTitle,
            courseCode:  formData.courseCode,
            aatNo:       formData.aatNo,
        };
        for (let i = 0; i < 10; i++) {
            docData[`question${i + 1}`] = questions[i];
            docData[`answer${i + 1}`]   = markdownToWordXML(rawAnswers[i] || "AI Generation Failed.");
        }

        const content = fs.readFileSync(path.join(__dirname, "assets", "ReportTemplate.docx"), "binary");
        const zip = new PizZip(content);
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
        doc.render(docData);
        const buf = doc.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });

        // Store as base64 — Redis only holds strings
        await redis.set(
            `report_result:${userId}`,
            JSON.stringify({
                status:   "done",
                fileName: (formData.name || "Student").replace(/[^a-zA-Z0-9_]/g, "_") + "_Report.docx",
                docBase64: buf.toString("base64")
            }),
            { ex: REPORT_TTL }
        );

        console.log(`[ReportQueue] Job done for ${userId}`);

    } catch (err) {
        console.error(`[ReportQueue] Failed for ${userId}:`, err.message);
        await redis.set(
            `report_result:${userId}`,
            JSON.stringify({ status: "error", message: err.message }),
            { ex: 120 }
        );
    } finally {
        const n = await redis.decr("report_active");
        if (n < 0) await redis.set("report_active", "0");

        await redis.lrem("report_processing", 0, userId);

        processNextReport().catch(err =>
            console.error("[ReportQueue] Failed to trigger next job:", err)
        );
    }
}

// ─── Report Status Polling ────────────────────────────────────────────────────
app.get("/report-status/:userId", async (req, res) => {
    const { userId } = req.params;
    try {
        recoverStuckReports().catch(console.error);
        const raw = await redis.get(`report_result:${userId}`);
        if (!raw) return res.json({ status: "not_found" });
        const data = typeof raw === "string" ? JSON.parse(raw) : raw;
        // Don't send the base64 blob in the status poll — only send metadata
        if (data.status === "done") {
            return res.json({ status: "done", fileName: data.fileName });
        }
        return res.json(data);
    } catch (err) {
        console.error("[ReportStatus] error:", err);
        return res.status(500).json({ error: "Status check failed." });
    }
});

// ─── Report Download ──────────────────────────────────────────────────────────
app.get("/download-report/:userId", async (req, res) => {
    const { userId } = req.params;
    try {
        const raw = await redis.get(`report_result:${userId}`);
        if (!raw) return res.status(404).send("Report not found. Please generate again.");

        const data = typeof raw === "string" ? JSON.parse(raw) : raw;
        if (data.status !== "done") return res.status(400).send("Report not ready yet.");

        const buf      = Buffer.from(data.docBase64, "base64");
        const fileName = data.fileName || `${userId}_Report.docx`;

        res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        res.send(buf);

        // Clean up Redis after successful download
        await redis.del(`report_result:${userId}`);
        await redis.del(`report_form:${userId}`);

    } catch (err) {
        console.error("[ReportDownload] error:", err);
        res.status(500).send("Download failed: " + err.message);
    }
});

// ─── Report Queue Waiting Page ────────────────────────────────────────────────
app.get("/report-queue", isAuthenticated, (req, res) => {
    res.render("report-queue", { user: req.session.user });
});

// /report-done route removed — download is automatic, no done page needed

// ══════════════════════════════════════════════════════════════════════════════
// COMPLEX ENGINEERING PROBLEMS — Queue-based generation with Docxtemplater
// ══════════════════════════════════════════════════════════════════════════════

const MAX_CEP_JOBS = 4;
const CEP_TTL      = 600; // 10 min

// ─── Enqueue CEP ─────────────────────────────────────────────────────────────
app.post("/generate-cep", isAuthenticated, async (req, res) => {
    const userId = req.session.user.studentId;
    try {
        // Don't double-queue
        const existingRaw = await redis.get(`cep_result:${userId}`);
        if (existingRaw) {
            const existing = typeof existingRaw === "string" ? JSON.parse(existingRaw) : existingRaw;
            if (existing.status === "waiting" || existing.status === "generating") {
                return res.json({ queued: true, position: existing.position || 1 });
            }
        }

        // Save form data to Redis
        await redis.set(`cep_form:${userId}`, JSON.stringify(req.body), { ex: CEP_TTL });

        await redis.lrem("cep_queue", 0, userId);
        await redis.rpush("cep_queue", userId);

        const queue    = await redis.lrange("cep_queue", 0, -1);
        const position = queue.indexOf(userId) + 1;

        await redis.set(
            `cep_result:${userId}`,
            JSON.stringify({ status: "waiting", position }),
            { ex: CEP_TTL }
        );

        processNextCEP().catch(err => console.error("[CEP] boot error:", err));
        return res.json({ queued: true, position });

    } catch (err) {
        console.error("[CEP] Enqueue error:", err);
        return res.status(500).json({ error: "Failed to queue: " + err.message });
    }
});

// ─── CEP Queue Worker ─────────────────────────────────────────────────────────
async function processNextCEP() {
    const lock = await redis.set("cep_lock", "1", { nx: true, ex: 15 });
    if (!lock) return;

    try {
        let active = parseInt(await redis.get("cep_active") || "0", 10);
        if (isNaN(active) || active < 0) { active = 0; await redis.set("cep_active", "0"); }

        if (active >= MAX_CEP_JOBS) return;

        const userId = await redis.lpop("cep_queue");
        if (!userId) return;

        await redis.lpush("cep_processing", userId);
        await redis.incr("cep_active");
        await redis.set(
            `cep_result:${userId}`,
            JSON.stringify({ status: "generating", startedAt: Date.now() }),
            { ex: CEP_TTL }
        );

        const remaining = await redis.lrange("cep_queue", 0, -1);
        for (let i = 0; i < remaining.length; i++) {
            await redis.set(
                `cep_result:${remaining[i]}`,
                JSON.stringify({ status: "waiting", position: i + 1 }),
                { ex: CEP_TTL }
            );
        }

        console.log(`[CEP] Dispatching ${userId} | active=${active + 1}`);
        runCEPJob(userId).catch(err => console.error(`[CEP] job error ${userId}:`, err));

    } catch (err) {
        console.error("[CEP] processNextCEP error:", err);
    } finally {
        await redis.del("cep_lock");
    }
}

// ─── Run one CEP generation job ───────────────────────────────────────────────
async function runCEPJob(userId) {
    try {
        const formRaw = await redis.get(`cep_form:${userId}`);
        if (!formRaw) throw new Error("Form data missing for user: " + userId);
        const form = typeof formRaw === "string" ? JSON.parse(formRaw) : formRaw;

        const ps = form.problemStatement || "";

        // Format date nicely
        const rawDate  = form.date || new Date().toISOString().split("T")[0];
        const dateObj  = new Date(rawDate);
        const formattedDate = dateObj.toLocaleDateString("en-IN", { day: "2-digit", month: "long", year: "numeric" });

        // Single Gemini call — generate all sections as structured JSON
        const prompt = `
You are an expert engineering academic content generator.
Analyse the following Complex Engineering Problem Statement and generate detailed content for a university project report.

Problem Statement: "${ps}"

Return ONLY a valid JSON object with these exact keys — no markdown fences, no extra text:

{
  "abstract": "An 8-line detailed summary of the problem, approach, and expected outcomes. Should read as a formal academic abstract.",
  "introduction": "A 5-6 sentence introduction to the topic, its background, relevance, and importance in engineering.",
  "overview": "A 4-5 sentence overview of the project scope, what it covers, and how it approaches the problem.",
  "objectives": "List 5 specific objectives of this project as a numbered list (1. ... 2. ... 3. ... 4. ... 5. ...). Each objective should be one clear sentence.",
  "prerequisites": "List exactly 4 prerequisite knowledge areas needed for this project (1. ... 2. ... 3. ... 4. ...). Each should name the subject and briefly explain why it is needed.",
  "requirements": "List 5-6 technical and resource requirements for this project (hardware, software, datasets, tools, or skills). Format as a numbered list.",
  "methodology": "Describe in 5-6 sentences the methodology and approach that will be used to solve this problem. Be specific to the domain.",
  "workflow": "List exactly 5 workflow steps as: Step 1: ... Step 2: ... Step 3: ... Step 4: ... Step 5: ... Each step should be one sentence describing what is done at that stage.",
  "content": "A detailed explanation (8-10 sentences) of the technical solution corresponding to the workflow steps. Include relevant techniques, algorithms, formulas, or design approaches specific to this problem.",
  "result": "A 3-4 sentence description of the expected results and outcomes of solving this problem.",
  "conclusion": "A 4-5 sentence conclusion summarising what was achieved, its significance, and what it demonstrates.",
  "futureScope": "List 4-5 future scope items as a numbered list describing how this project can be extended or improved."
}

Rules:
- Be specific and technical — this is for a university engineering report.
- Do NOT use markdown formatting inside the JSON values.
- Do NOT include bullet characters (•) or asterisks (*) in values.
- Use plain numbered lists (1. 2. 3.) where lists are needed.
- All values must be plain text strings.
`;

        const aiResult = await generateWithFallback(m => m.generateContent(prompt), true);
        let aiText = aiResult.response.text()
            .replace(/\`\`\`json/g, "").replace(/\`\`\`/g, "").trim();
        const aiData = JSON.parse(aiText);

        // Build Docxtemplater data — every key matches a {placeholder} in the template exactly
        // Template placeholders confirmed: {name} {rollNumber} {program} {semester}
        // {branch} {class} {subject} {courseCode} {couseCode} {date} {topic}
        // {lecturerName} {hodName}
        // {abstract} {introduction} {overview} {objectives} {prerequisites}
        // {requirements} {methodology} {workflow} {content} {results} {conclusion} {futureScope}
        const docData = {
            // Student details
            name:         form.name         || "",
            rollNumber:   form.rollNo        || "",
            program:      form.program       || "",
            semester:     form.semester      || "",
            branch:       form.class         || "",  // {branch} — used in paragraphs
            class:        form.class         || "",  // {class}  — used in tables
            subject:      form.courseTitle   || "",  // {subject} — used in tables
            courseName:   form.courseTitle   || "",  // {courseName} — fallback
            courseCode:   form.courseCode    || "",  // {courseCode}
            couseCode:    form.courseCode    || "",  // {couseCode} — template typo, keep both
            date:         formattedDate,
            topic:        ps,
            // Faculty
            lecturerName: form.lecturerName  || "",
            hodName:      form.hodName       || "",
            // AI-generated content
            abstract:     aiData.abstract       || "",
            introduction: aiData.introduction   || "",
            overview:     aiData.overview       || "",
            objectives:   aiData.objectives     || "",
            prerequisites:aiData.prerequisites  || "",
            requirements: aiData.requirements   || "",
            methodology:  aiData.methodology    || "",
            workflow:     aiData.workflow       || "",
            content:      aiData.content        || "",
            results:      aiData.result         || "",  // template uses {results} (plural)
            result:       aiData.result         || "",  // keep singular too for safety
            conclusion:   aiData.conclusion     || "",
            futureScope:  aiData.futureScope    || "",
        };

        // Fill template
        const templateBuf = fs.readFileSync(
            path.join(__dirname, "assets", "ComplexEngineeringTemplate.docx"),
            "binary"
        );
        const zip = new PizZip(templateBuf);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks:    true,
        });
        doc.render(docData);
        const buf = doc.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });

        const safeName = (form.name || "Student").replace(/[^a-zA-Z0-9_]/g, "_");

        await redis.set(
            `cep_result:${userId}`,
            JSON.stringify({
                status:    "done",
                fileName:  `${safeName}_CEP.docx`,
                docBase64: buf.toString("base64")
            }),
            { ex: CEP_TTL }
        );

        console.log(`[CEP] Done for ${userId}`);

    } catch (err) {
        console.error(`[CEP] Failed for ${userId}:`, err.message);
        await redis.set(
            `cep_result:${userId}`,
            JSON.stringify({ status: "error", message: err.message }),
            { ex: 120 }
        );
    } finally {
        const n = await redis.decr("cep_active");
        if (n < 0) await redis.set("cep_active", "0");
        await redis.lrem("cep_processing", 0, userId);
        processNextCEP().catch(err => console.error("[CEP] next-job error:", err));
    }
}

// ─── CEP Status Polling ───────────────────────────────────────────────────────
app.get("/cep-status/:userId", async (req, res) => {
    const { userId } = req.params;
    try {
        const raw = await redis.get(`cep_result:${userId}`);
        if (!raw) return res.json({ status: "not_found" });
        const data = typeof raw === "string" ? JSON.parse(raw) : raw;
        if (data.status === "done") return res.json({ status: "done", fileName: data.fileName });
        return res.json(data);
    } catch (err) {
        return res.status(500).json({ error: "Status check failed." });
    }
});

// ─── CEP Download ─────────────────────────────────────────────────────────────
app.get("/download-cep/:userId", async (req, res) => {
    const { userId } = req.params;
    try {
        const raw = await redis.get(`cep_result:${userId}`);
        if (!raw) return res.status(404).send("Document not found. Please generate again.");
        const data = typeof raw === "string" ? JSON.parse(raw) : raw;
        if (data.status !== "done") return res.status(400).send("Document not ready yet.");

        const buf      = Buffer.from(data.docBase64, "base64");
        const fileName = data.fileName || `${userId}_CEP.docx`;

        res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        res.send(buf);

        await redis.del(`cep_result:${userId}`);
        await redis.del(`cep_form:${userId}`);
    } catch (err) {
        console.error("[CEP] Download error:", err);
        res.status(500).send("Download failed: " + err.message);
    }
});

// ─── CEP Queue Waiting Page ───────────────────────────────────────────────────
app.get("/cep-queue", isAuthenticated, (req, res) => {
    res.render("cep-queue", { user: req.session.user });
});

// ─── Start Server ─────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 8080;

if (process.env.VERCEL !== "1") {
    app.listen(PORT, () => {
        console.log(`Server listening on port ${PORT}`);
    });
}

module.exports = app;