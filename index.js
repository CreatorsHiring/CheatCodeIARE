const dotenv = require("dotenv");
const express = require("express");
const { Redis } = require("@upstash/redis");
const path = require("path");
const { GoogleGenerativeAI, HarmCategory, HarmBlockThreshold } = require("@google/generative-ai");
const PptxGenJS = require("pptxgenjs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const fs = require("fs");
const cookieParser = require("cookie-parser");



const { inject } = require("@vercel/analytics");

dotenv.config();
inject();


const redis = new Redis({
  url: process.env.UPSTASH_REDIS_REST_URL,
  token: process.env.UPSTASH_REDIS_REST_TOKEN
});

const app = express();


app.use(cookieParser());


// ─── Middleware ───────────────────────────────────────────────────────────────
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

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
const USER_COOKIE_NAME = "user";
const AUTH_COOKIE_OPTIONS = {
    httpOnly: true,
    secure: process.env.NODE_ENV === "production",
    sameSite: "lax"
};

function safeJsonParse(value) {
    if (!value || typeof value !== "string") return null;

    try {
        return JSON.parse(value);
    } catch {
        return null;
    }
}

function stripMarkdownCodeFences(text) {
    return String(text || "")
        .replace(/```json/gi, "")
        .replace(/```/g, "")
        .trim();
}

function extractJsonPayload(text) {
    const cleaned = stripMarkdownCodeFences(text);
    const startIndex = cleaned.search(/[\[{]/);

    if (startIndex === -1) return cleaned;

    const openChar = cleaned[startIndex];
    const closeChar = openChar === "{" ? "}" : "]";
    let depth = 0;
    let inString = false;
    let escaped = false;

    for (let i = startIndex; i < cleaned.length; i++) {
        const ch = cleaned[i];

        if (inString) {
            if (escaped) {
                escaped = false;
            } else if (ch === "\\") {
                escaped = true;
            } else if (ch === "\"") {
                inString = false;
            }
            continue;
        }

        if (ch === "\"") {
            inString = true;
            continue;
        }

        if (ch === openChar) depth++;
        if (ch === closeChar) {
            depth--;
            if (depth === 0) {
                return cleaned.slice(startIndex, i + 1);
            }
        }
    }

    return cleaned.slice(startIndex);
}

function escapeJsonControlCharsInStrings(jsonText) {
    let output = "";
    let inString = false;
    let escaped = false;

    for (const ch of jsonText) {
        if (inString) {
            if (escaped) {
                output += ch;
                escaped = false;
                continue;
            }

            if (ch === "\\") {
                output += ch;
                escaped = true;
                continue;
            }

            if (ch === "\"") {
                output += ch;
                inString = false;
                continue;
            }

            const code = ch.charCodeAt(0);
            if (code <= 0x1f) {
                if (ch === "\n") output += "\\n";
                else if (ch === "\r") output += "\\r";
                else if (ch === "\t") output += "\\t";
                else if (ch === "\b") output += "\\b";
                else if (ch === "\f") output += "\\f";
                else output += `\\u${code.toString(16).padStart(4, "0")}`;
                continue;
            }
        } else if (ch === "\"") {
            inString = true;
        }

        output += ch;
    }

    return output;
}

function parseAiJson(text) {
    const payload = extractJsonPayload(text);

    try {
        return JSON.parse(payload);
    } catch (error) {
        const repaired = escapeJsonControlCharsInStrings(payload);

        try {
            return JSON.parse(repaired);
        } catch (repairError) {
            throw new Error(`Failed to parse AI JSON: ${repairError.message}`);
        }
    }
}

function normalizeUser(value) {
    const source = value && typeof value === "object" ? value : {};
    const name = String(source.name || "").trim();
    const studentId = String(source.studentId || "").trim();

    if (!name || !studentId) return null;

    return { name, studentId };
}

function parseUserCookie(rawCookie) {
    return normalizeUser(safeJsonParse(rawCookie));
}

function setUserCookie(res, user) {
    res.cookie(USER_COOKIE_NAME, JSON.stringify(user), AUTH_COOKIE_OPTIONS);
}

function clearUserCookie(res) {
    res.clearCookie(USER_COOKIE_NAME, AUTH_COOKIE_OPTIONS);
}

function sanitizeReturnTo(returnTo) {
    if (typeof returnTo !== "string") return "/";
    if (!returnTo.startsWith("/") || returnTo.startsWith("//")) return "/";
    return returnTo;
}

function isOwnUser(req, userId) {
    return !!req.user && String(req.user.studentId) === String(userId);
}

async function getRedisJson(key) {
    const raw = await redis.get(key);
    if (!raw) return null;
    return typeof raw === "string" ? safeJsonParse(raw) : raw;
}

async function setRedisJson(key, value, ttlSeconds) {
    await redis.set(key, JSON.stringify(value), { ex: ttlSeconds });
}

async function getCompletedPptResult(userId) {
    const data = await getRedisJson(`ppt_result:${userId}`);
    if (!data || data.status !== "done" || !data.pptData) return null;
    return data;
}

app.use((req, res, next) => {
    const parsedUser = parseUserCookie(req.cookies?.[USER_COOKIE_NAME]);

    if (parsedUser) {
        req.user = parsedUser;
    } else if (req.cookies?.[USER_COOKIE_NAME]) {
        console.warn("[Auth] Clearing malformed user cookie.");
        clearUserCookie(res);
    }

    next();
});

const isAuthenticated = (req, res, next) => {
    if (req.user) {
        return next();
    }

    console.warn(`[Auth] Blocked ${req.method} ${req.originalUrl}`);
    return res.redirect(`/login?returnTo=${encodeURIComponent(req.originalUrl)}`);
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

    return parseAiJson(responseText);
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
    res.render("login", { returnTo: sanitizeReturnTo(req.query.returnTo) });
});

app.post("/login", (req, res) => {
    const user = normalizeUser(req.body);
    const returnTo = sanitizeReturnTo(req.body.returnTo || req.query.returnTo);

    if (!user) {
        console.warn("[Auth] Login failed: missing name or studentId.");
        return res.redirect(`/login?returnTo=${encodeURIComponent(returnTo)}`);
    }

    setUserCookie(res, user);
    console.log(`[Auth] Login success for ${user.studentId}`);
    return res.redirect(returnTo);
});

// ─── Signup ───────────────────────────────────────────────────────────────────
app.get("/signup", (req, res) => {
    res.render("signup");
});

app.post("/signup", (req, res) => {
    const { name, studentId, email, password } = req.body;
    if (name && studentId && email && password) {
        setUserCookie(res, { name: String(name).trim(), studentId: String(studentId).trim() });
        res.redirect("/");
    } else {
        res.redirect("/signup");
    }
});

app.get("/logout", (req, res) => {
    if (req.user) {
        console.log(`[Auth] Logout for ${req.user.studentId}`);
    }
    clearUserCookie(res);
    res.redirect("/login");
});

app.post("/logout", (req, res) => {
    if (req.user) {
        console.log(`[Auth] Logout for ${req.user.studentId}`);
    }
    clearUserCookie(res);
    res.redirect("/login");
});

// ─── PPT Form Page ────────────────────────────────────────────────────────────
app.get("/ppt", isAuthenticated, (req, res) => {
    res.render("ppt", { user: req.user });
});

// ─── Redis Queue Constants ────────────────────────────────────────────────────
const MAX_ACTIVE_JOBS = 4;
const EFFECTIVE_JOB_TTL_SECONDS = 60 * 60;
const PPT_RUNNER_STALE_MS = 90 * 1000;
const PPT_RUNNER_LOCK_TTL = 120;

function getPptRunnerKey(userId) {
    return `ppt_runner:${userId}`;
}

async function touchPptRunner(userId) {
    await redis.set(getPptRunnerKey(userId), String(Date.now()), { ex: PPT_RUNNER_LOCK_TTL });
}

async function ensurePptWorkerRunning(userId) {
    const runnerKey = getPptRunnerKey(userId);
    const heartbeat = parseInt(await redis.get(runnerKey) || "0", 10);

    if (!isNaN(heartbeat) && heartbeat > 0 && Date.now() - heartbeat <= PPT_RUNNER_STALE_MS) {
        return false;
    }

    if (!isNaN(heartbeat) && heartbeat > 0) {
        console.warn(`[Status] Stale PPT worker heartbeat for ${userId}, reclaiming lock`);
        await redis.del(runnerKey);
    }

    console.log(`[Status] Resuming PPT worker for ${userId}`);
    runPptJob(userId).catch(err => console.error(`[Status] PPT resume failed ${userId}:`, err));
    return true;
}

async function refreshPptQueuePositions() {
    const queue = await redis.lrange("ppt_queue", 0, -1);

    for (let i = 0; i < queue.length; i++) {
        await setRedisJson(
            `ppt_result:${queue[i]}`,
            { status: "waiting", position: i + 1 },
            EFFECTIVE_JOB_TTL_SECONDS
        );
    }
}
const JOB_TTL_SECONDS = 600; // 10 min — auto-cleanup if something crashes

// ─── Generate PPT → Queue-based with Redis ────────────────────────────────────
app.post("/generate-ppt", isAuthenticated, async (req, res) => {
    const { department, subject, problemStatement } = req.body;
    const userId = req.user.studentId;

    try {
        // Check if user already has a job running or waiting
        const existing = await getRedisJson(`ppt_result:${userId}`);
        if (existing) {
            if (existing.status === "waiting" || existing.status === "generating") {
                return res.json({ queued: true, position: existing.position || 0 });
            }
        }

        const existingQueuePos = await redis.lpos("ppt_queue", userId);
        if (existingQueuePos !== null) {
            await setRedisJson(
                `ppt_result:${userId}`,
                { status: "waiting", position: existingQueuePos + 1 },
                EFFECTIVE_JOB_TTL_SECONDS
            );
            return res.json({ queued: true, position: existingQueuePos + 1 });
        }

        const existingProcessingPos = await redis.lpos("ppt_processing", userId);
        if (existingProcessingPos !== null) {
            await ensurePptWorkerRunning(userId);
            await setRedisJson(
                `ppt_result:${userId}`,
                { status: "generating", startedAt: Date.now() },
                EFFECTIVE_JOB_TTL_SECONDS
            );
            return res.json({ queued: true, processing: true });
        }

        // Save form data + user identity to Redis (session not reliable across Vercel instances)
        await setRedisJson(
            `ppt_form:${userId}`,
            {
                department, subject, problemStatement,
                userName: req.user.name,
                studentId: req.user.studentId
            },
            EFFECTIVE_JOB_TTL_SECONDS
        );

        // Remove user from queue if already in it (re-submit scenario)
        await redis.lrem("ppt_queue", 0, userId);

        // Add to queue
        await redis.rpush("ppt_queue", userId);

        // Get their position (1-based)
        const queue = await redis.lrange("ppt_queue", 0, -1);
        const position = queue.indexOf(userId) + 1;

        // Save initial waiting status with TTL
        await setRedisJson(
            `ppt_result:${userId}`,
            { status: "waiting", position },
            EFFECTIVE_JOB_TTL_SECONDS
        );

        console.log(`[Queue] Enqueued PPT for ${userId} at position ${position}`);

        // Try to start processing immediately before returning on Vercel.
        await processNextInQueue().catch(err => console.error("[Queue] processNextInQueue error:", err));

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
        const processing = await redis.lrange("ppt_processing", 0, -1);
        if (processing.length === 0 && activeJobs > 0) {
            console.warn(`[Queue] Resetting stale active count ${activeJobs} before dispatch`);
            activeJobs = 0;
            await redis.set("active_jobs", "0");
        } else if (activeJobs > processing.length) {
            console.warn(`[Queue] Clamping stale active count ${activeJobs} -> ${processing.length}`);
            activeJobs = processing.length;
            await redis.set("active_jobs", String(processing.length));
        }

        console.log(`[Queue] active=${activeJobs}/${MAX_ACTIVE_JOBS}`);

        if (activeJobs >= MAX_ACTIVE_JOBS) return; // finally releases lock

        const userId = await redis.lpop("ppt_queue");
        if (!userId) { console.log("[Queue] Queue empty."); return; }

        await redis.lpush("ppt_processing", userId);
        await redis.incr("active_jobs");
        await setRedisJson(
            `ppt_result:${userId}`,
            { status: "generating", startedAt: Date.now() },
            EFFECTIVE_JOB_TTL_SECONDS
        );

        await refreshPptQueuePositions();

        console.log(`[Queue] Worker start for ${userId} | active=${activeJobs + 1}`);
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
    const runnerKey = getPptRunnerKey(userId);
    let runnerLock = await redis.set(runnerKey, String(Date.now()), { nx: true, ex: PPT_RUNNER_LOCK_TTL });

    if (!runnerLock) {
        const heartbeat = parseInt(await redis.get(runnerKey) || "0", 10);
        if (!isNaN(heartbeat) && heartbeat > 0 && Date.now() - heartbeat <= PPT_RUNNER_STALE_MS) {
            console.log(`[Queue] PPT worker already running for ${userId}, skipping duplicate start`);
            return;
        }

        console.warn(`[Queue] Reclaiming stale PPT worker lock for ${userId}`);
        await redis.del(runnerKey);
        runnerLock = await redis.set(runnerKey, String(Date.now()), { nx: true, ex: PPT_RUNNER_LOCK_TTL });
        if (!runnerLock) {
            console.log(`[Queue] PPT worker lock still busy for ${userId}, skipping`);
            return;
        }
    }

    try {
        await touchPptRunner(userId);
        const active = await redis.get(`user_active:${userId}`);

        if (!active) {
            console.log(`[Queue] No recent heartbeat for ${userId}, continuing generation.`);
        }
        // Retrieve form data stored during /generate-ppt — stored in Redis too for safety
        const formData = await getRedisJson(`ppt_form:${userId}`);

        if (!formData) {
            throw new Error("Form data not found in Redis for user: " + userId);
        }
        await touchPptRunner(userId);

        const content = await generatePptContent(formData.problemStatement);
        await touchPptRunner(userId);

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
        await setRedisJson(`ppt_result:${userId}`, result, EFFECTIVE_JOB_TTL_SECONDS);
        await touchPptRunner(userId);
        console.log(`[Queue] Job done for user: ${userId}`);

    } catch (err) {
        console.error(`[Queue] Generation failed for ${userId}:`, err.message);
        await setRedisJson(
            `ppt_result:${userId}`,
            { status: "error", message: err.message },
            10 * 60
        );
    } finally {
        await redis.del(runnerKey);
        const nextActive = await redis.decr("active_jobs");
        if (nextActive < 0) {
            await redis.set("active_jobs", "0");
        }

        await redis.lrem("ppt_processing", 0, userId);
        console.log(`[Queue] Cleanup success for ${userId}`);

        processNextInQueue().catch(err =>
            console.error("[Queue] Failed to trigger next job:", err)
        );
    }
}

async function recoverStuckJobs() {
    // Throttle: only run once per 60s globally
    const throttle = await redis.set("recovery_lock", "1", { nx: true, ex: 60 });
    if (!throttle) return;

    try {
        let freedAny = false;
        const processing = await redis.lrange("ppt_processing", 0, -1);
        const STUCK_MS = 8 * 60 * 1000; // 8 min — Gemini never takes this long

        let active = parseInt(await redis.get("active_jobs") || "0", 10);
        if (isNaN(active) || active < 0) active = 0;

        if (processing.length === 0 && active > 0) {
            console.warn(`[Recovery] Resetting stale PPT active count ${active} with empty processing list`);
            await redis.set("active_jobs", "0");
            active = 0;
            freedAny = true;
        } else if (active > processing.length) {
            console.warn(`[Recovery] Clamping stale PPT active count ${active} -> ${processing.length}`);
            await redis.set("active_jobs", String(processing.length));
            active = processing.length;
            freedAny = true;
        }

        for (const userId of processing) {
            const result = await getRedisJson(`ppt_result:${userId}`);

            if (!result) {
                // Key expired — job never finished, free the slot
                console.warn(`[Recovery] Key expired for ${userId}, freeing slot`);
                await redis.lrem("ppt_processing", 0, userId);
                const n = await redis.decr("active_jobs");
                if (n < 0) await redis.set("active_jobs", "0");
                const pos = await redis.lpos("ppt_queue", userId);
                if (pos === null) await redis.rpush("ppt_queue", userId);
                await setRedisJson(
                    `ppt_result:${userId}`,
                    { status: "waiting", position: 99 },
                    EFFECTIVE_JOB_TTL_SECONDS
                );
                freedAny = true;
                continue;
            }

            if (result.status === "generating") {
                const age = Date.now() - (result.startedAt || 0);
                if (!result.startedAt || age > STUCK_MS) {
                    console.warn(`[Recovery] Stuck job for ${userId} (age=${Math.round(age/1000)}s), re-queuing`);
                    await redis.lrem("ppt_processing", 0, userId);
                    const n = await redis.decr("active_jobs");
                    if (n < 0) await redis.set("active_jobs", "0");
                    const pos = await redis.lpos("ppt_queue", userId);
                    if (pos === null) await redis.rpush("ppt_queue", userId);
                    await setRedisJson(
                        `ppt_result:${userId}`,
                        { status: "waiting", position: 99 },
                        EFFECTIVE_JOB_TTL_SECONDS
                    );
                    freedAny = true;
                }
            }
        }
        await refreshPptQueuePositions();

        const queueLength = await redis.llen("ppt_queue");
        const refreshedActive = parseInt(await redis.get("active_jobs") || "0", 10);
        if (queueLength > 0 && (isNaN(refreshedActive) || refreshedActive < MAX_ACTIVE_JOBS)) {
            freedAny = true;
        }

        if (freedAny) {
            console.log("[Recovery] Freed stuck PPT slots, restarting queue...");
            processNextInQueue().catch(err =>
                console.error("[Recovery] Failed to restart PPT queue:", err)
            );
        }
    } catch (err) {
        console.error("[Recovery] error:", err);
    }
}

async function recoverStuckReports() {
    const throttle = await redis.set("report_recovery_lock", "1", { nx: true, ex: 60 });
    if (!throttle) return;

    try {
        let freedAny = false;

        const processing = await redis.lrange("report_processing", 0, -1);
        const STUCK_MS = 8 * 60 * 1000; // 8 minutes
        let active = parseInt(await redis.get("report_active") || "0", 10);
        if (isNaN(active) || active < 0) active = 0;

        if (processing.length === 0 && active > 0) {
            console.warn(`[ReportRecovery] Resetting stale active count ${active} with empty processing list`);
            await redis.set("report_active", "0");
            active = 0;
            freedAny = true;
        } else if (active > processing.length) {
            console.warn(`[ReportRecovery] Clamping stale active count ${active} -> ${processing.length}`);
            await redis.set("report_active", String(processing.length));
            active = processing.length;
            freedAny = true;
        }

        for (const userId of processing) {
            const result = await getRedisJson(`report_result:${userId}`);

            if (!result) {
                console.warn(`[ReportRecovery] Missing result for ${userId}, freeing slot`);
                await redis.lrem("report_processing", 0, userId);

                const n = await redis.decr("report_active");
                if (n < 0) await redis.set("report_active", "0");

                const pos = await redis.lpos("report_queue", userId);
                if (pos === null) await redis.rpush("report_queue", userId);

                await redis.set(
                    `report_result:${userId}`,
                    JSON.stringify({ status: "waiting", position: 99 }),
                    { ex: EFFECTIVE_REPORT_TTL }
                );

                freedAny = true; // <-- ADD THIS
                continue;
            }

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
                        { ex: EFFECTIVE_REPORT_TTL }
                    );

                    freedAny = true; // <-- ADD THIS
                }
            }
        }

        await refreshReportQueuePositions();

        const queueLength = await redis.llen("report_queue");
        const refreshedActive = parseInt(await redis.get("report_active") || "0", 10);
        if (queueLength > 0 && (isNaN(refreshedActive) || refreshedActive < MAX_REPORT_JOBS)) {
            freedAny = true;
        }

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
app.post("/heartbeat/:userId", isAuthenticated, async (req,res)=>{

    const { userId } = req.params;

    if (!isOwnUser(req, userId)) {
        console.warn(`[Auth] Heartbeat forbidden for ${req.user.studentId} -> ${userId}`);
        return res.sendStatus(403);
    }

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
app.get("/ppt-status/:userId", isAuthenticated, async (req, res) => {
    const { userId } = req.params;

    if (!isOwnUser(req, userId)) {
        console.warn(`[Auth] PPT status forbidden for ${req.user.studentId} -> ${userId}`);
        return res.status(403).json({ error: "Forbidden" });
    }

    try {
        // Throttled recovery — max once per 60s, not on every 3s poll
        await recoverStuckJobs().catch(console.error);
        await processNextInQueue().catch(console.error);

        const data = await getRedisJson(`ppt_result:${userId}`);
        if (!data) {
            const queuePos = await redis.lpos("ppt_queue", userId);
            const processingPos = await redis.lpos("ppt_processing", userId);

            if (queuePos !== null) {
                await setRedisJson(
                    `ppt_result:${userId}`,
                    { status: "waiting", position: queuePos + 1 },
                    EFFECTIVE_JOB_TTL_SECONDS
                );
                return res.json({ status: "waiting", position: queuePos + 1 });
            }

            if (processingPos !== null) {
                await ensurePptWorkerRunning(userId);
                await setRedisJson(
                    `ppt_result:${userId}`,
                    { status: "generating", startedAt: Date.now() },
                    EFFECTIVE_JOB_TTL_SECONDS
                );
                return res.json({ status: "generating" });
            }

            return res.json({ status: "not_found" });
        }

        if (data.status === "waiting") {
            const queuePos = await redis.lpos("ppt_queue", userId);
            const processingPos = await redis.lpos("ppt_processing", userId);

            if (processingPos !== null) {
                await ensurePptWorkerRunning(userId);
                await setRedisJson(
                    `ppt_result:${userId}`,
                    { status: "generating", startedAt: data.startedAt || Date.now() },
                    EFFECTIVE_JOB_TTL_SECONDS
                );
                return res.json({ status: "generating" });
            }

            if (queuePos === null && processingPos === null) {
                console.warn(`[Status] Re-queueing lost PPT job for ${userId}`);
                await redis.rpush("ppt_queue", userId);
                await refreshPptQueuePositions();
                await processNextInQueue().catch(console.error);
                const repaired = await getRedisJson(`ppt_result:${userId}`);
                return res.json(repaired || { status: "waiting", position: 1 });
            }

            return res.json({
                ...data,
                position: queuePos === null ? data.position || 1 : queuePos + 1
            });
        }

        if (data.status === "generating") {
            const processingPos = await redis.lpos("ppt_processing", userId);
            if (processingPos === null) {
                console.warn(`[Status] Generating PPT missing from processing for ${userId}, forcing recovery`);
                await recoverStuckJobs().catch(console.error);
                await processNextInQueue().catch(console.error);
            } else {
                await ensurePptWorkerRunning(userId);
            }
        }

        if (data.status === "done") {
            // Populate session so /edit and /download work on this instance
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
    res.render("queue", { user: req.user });
});

// ─── Edit Page ────────────────────────────────────────────────────────────────
app.get("/edit/:userId", isAuthenticated, async (req, res) => {

    const { userId } = req.params;

    if (!isOwnUser(req, userId)) {
        console.warn(`[Auth] Edit forbidden for ${req.user.studentId} -> ${userId}`);
        return res.status(403).send("Forbidden");
    }

    const data = await getCompletedPptResult(userId);

    if (!data) {
        return res.redirect("/ppt");
    }

    res.render("edit", { user: req.user, pptData: data.pptData });
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

        const pptResult = await getCompletedPptResult(req.user.studentId);
        const pptData = pptResult?.pptData;

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

        const pptResult = await getCompletedPptResult(req.user.studentId);
        const pptData = pptResult?.pptData;
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

        const updatedContent = parseAiJson(responseText);

        await setRedisJson(
            `ppt_result:${req.user.studentId}`,
            {
                ...pptResult,
                pptData: {
                    ...pptData,
                    content: updatedContent
                }
            },
            EFFECTIVE_JOB_TTL_SECONDS
        );

        res.json({ success: true, content: updatedContent });

    } catch (error) {
        console.error("Edit PPT API Error:", error);
        res.status(500).json({ success: false, error: "Failed to update presentation: " + error.message });
    }
});

// ─── Download PPT ─────────────────────────────────────────────────────────────
app.post("/download-ppt", isAuthenticated, async (req, res) => {
    try {
        const userId = String(req.body.userId || req.user.studentId || "").trim();
        if (!userId) return res.status(400).send("Missing userId.");
        if (!isOwnUser(req, userId)) {
            console.warn(`[Auth] PPT download forbidden for ${req.user.studentId} -> ${userId}`);
            return res.status(403).send("Forbidden");
        }

        const data = await getCompletedPptResult(userId);
        if (!data) return res.status(404).send("Presentation not found. Please generate again.");

        const pptData = data.pptData;
        const userName = data.userName || req.user.name || "Student";

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
        console.log(`[Download] PPT success for ${userId}`);

    } catch (err) {
        console.error("[Download] Error:", err);
        res.status(500).send("Download failed: " + err.message);

    }

});

// ─── Report Page ──────────────────────────────────────────────────────────────
app.get("/report", isAuthenticated, (req, res) => {
    res.render("report", { user: req.user });
});

// ─── Worksheets Page (Coming Soon) ───────────────────────────────────────────
app.get("/worksheets", isAuthenticated, (req, res) => {
    res.render("worksheets");
});

app.get("/worksheet", isAuthenticated, (req, res) => {
    res.redirect("/worksheets");
});

// ─── Complex Engineering Problems ────────────────────────────────────────────
app.get("/complex-engineering", isAuthenticated, (req, res) => {
    res.render("complex-engineering", { user: req.user });
});

app.get("/cep", isAuthenticated, (req, res) => {
    res.redirect("/complex-engineering");
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
            const parsed = parseAiJson(result.response.text());
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
const EFFECTIVE_REPORT_TTL = 60 * 60;
const REPORT_TTL       = 600; // 10 min
const REPORT_RUNNER_STALE_MS = 90 * 1000;
const REPORT_RUNNER_LOCK_TTL = 120;

function getReportRunnerKey(userId) {
    return `report_runner:${userId}`;
}

async function touchReportRunner(userId) {
    await redis.set(getReportRunnerKey(userId), String(Date.now()), { ex: REPORT_RUNNER_LOCK_TTL });
}

async function ensureReportWorkerRunning(userId) {
    const runnerKey = getReportRunnerKey(userId);
    const heartbeat = parseInt(await redis.get(runnerKey) || "0", 10);

    if (!isNaN(heartbeat) && heartbeat > 0 && Date.now() - heartbeat <= REPORT_RUNNER_STALE_MS) {
        return false;
    }

    if (!isNaN(heartbeat) && heartbeat > 0) {
        console.warn(`[ReportStatus] Stale worker heartbeat for ${userId}, reclaiming lock`);
        await redis.del(runnerKey);
    }

    console.log(`[ReportStatus] Resuming worker for ${userId}`);
    runReportJob(userId).catch(err => console.error(`[ReportStatus] resume failed ${userId}:`, err));
    return true;
}

async function refreshReportQueuePositions() {
    const queue = await redis.lrange("report_queue", 0, -1);

    for (let i = 0; i < queue.length; i++) {
        await setRedisJson(
            `report_result:${queue[i]}`,
            { status: "waiting", position: i + 1 },
            EFFECTIVE_REPORT_TTL
        );
    }
}

// ─── Enqueue Report ───────────────────────────────────────────────────────────
app.post("/generate-report", isAuthenticated, async (req, res) => {
    const userId = req.user.studentId;

    try {
        // Don't double-queue if already waiting/generating
        const existing = await getRedisJson(`report_result:${userId}`);
        if (existing) {
            if (existing.status === "waiting" || existing.status === "generating") {
                return res.json({ queued: true, position: existing.position || 1 });
            }
        }

        const existingQueuePos = await redis.lpos("report_queue", userId);
        if (existingQueuePos !== null) {
            await setRedisJson(
                `report_result:${userId}`,
                { status: "waiting", position: existingQueuePos + 1 },
                EFFECTIVE_REPORT_TTL
            );
            return res.json({ queued: true, position: existingQueuePos + 1 });
        }

        const existingProcessingPos = await redis.lpos("report_processing", userId);
        if (existingProcessingPos !== null) {
            await ensureReportWorkerRunning(userId);
            await setRedisJson(
                `report_result:${userId}`,
                { status: "generating", startedAt: Date.now() },
                EFFECTIVE_REPORT_TTL
            );
            return res.json({ queued: true, processing: true });
        }

        // Persist entire form to Redis — session unreliable across Vercel instances
        const formData = { ...req.body, authenticatedUser: req.user };
        await setRedisJson(
            `report_form:${userId}`,
            formData,
            EFFECTIVE_REPORT_TTL
        );

        // Add to queue
        await redis.lrem("report_queue", 0, userId);
        await redis.rpush("report_queue", userId);

        const queue    = await redis.lrange("report_queue", 0, -1);
        const position = queue.indexOf(userId) + 1;

        await setRedisJson(
            `report_result:${userId}`,
            { status: "waiting", position },
            EFFECTIVE_REPORT_TTL
        );

        console.log(`[ReportQueue] Enqueued ${userId} at position ${position}`);

        // Kick the worker before returning so Vercel does not drop the dispatch.
        await processNextReport().catch(err => console.error("[ReportQueue] boot error:", err));

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
        const processing = await redis.lrange("report_processing", 0, -1);
        if (processing.length === 0 && active > 0) {
            console.warn(`[ReportQueue] Resetting stale active count ${active} before dispatch`);
            active = 0;
            await redis.set("report_active", "0");
        } else if (active > processing.length) {
            console.warn(`[ReportQueue] Clamping stale active count ${active} -> ${processing.length}`);
            active = processing.length;
            await redis.set("report_active", String(processing.length));
        }

        console.log(`[ReportQueue] active=${active}/${MAX_REPORT_JOBS}`);
        if (active >= MAX_REPORT_JOBS) return;

        const userId = await redis.lpop("report_queue");
        if (!userId) { console.log("[ReportQueue] Queue empty."); return; }

        await redis.lpush("report_processing", userId);
        await redis.incr("report_active");
        await setRedisJson(
            `report_result:${userId}`,
            { status: "generating", startedAt: Date.now() },
            EFFECTIVE_REPORT_TTL
        );

        // Update positions for remaining waiters
        await refreshReportQueuePositions();

        console.log(`[ReportQueue] Worker start for ${userId} | active=${active + 1}`);
        runReportJob(userId).catch(err => console.error(`[ReportQueue] job error ${userId}:`, err));

    } catch (err) {
        console.error("[ReportQueue] processNextReport error:", err);
    } finally {
        await redis.del("report_lock");
    }
}

// ─── Run one report generation job ───────────────────────────────────────────
async function runReportJob(userId) {
    const runnerKey = getReportRunnerKey(userId);
    let runnerLock = await redis.set(runnerKey, String(Date.now()), { nx: true, ex: REPORT_RUNNER_LOCK_TTL });

    if (!runnerLock) {
        const heartbeat = parseInt(await redis.get(runnerKey) || "0", 10);
        if (!isNaN(heartbeat) && heartbeat > 0 && Date.now() - heartbeat <= REPORT_RUNNER_STALE_MS) {
            console.log(`[ReportQueue] Worker already running for ${userId}, skipping duplicate start`);
            return;
        }

        console.warn(`[ReportQueue] Reclaiming stale worker lock for ${userId}`);
        await redis.del(runnerKey);
        runnerLock = await redis.set(runnerKey, String(Date.now()), { nx: true, ex: REPORT_RUNNER_LOCK_TTL });
        if (!runnerLock) {
            console.log(`[ReportQueue] Worker lock still busy for ${userId}, skipping`);
            return;
        }
    }

    try {
        await touchReportRunner(userId);
        console.log(`[ReportQueue] Reading form data for ${userId}`);
        const formData = await getRedisJson(`report_form:${userId}`);
        if (!formData) throw new Error("Form data missing for user: " + userId);
        await touchReportRunner(userId);

        const questions = [];
        for (let i = 1; i <= 10; i++) questions.push(formData[`question${i}`] || "");

        const rawAnswers = await batchGenerateAnswers(questions);
        await touchReportRunner(userId);

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
        await setRedisJson(
            `report_result:${userId}`,
            {
                status: "done",
                fileName: (formData.name || "Student").replace(/[^a-zA-Z0-9_]/g, "_") + "_Report.docx",
                docBase64: buf.toString("base64")
            },
            EFFECTIVE_REPORT_TTL
        );
        await touchReportRunner(userId);

        console.log(`[ReportQueue] Job done for ${userId}`);

    } catch (err) {
        console.error(`[ReportQueue] Failed for ${userId}:`, err.message);
        await setRedisJson(
            `report_result:${userId}`,
            { status: "error", message: err.message },
            10 * 60
        );
    } finally {
        await redis.del(runnerKey);
        const n = await redis.decr("report_active");
        if (n < 0) await redis.set("report_active", "0");

        await redis.lrem("report_processing", 0, userId);
        console.log(`[ReportQueue] Cleanup success for ${userId}`);

        processNextReport().catch(err =>
            console.error("[ReportQueue] Failed to trigger next job:", err)
        );
    }
}

// ─── Report Status Polling ────────────────────────────────────────────────────
app.get("/report-status/:userId", isAuthenticated, async (req, res) => {
    const { userId } = req.params;

    if (!isOwnUser(req, userId)) {
        console.warn(`[Auth] Report status forbidden for ${req.user.studentId} -> ${userId}`);
        return res.status(403).json({ error: "Forbidden" });
    }

    try {
        await recoverStuckReports().catch(console.error);
        await processNextReport().catch(console.error);
        const data = await getRedisJson(`report_result:${userId}`);
        if (!data) {
            const queuePos = await redis.lpos("report_queue", userId);
            const processingPos = await redis.lpos("report_processing", userId);

            if (queuePos !== null) {
                await setRedisJson(
                    `report_result:${userId}`,
                    { status: "waiting", position: queuePos + 1 },
                    EFFECTIVE_REPORT_TTL
                );
                return res.json({ status: "waiting", position: queuePos + 1 });
            }

            if (processingPos !== null) {
                await ensureReportWorkerRunning(userId);
                await setRedisJson(
                    `report_result:${userId}`,
                    { status: "generating", startedAt: Date.now() },
                    EFFECTIVE_REPORT_TTL
                );
                return res.json({ status: "generating" });
            }

            return res.json({ status: "not_found" });
        }

        if (data.status === "waiting") {
            const queuePos = await redis.lpos("report_queue", userId);
            const processingPos = await redis.lpos("report_processing", userId);

            if (processingPos !== null) {
                await ensureReportWorkerRunning(userId);
                await setRedisJson(
                    `report_result:${userId}`,
                    { status: "generating", startedAt: data.startedAt || Date.now() },
                    EFFECTIVE_REPORT_TTL
                );
                return res.json({ status: "generating" });
            }

            if (queuePos === null && processingPos === null) {
                console.warn(`[ReportStatus] Re-queueing lost waiting job for ${userId}`);
                await redis.rpush("report_queue", userId);
                await refreshReportQueuePositions();
                await processNextReport().catch(console.error);
                const repaired = await getRedisJson(`report_result:${userId}`);
                return res.json(repaired || { status: "waiting", position: 1 });
            }

            return res.json({
                ...data,
                position: queuePos === null ? data.position || 1 : queuePos + 1
            });
        }

        if (data.status === "generating") {
            const processingPos = await redis.lpos("report_processing", userId);
            if (processingPos === null) {
                console.warn(`[ReportStatus] Generating job missing from processing for ${userId}, forcing recovery`);
                await recoverStuckReports().catch(console.error);
                await processNextReport().catch(console.error);
            } else {
                await ensureReportWorkerRunning(userId);
            }
        }

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
app.get("/download-report/:userId", isAuthenticated, async (req, res) => {
    const { userId } = req.params;

    if (!isOwnUser(req, userId)) {
        console.warn(`[Auth] Report download forbidden for ${req.user.studentId} -> ${userId}`);
        return res.status(403).send("Forbidden");
    }

    try {
        const data = await getRedisJson(`report_result:${userId}`);
        if (!data) return res.status(404).send("Report not found. Please generate again.");
        if (data.status !== "done") return res.status(400).send("Report not ready yet.");

        const buf      = Buffer.from(data.docBase64, "base64");
        const fileName = data.fileName || `${userId}_Report.docx`;

        res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        res.send(buf);

        // Clean up Redis after successful download
        await redis.del(`report_result:${userId}`);
        await redis.del(`report_form:${userId}`);
        console.log(`[ReportDownload] Success for ${userId}`);

    } catch (err) {
        console.error("[ReportDownload] error:", err);
        res.status(500).send("Download failed: " + err.message);
    }
});

// ─── Report Queue Waiting Page ────────────────────────────────────────────────
app.get("/report-queue", isAuthenticated, (req, res) => {
    res.render("report-queue", { user: req.user });
});

// /report-done route removed — download is automatic, no done page needed

// ══════════════════════════════════════════════════════════════════════════════
// COMPLEX ENGINEERING PROBLEMS — Queue-based generation with Docxtemplater
// ══════════════════════════════════════════════════════════════════════════════

const MAX_CEP_JOBS = 4;
const EFFECTIVE_CEP_TTL = 60 * 60;
const CEP_TTL      = 600; // 10 min

async function refreshCepQueuePositions() {
    const queue = await redis.lrange("cep_queue", 0, -1);

    for (let i = 0; i < queue.length; i++) {
        await setRedisJson(
            `cep_result:${queue[i]}`,
            { status: "waiting", position: i + 1 },
            EFFECTIVE_CEP_TTL
        );
    }
}

async function recoverStuckCEP() {
    const throttle = await redis.set("cep_recovery_lock", "1", { nx: true, ex: 60 });
    if (!throttle) return;

    try {
        let freedAny = false;
        const processing = await redis.lrange("cep_processing", 0, -1);
        const STUCK_MS = 8 * 60 * 1000;
        let active = parseInt(await redis.get("cep_active") || "0", 10);
        if (isNaN(active) || active < 0) active = 0;

        if (processing.length === 0 && active > 0) {
            console.warn(`[CEPRecovery] Resetting stale active count ${active} with empty processing list`);
            await redis.set("cep_active", "0");
            active = 0;
            freedAny = true;
        } else if (active > processing.length) {
            console.warn(`[CEPRecovery] Clamping stale active count ${active} -> ${processing.length}`);
            await redis.set("cep_active", String(processing.length));
            active = processing.length;
            freedAny = true;
        }

        for (const userId of processing) {
            const result = await getRedisJson(`cep_result:${userId}`);

            if (!result) {
                console.warn(`[CEPRecovery] Missing result for ${userId}, re-queuing`);
                await redis.lrem("cep_processing", 0, userId);
                const n = await redis.decr("cep_active");
                if (n < 0) await redis.set("cep_active", "0");
                const pos = await redis.lpos("cep_queue", userId);
                if (pos === null) await redis.rpush("cep_queue", userId);
                await setRedisJson(
                    `cep_result:${userId}`,
                    { status: "waiting", position: 99 },
                    EFFECTIVE_CEP_TTL
                );
                freedAny = true;
                continue;
            }

            if (result.status === "generating") {
                const age = Date.now() - (result.startedAt || 0);
                if (!result.startedAt || age > STUCK_MS) {
                    console.warn(`[CEPRecovery] Stuck CEP for ${userId}, re-queuing`);
                    await redis.lrem("cep_processing", 0, userId);
                    const n = await redis.decr("cep_active");
                    if (n < 0) await redis.set("cep_active", "0");
                    const pos = await redis.lpos("cep_queue", userId);
                    if (pos === null) await redis.rpush("cep_queue", userId);
                    await setRedisJson(
                        `cep_result:${userId}`,
                        { status: "waiting", position: 99 },
                        EFFECTIVE_CEP_TTL
                    );
                    freedAny = true;
                }
            }
        }

        await refreshCepQueuePositions();
        const queueLength = await redis.llen("cep_queue");
        const refreshedActive = parseInt(await redis.get("cep_active") || "0", 10);
        if (queueLength > 0 && (isNaN(refreshedActive) || refreshedActive < MAX_CEP_JOBS)) {
            freedAny = true;
        }

        if (freedAny) {
            console.log("[CEPRecovery] Freed stuck CEP slots, restarting queue...");
            processNextCEP().catch(err => console.error("[CEPRecovery] restart error:", err));
        }
    } catch (err) {
        console.error("[CEPRecovery] error:", err);
    }
}

// ─── Enqueue CEP ─────────────────────────────────────────────────────────────
app.post("/generate-cep", isAuthenticated, async (req, res) => {
    const userId = req.user.studentId;
    try {
        // Don't double-queue
        const existing = await getRedisJson(`cep_result:${userId}`);
        if (existing) {
            if (existing.status === "waiting" || existing.status === "generating") {
                return res.json({ queued: true, position: existing.position || 1 });
            }
        }

        const existingQueuePos = await redis.lpos("cep_queue", userId);
        if (existingQueuePos !== null) {
            await setRedisJson(
                `cep_result:${userId}`,
                { status: "waiting", position: existingQueuePos + 1 },
                EFFECTIVE_CEP_TTL
            );
            return res.json({ queued: true, position: existingQueuePos + 1 });
        }

        const existingProcessingPos = await redis.lpos("cep_processing", userId);
        if (existingProcessingPos !== null) {
            await setRedisJson(
                `cep_result:${userId}`,
                { status: "generating", startedAt: Date.now() },
                EFFECTIVE_CEP_TTL
            );
            return res.json({ queued: true, processing: true });
        }

        // Save form data to Redis
        await setRedisJson(`cep_form:${userId}`, { ...req.body, authenticatedUser: req.user }, EFFECTIVE_CEP_TTL);

        await redis.lrem("cep_queue", 0, userId);
        await redis.rpush("cep_queue", userId);

        const queue    = await redis.lrange("cep_queue", 0, -1);
        const position = queue.indexOf(userId) + 1;

        await setRedisJson(
            `cep_result:${userId}`,
            { status: "waiting", position },
            EFFECTIVE_CEP_TTL
        );

        console.log(`[CEP] Enqueued ${userId} at position ${position}`);
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
        const processing = await redis.lrange("cep_processing", 0, -1);
        if (processing.length === 0 && active > 0) {
            console.warn(`[CEP] Resetting stale active count ${active} before dispatch`);
            active = 0;
            await redis.set("cep_active", "0");
        } else if (active > processing.length) {
            console.warn(`[CEP] Clamping stale active count ${active} -> ${processing.length}`);
            active = processing.length;
            await redis.set("cep_active", String(processing.length));
        }

        console.log(`[CEP] active=${active}/${MAX_CEP_JOBS}`);
        if (active >= MAX_CEP_JOBS) return;

        const userId = await redis.lpop("cep_queue");
        if (!userId) { console.log("[CEP] Queue empty."); return; }

        await redis.lpush("cep_processing", userId);
        await redis.incr("cep_active");
        await setRedisJson(
            `cep_result:${userId}`,
            { status: "generating", startedAt: Date.now() },
            EFFECTIVE_CEP_TTL
        );

        await refreshCepQueuePositions();

        console.log(`[CEP] Worker start for ${userId} | active=${active + 1}`);
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
        const form = await getRedisJson(`cep_form:${userId}`);
        if (!form) throw new Error("Form data missing for user: " + userId);
        console.log(`[CEP] Reading form data for ${userId}`);

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
  "abstract": "Write 6-8 sentences as a formal academic abstract covering: problem context, approach, techniques used, and expected outcomes. Separate each sentence with \n.",
  "introduction": "Write 5-6 sentences introducing the topic, its background, real-world relevance, and engineering importance. Separate each sentence with \n.",
  "overview": "Write 4-5 sentences giving an overview of the project scope and approach. Separate each sentence with \n.",
  "objectives": "List exactly 5 objectives. Format as: 1. <sentence>\n2. <sentence>\n3. <sentence>\n4. <sentence>\n5. <sentence>",
  "prerequisites": "List exactly 4 prerequisite knowledge areas. Format as: 1. <subject — why needed>\n2. <subject — why needed>\n3. <subject — why needed>\n4. <subject — why needed>",
  "requirements": "List 5-6 technical/resource requirements. Format as: 1. <requirement>\n2. <requirement>\n3. <requirement>\n4. <requirement>\n5. <requirement>",
  "methodology": "Write 5-6 sentences describing the methodology and approach specific to this domain. Separate each sentence with \n.",
  "workflow": "List exactly 5 workflow steps. Format as: Step 1: <sentence>\nStep 2: <sentence>\nStep 3: <sentence>\nStep 4: <sentence>\nStep 5: <sentence>",
  "content": "Write a detailed technical explanation (8-10 sentences) of the solution corresponding to the workflow steps. Include techniques, algorithms, or design approaches. Separate each sentence with \n.",
  "result": "Write 3-4 sentences describing the expected results and outcomes. Separate each sentence with \n.",
  "conclusion": "Write 4-5 sentences summarising what was achieved and its significance. Separate each sentence with \n.",
  "futureScope": "List 4-5 future scope items. Format as: 1. <item>\n2. <item>\n3. <item>\n4. <item>"
}

CRITICAL RULES:
- Use literal \n (backslash-n) inside the JSON string values to separate items and sentences.
- Do NOT use markdown (no **, no ##, no bullets •).
- Do NOT use actual newlines inside JSON strings — use \n only.
- All values must be plain text strings with \n separators.
- Be specific and technical — this is for a university engineering report.
`;

        const aiResult = await generateWithFallback(m => m.generateContent(prompt), true);
        const aiData = parseAiJson(aiResult.response.text());

        // ── Post-process AI text: ensure numbered items are on their own lines ──────
        // Handles cases where the AI returns "1. Foo 2. Bar" instead of "1. Foo\n2. Bar"
        function ensureLineBreaks(text) {
            if (!text) return "";
            // Already has newlines — just clean up
            if (text.includes("\n")) {
                return text
                    .replace(/\n+/g, "\n")   // collapse multiple newlines
                    .trim();
            }
            // Insert newline before numbered items: "1." "2." "3." etc
            text = text.replace(/(?<=[\w,.!?])\s+(?=\d+\.\s)/g, "\n");
            // Insert newline before "Step N:"
            text = text.replace(/(?<=[\w,.!?])\s+(?=Step\s+\d+:)/gi, "\n");
            return text.trim();
        }

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
            // AI-generated content — run through ensureLineBreaks so every
            // numbered item / sentence gets its own line in the Word doc
            abstract:      ensureLineBreaks(aiData.abstract      || ""),
            introduction:  ensureLineBreaks(aiData.introduction  || ""),
            overview:      ensureLineBreaks(aiData.overview      || ""),
            objectives:    ensureLineBreaks(aiData.objectives    || ""),
            prerequisites: ensureLineBreaks(aiData.prerequisites || ""),
            requirements:  ensureLineBreaks(aiData.requirements  || ""),
            methodology:   ensureLineBreaks(aiData.methodology   || ""),
            workflow:      ensureLineBreaks(aiData.workflow      || ""),
            content:       ensureLineBreaks(aiData.content       || ""),
            results:       ensureLineBreaks(aiData.result        || ""),  // template uses {results}
            result:        ensureLineBreaks(aiData.result        || ""),  // keep singular too
            conclusion:    ensureLineBreaks(aiData.conclusion    || ""),
            futureScope:   ensureLineBreaks(aiData.futureScope   || ""),
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

        await setRedisJson(
            `cep_result:${userId}`,
            {
                status: "done",
                fileName: `${safeName}_CEP.docx`,
                docBase64: buf.toString("base64")
            },
            EFFECTIVE_CEP_TTL
        );

        console.log(`[CEP] Done for ${userId}`);

    } catch (err) {
        console.error(`[CEP] Failed for ${userId}:`, err.message);
        await setRedisJson(
            `cep_result:${userId}`,
            { status: "error", message: err.message },
            10 * 60
        );
    } finally {
        const n = await redis.decr("cep_active");
        if (n < 0) await redis.set("cep_active", "0");
        await redis.lrem("cep_processing", 0, userId);
        console.log(`[CEP] Cleanup success for ${userId}`);
        processNextCEP().catch(err => console.error("[CEP] next-job error:", err));
    }
}

// ─── CEP Status Polling ───────────────────────────────────────────────────────
app.get("/cep-status/:userId", isAuthenticated, async (req, res) => {
    const { userId } = req.params;

    if (!isOwnUser(req, userId)) {
        console.warn(`[Auth] CEP status forbidden for ${req.user.studentId} -> ${userId}`);
        return res.status(403).json({ error: "Forbidden" });
    }

    try {
        recoverStuckCEP().catch(console.error);
        processNextCEP().catch(console.error);
        const data = await getRedisJson(`cep_result:${userId}`);
        if (!data) {
            const queuePos = await redis.lpos("cep_queue", userId);
            const processingPos = await redis.lpos("cep_processing", userId);

            if (queuePos !== null) {
                await setRedisJson(
                    `cep_result:${userId}`,
                    { status: "waiting", position: queuePos + 1 },
                    EFFECTIVE_CEP_TTL
                );
                return res.json({ status: "waiting", position: queuePos + 1 });
            }

            if (processingPos !== null) {
                await setRedisJson(
                    `cep_result:${userId}`,
                    { status: "generating", startedAt: Date.now() },
                    EFFECTIVE_CEP_TTL
                );
                return res.json({ status: "generating" });
            }

            return res.json({ status: "not_found" });
        }

        if (data.status === "waiting") {
            const queuePos = await redis.lpos("cep_queue", userId);
            const processingPos = await redis.lpos("cep_processing", userId);

            if (queuePos === null && processingPos === null) {
                console.warn(`[CEPStatus] Re-queueing lost waiting job for ${userId}`);
                await redis.rpush("cep_queue", userId);
                await refreshCepQueuePositions();
                processNextCEP().catch(console.error);
                const repaired = await getRedisJson(`cep_result:${userId}`);
                return res.json(repaired || { status: "waiting", position: 1 });
            }

            return res.json({
                ...data,
                position: queuePos === null ? data.position || 1 : queuePos + 1
            });
        }

        if (data.status === "generating") {
            const processingPos = await redis.lpos("cep_processing", userId);
            if (processingPos === null) {
                console.warn(`[CEPStatus] Generating job missing from processing for ${userId}, forcing recovery`);
                recoverStuckCEP().catch(console.error);
                processNextCEP().catch(console.error);
            }
        }

        if (data.status === "done") return res.json({ status: "done", fileName: data.fileName });
        return res.json(data);
    } catch (err) {
        console.error("[CEP] Status error:", err);
        return res.status(500).json({ error: "Status check failed." });
    }
});

// ─── CEP Download ─────────────────────────────────────────────────────────────
app.get("/download-cep/:userId", isAuthenticated, async (req, res) => {
    const { userId } = req.params;

    if (!isOwnUser(req, userId)) {
        console.warn(`[Auth] CEP download forbidden for ${req.user.studentId} -> ${userId}`);
        return res.status(403).send("Forbidden");
    }

    try {
        const data = await getRedisJson(`cep_result:${userId}`);
        if (!data) return res.status(404).send("Document not found. Please generate again.");
        if (data.status !== "done") return res.status(400).send("Document not ready yet.");

        const buf      = Buffer.from(data.docBase64, "base64");
        const fileName = data.fileName || `${userId}_CEP.docx`;

        res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        res.send(buf);

        await redis.del(`cep_result:${userId}`);
        await redis.del(`cep_form:${userId}`);
        console.log(`[CEP] Download success for ${userId}`);
    } catch (err) {
        console.error("[CEP] Download error:", err);
        res.status(500).send("Download failed: " + err.message);
    }
});

// ─── CEP Queue Waiting Page ───────────────────────────────────────────────────
app.get("/cep-queue", isAuthenticated, (req, res) => {
    res.render("cep-queue", { user: req.user });
});

// ══════════════════════════════════════════════════════════════════════════════
// CASE STUDY — Queue-based generation with Docxtemplater
// ══════════════════════════════════════════════════════════════════════════════

const MAX_CASE_STUDY_JOBS = 4;
const EFFECTIVE_CASE_STUDY_TTL = 60 * 60;

async function refreshCaseStudyQueuePositions() {
    const queue = await redis.lrange("case_study_queue", 0, -1);

    for (let i = 0; i < queue.length; i++) {
        await setRedisJson(
            `case_study_result:${queue[i]}`,
            { status: "waiting", position: i + 1 },
            EFFECTIVE_CASE_STUDY_TTL
        );
    }
}

async function recoverStuckCaseStudy() {
    const throttle = await redis.set("case_study_recovery_lock", "1", { nx: true, ex: 60 });
    if (!throttle) return;

    try {
        let freedAny = false;
        const processing = await redis.lrange("case_study_processing", 0, -1);
        const STUCK_MS = 8 * 60 * 1000;
        let active = parseInt(await redis.get("case_study_active") || "0", 10);
        if (isNaN(active) || active < 0) active = 0;

        if (processing.length === 0 && active > 0) {
            console.warn(`[CaseStudyRecovery] Resetting stale active count ${active} with empty processing list`);
            await redis.set("case_study_active", "0");
            active = 0;
            freedAny = true;
        } else if (active > processing.length) {
            console.warn(`[CaseStudyRecovery] Clamping stale active count ${active} -> ${processing.length}`);
            await redis.set("case_study_active", String(processing.length));
            active = processing.length;
            freedAny = true;
        }

        for (const userId of processing) {
            const result = await getRedisJson(`case_study_result:${userId}`);

            if (!result) {
                console.warn(`[CaseStudyRecovery] Missing result for ${userId}, re-queuing`);
                await redis.lrem("case_study_processing", 0, userId);
                const n = await redis.decr("case_study_active");
                if (n < 0) await redis.set("case_study_active", "0");
                const pos = await redis.lpos("case_study_queue", userId);
                if (pos === null) await redis.rpush("case_study_queue", userId);
                await setRedisJson(
                    `case_study_result:${userId}`,
                    { status: "waiting", position: 99 },
                    EFFECTIVE_CASE_STUDY_TTL
                );
                freedAny = true;
                continue;
            }

            if (result.status === "generating") {
                const age = Date.now() - (result.startedAt || 0);
                if (!result.startedAt || age > STUCK_MS) {
                    console.warn(`[CaseStudyRecovery] Stuck case study for ${userId}, re-queuing`);
                    await redis.lrem("case_study_processing", 0, userId);
                    const n = await redis.decr("case_study_active");
                    if (n < 0) await redis.set("case_study_active", "0");
                    const pos = await redis.lpos("case_study_queue", userId);
                    if (pos === null) await redis.rpush("case_study_queue", userId);
                    await setRedisJson(
                        `case_study_result:${userId}`,
                        { status: "waiting", position: 99 },
                        EFFECTIVE_CASE_STUDY_TTL
                    );
                    freedAny = true;
                }
            }
        }

        await refreshCaseStudyQueuePositions();
        const queueLength = await redis.llen("case_study_queue");
        const refreshedActive = parseInt(await redis.get("case_study_active") || "0", 10);
        if (queueLength > 0 && (isNaN(refreshedActive) || refreshedActive < MAX_CASE_STUDY_JOBS)) {
            freedAny = true;
        }

        if (freedAny) {
            console.log("[CaseStudyRecovery] Freed stuck case study slots, restarting queue...");
            processNextCaseStudy().catch(err => console.error("[CaseStudyRecovery] restart error:", err));
        }
    } catch (err) {
        console.error("[CaseStudyRecovery] error:", err);
    }
}

// ─── Enqueue Case Study ───────────────────────────────────────────────────────
app.post("/generate-case-study", isAuthenticated, async (req, res) => {
    const userId = req.user.studentId;
    try {
        // Don't double-queue
        const existing = await getRedisJson(`case_study_result:${userId}`);
        if (existing) {
            if (existing.status === "waiting" || existing.status === "generating") {
                return res.json({ queued: true, position: existing.position || 1 });
            }
        }

        const existingQueuePos = await redis.lpos("case_study_queue", userId);
        if (existingQueuePos !== null) {
            await setRedisJson(
                `case_study_result:${userId}`,
                { status: "waiting", position: existingQueuePos + 1 },
                EFFECTIVE_CASE_STUDY_TTL
            );
            return res.json({ queued: true, position: existingQueuePos + 1 });
        }

        const existingProcessingPos = await redis.lpos("case_study_processing", userId);
        if (existingProcessingPos !== null) {
            await setRedisJson(
                `case_study_result:${userId}`,
                { status: "generating", startedAt: Date.now() },
                EFFECTIVE_CASE_STUDY_TTL
            );
            return res.json({ queued: true, processing: true });
        }

        // Save form data to Redis
        await setRedisJson(`case_study_form:${userId}`, { ...req.body, authenticatedUser: req.user }, EFFECTIVE_CASE_STUDY_TTL);

        await redis.lrem("case_study_queue", 0, userId);
        await redis.rpush("case_study_queue", userId);

        const queue    = await redis.lrange("case_study_queue", 0, -1);
        const position = queue.indexOf(userId) + 1;

        await setRedisJson(
            `case_study_result:${userId}`,
            { status: "waiting", position },
            EFFECTIVE_CASE_STUDY_TTL
        );

        console.log(`[CaseStudy] Enqueued ${userId} at position ${position}`);
        processNextCaseStudy().catch(err => console.error("[CaseStudy] boot error:", err));
        return res.json({ queued: true, position });

    } catch (err) {
        console.error("[CaseStudy] Enqueue error:", err);
        return res.status(500).json({ error: "Failed to queue: " + err.message });
    }
});

// ─── Case Study Queue Worker ──────────────────────────────────────────────────
async function processNextCaseStudy() {
    const lock = await redis.set("case_study_lock", "1", { nx: true, ex: 15 });
    if (!lock) return;

    try {
        let active = parseInt(await redis.get("case_study_active") || "0", 10);
        if (isNaN(active) || active < 0) { active = 0; await redis.set("case_study_active", "0"); }
        const processing = await redis.lrange("case_study_processing", 0, -1);
        if (processing.length === 0 && active > 0) {
            console.warn(`[CaseStudy] Resetting stale active count ${active} before dispatch`);
            active = 0;
            await redis.set("case_study_active", "0");
        } else if (active > processing.length) {
            console.warn(`[CaseStudy] Clamping stale active count ${active} -> ${processing.length}`);
            active = processing.length;
            await redis.set("case_study_active", String(processing.length));
        }

        console.log(`[CaseStudy] active=${active}/${MAX_CASE_STUDY_JOBS}`);
        if (active >= MAX_CASE_STUDY_JOBS) return;

        const userId = await redis.lpop("case_study_queue");
        if (!userId) { console.log("[CaseStudy] Queue empty."); return; }

        await redis.lpush("case_study_processing", userId);
        await redis.incr("case_study_active");
        await setRedisJson(
            `case_study_result:${userId}`,
            { status: "generating", startedAt: Date.now() },
            EFFECTIVE_CASE_STUDY_TTL
        );

        await refreshCaseStudyQueuePositions();

        console.log(`[CaseStudy] Worker start for ${userId} | active=${active + 1}`);
        runCaseStudyJob(userId).catch(err => console.error(`[CaseStudy] job error ${userId}:`, err));

    } catch (err) {
        console.error("[CaseStudy] processNextCaseStudy error:", err);
    } finally {
        await redis.del("case_study_lock");
    }
}

// ─── Run one Case Study generation job ────────────────────────────────────────
async function runCaseStudyJob(userId) {
    try {
        const form = await getRedisJson(`case_study_form:${userId}`);
        if (!form) throw new Error("Form data missing for user: " + userId);
        console.log(`[CaseStudy] Reading form data for ${userId}`);

        const ps = form.problemStatement || "";
        const topicInput = form.topicName || form.topic || "";

        const rawDate = form.date || new Date().toISOString().split("T")[0];
        const dateObj = new Date(rawDate);
        const formattedDate = dateObj.toLocaleDateString("en-IN", {
            day: "2-digit",
            month: "long",
            year: "numeric"
        });

        const prompt = `
You are an expert academic case-study writer for engineering students.

Generate content for a university case study report for:
Topic: "${topicInput}"
Problem Statement: "${ps}"
Institute context: "IARE (Institute of Aeronautical Engineering)"

Return ONLY a valid JSON object with these exact keys and plain string values:
{
  "problemStatement": "One concise sentence (12-18 words) describing the central problem. It must work in both heading and summary contexts.",
  "principleTopic1": "Short side heading (3-7 words) for principles section",
  "principleTopic2": "Another short side heading (3-7 words) for principles section",
  "problemTopic1": "Short side heading (3-7 words) for problem subsection",
  "problemTopic2": "Another short side heading (3-7 words) for problem subsection",
  "differentApproaches": "A short heading (4-9 words) for section 5",
  "approachTopic1": "Short side heading (3-7 words) for approach subsection",
  "approachTopic2": "Another short side heading (3-7 words) for approach subsection",
  "application": "A short heading (3-8 words) for applications section",
  "applicationTopic1": "Short side heading (3-7 words) for applications subsection",
  "applicationTopic2": "Another short side heading (3-7 words) for applications subsection",
  "designAndAnalysis": "A short heading (4-9 words) for section 7",
  "pageNo": "Use a simple repeated value such as '1'",
  "introduction": "Paragraph about topic introduction",
  "useInCampus": "Paragraph on usage in IARE campus context",
  "focus": "Paragraph on the focus and solution direction",
  "whyThisTechnique": "Paragraph on why this technique/solution is chosen",
  "observation": "Paragraph on observations from the problem context",
  "scope": "Paragraph on solution scope and boundaries",
  "theory": "Paragraph on core theory related to topic + problem",
  "background": "Paragraph on background and prerequisites",
  "historicalContext": "Paragraph on historical context",
  "theoreticalFramework": "Paragraph on framework/model used",
  "principles": "Paragraph explaining main principles",
  "propertiesOfTopic": "Paragraph on properties relevant to solving the problem",
  "howItSolves": "Paragraph explaining how the solution addresses the problem",
  "relatedAns1": "Paragraph answering problem subsection 1",
  "relatedAns2": "Paragraph answering problem subsection 2",
  "differentApproachesAns": "Paragraph summarizing multiple approaches",
  "approachRelatedAns1": "Paragraph answering approach subsection 1",
  "approachRelatedAns2": "Paragraph answering approach subsection 2",
  "applicationRelatedAnsMain": "Paragraph on overall application",
  "applicationRelatedAns1": "Paragraph answering application subsection 1",
  "applicationRelatedAns2": "Paragraph answering application subsection 2",
  "applicationRelatedAns3": "Paragraph covering a third application subsection or real-world example.",
  "designAndAnalysisRelatedAns": "Paragraph describing the design approach and analysis outcomes for section 7.",
  "problemStatementBrief": "One concise paragraph (3-4 sentences) briefly summarising the problem statement for the introduction section.",
  "problemStatementSummary": "One short paragraph (2-3 sentences) summarising the problem statement at the start of section 4.",
  "conclusion": "Final conclusion paragraph",
  "references": "Exactly 5 numbered research-paper references in this format: 1. ...\n2. ...\n3. ...\n4. ...\n5. ..."
}

CRITICAL RULES:
- Use literal \\n where line separation is needed.
- Do not use markdown symbols.
- Keep content formal, clear, and academically meaningful.
- Keep all headings concise and title-friendly.
`;

        const aiResult = await generateWithFallback(m => m.generateContent(prompt), true);
        const aiData = parseAiJson(aiResult.response.text());

        function compactLine(text) {
            return String(text || "").replace(/\s+/g, " ").trim();
        }

        function ensureParagraphs(text) {
            const value = String(text || "").replace(/\r/g, "").trim();
            if (!value) return "";
            return value.replace(/\n{3,}/g, "\n\n");
        }

        function mergeTwoHeadings(a, b, fallbackA, fallbackB) {
            const first = compactLine(a || fallbackA || "");
            const second = compactLine(b || fallbackB || "");
            if (first && second && first.toLowerCase() !== second.toLowerCase()) {
                return `${first} / ${second}`;
            }
            return first || second;
        }

        function mergeTwoAnswers(titleA, ansA, titleB, ansB) {
            const t1 = compactLine(titleA || "Aspect 1");
            const t2 = compactLine(titleB || "Aspect 2");
            const a1 = ensureParagraphs(ansA || "");
            const a2 = ensureParagraphs(ansB || "");
            const chunks = [];
            if (a1) chunks.push(`${t1}: ${a1}`);
            if (a2) chunks.push(`${t2}: ${a2}`);
            return chunks.join("\n\n").trim();
        }

        function normalizeReferences(text) {
            const rawLines = ensureParagraphs(text)
                .split("\n")
                .map(line => line.trim())
                .filter(Boolean)
                .map(line => line.replace(/^\d+\.\s*/, ""));
            while (rawLines.length < 5) rawLines.push("Reference details to be updated.");
            return rawLines.slice(0, 5).map((line, idx) => `${idx + 1}. ${line}`).join("\n");
        }

        function extractSectionFromClass(classValue) {
            const cls = compactLine(classValue || "");
            const match = cls.match(/[-\s]([A-Za-z])$/);
            return match ? match[1].toUpperCase() : cls;
        }

        // ── Individual answer paragraphs are mapped directly in docData below ──

        const docData = {
            aatNo: compactLine(form.aatNo || form.type || ""),
            name: compactLine(form.name || ""),
            rollNo: compactLine(form.rollNo || ""),
            branch: compactLine(form.class || ""),
            section: compactLine(form.section || extractSectionFromClass(form.class || "")),
            semester: compactLine(form.semester || ""),
            courseCode: compactLine(form.courseCode || ""),
            faculty: compactLine(form.faculty || form.lecturerName || ""),

            topic: compactLine(topicInput || ps),
            problemStatement: compactLine(aiData.problemStatement || ps),

            // ── Section headings (split into numbered variants for template) ──
            principleRelatedTopic1: compactLine(aiData.principleTopic1 || "Core Principles"),
            principleRelatedTopic2: compactLine(aiData.principleTopic2 || "Solution Principles"),
            problemStatementRelatedTopic1: compactLine(aiData.problemTopic1 || "Root Cause Analysis"),
            problemStatementRelatedTopic2: compactLine(aiData.problemTopic2 || "Impact Dimensions"),
            relatedTopic1: compactLine(aiData.problemTopic1 || "Root Cause Analysis"),
            relatedTopic2: compactLine(aiData.problemTopic2 || "Impact Dimensions"),
            differentApproaches: compactLine(aiData.differentApproaches || "Comparative Solution Approaches"),
            approachRelatedTopic1: compactLine(aiData.approachTopic1 || "Primary Approach"),
            approachRelatedTopic2: compactLine(aiData.approachTopic2 || "Alternative Approach"),
            appraochRelated1: compactLine(aiData.approachTopic1 || "Primary Approach"),   // template typo kept
            approachRelated2: compactLine(aiData.approachTopic2 || "Alternative Approach"),
            application: compactLine(aiData.application || "Applications in Academic Environment"),
            applicationRelatedTopic1: compactLine(aiData.applicationTopic1 || "Operational Application"),
            applicationRelatedTopic2: compactLine(aiData.applicationTopic2 || "Strategic Application"),
            designAndAnalysis: compactLine(aiData.designAndAnalysis || "Design and Impact Analysis"),

            // ── Page numbers (template has pageNo1..pageNo21, all default to "1") ──
            pageNo1: "1", pageNo2: "1", pageNo3: "1", pageNo4: "1", pageNo5: "1",
            pageNo6: "1", pageNo7: "1", pageNo8: "1", pageNo9: "1", pageNo10: "1",
            pageNo11: "1", pageNo12: "1", pageNo13: "1", pageNo14: "1", pageNo15: "1",
            pageNo16: "1", pageNo17: "1", pageNo18: "1", pageNo19: "1", pageNo20: "1",
            pageNo21: "1",

            // ── Section content ──
            introduction: ensureParagraphs(aiData.introduction || ""),
            useInCampus: ensureParagraphs(aiData.useInCampus || ""),
            focus: ensureParagraphs(aiData.focus || ""),
            problemStatementBrief: ensureParagraphs(aiData.problemStatementBrief || aiData.problemStatement || ps),
            whyThisTechnique: ensureParagraphs(aiData.whyThisTechnique || ""),
            observation: ensureParagraphs(aiData.observation || ""),
            scope: ensureParagraphs(aiData.scope || ""),
            theory: ensureParagraphs(aiData.theory || ""),
            background: ensureParagraphs(aiData.background || ""),
            historicalContext: ensureParagraphs(aiData.historicalContext || ""),
            theoreticalFramework: ensureParagraphs(aiData.theoreticalFramework || ""),
            principles: ensureParagraphs(aiData.principles || ""),
            propertiesOfTopic: ensureParagraphs(aiData.propertiesOfTopic || ""),
            howItSolves: ensureParagraphs(aiData.howItSolves || ""),
            problemStatementSummary: ensureParagraphs(aiData.problemStatementSummary || aiData.problemStatement || ps),
            relatedAns1: ensureParagraphs(aiData.relatedAns1 || ""),
            relatedAns2: ensureParagraphs(aiData.relatedAns2 || ""),
            differentApproachesAns1: ensureParagraphs(aiData.differentApproachesAns || ""),
            approachRelatedAns1: ensureParagraphs(aiData.approachRelatedAns1 || ""),
            approachRelatedAns2: ensureParagraphs(aiData.approachRelatedAns2 || ""),
            applicationRelatedAns1: ensureParagraphs(aiData.applicationRelatedAns1 || ""),
            applicationRelatedAns2: ensureParagraphs(aiData.applicationRelatedAns2 || ""),
            applicationRelatedAns3: ensureParagraphs(aiData.applicationRelatedAns3 || ""),
            designAndAnalysisRelatedAns: ensureParagraphs(aiData.designAndAnalysisRelatedAns || ""),
            conclusion: ensureParagraphs(aiData.conclusion || ""),
            references: normalizeReferences(aiData.references || ""),

            courseTitle: compactLine(form.courseTitle || ""),
            couseTitle: compactLine(form.courseTitle || ""),
            date: formattedDate
        };

        const templateBuf = fs.readFileSync(
            path.join(__dirname, "assets", "CaseStudy_Template.docx"),
            "binary"
        );
        const zip = new PizZip(templateBuf);
        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true
        });
        doc.render(docData);
        const buf = doc.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });

        const safeName = (form.name || "Student").replace(/[^a-zA-Z0-9_]/g, "_");

        await setRedisJson(
            `case_study_result:${userId}`,
            {
                status: "done",
                fileName: `${safeName}_CaseStudy.docx`,
                docBase64: buf.toString("base64")
            },
            EFFECTIVE_CASE_STUDY_TTL
        );

        console.log(`[CaseStudy] Done for ${userId}`);
    } catch (err) {
        console.error(`[CaseStudy] Failed for ${userId}:`, err.message);
        await setRedisJson(
            `case_study_result:${userId}`,
            { status: "error", message: err.message },
            10 * 60
        );
    } finally {
        const n = await redis.decr("case_study_active");
        if (n < 0) await redis.set("case_study_active", "0");
        await redis.lrem("case_study_processing", 0, userId);
        console.log(`[CaseStudy] Cleanup success for ${userId}`);
        processNextCaseStudy().catch(err => console.error("[CaseStudy] next-job error:", err));
    }
}

app.get("/case-study-status/:userId", isAuthenticated, async (req, res) => {
    const { userId } = req.params;

    if (!isOwnUser(req, userId)) {
        console.warn(`[Auth] Case Study status forbidden for ${req.user.studentId} -> ${userId}`);
        return res.status(403).json({ error: "Forbidden" });
    }

    try {
        recoverStuckCaseStudy().catch(console.error);
        processNextCaseStudy().catch(console.error);
        const data = await getRedisJson(`case_study_result:${userId}`);
        if (!data) {
            const queuePos = await redis.lpos("case_study_queue", userId);
            const processingPos = await redis.lpos("case_study_processing", userId);

            if (queuePos !== null) {
                await setRedisJson(
                    `case_study_result:${userId}`,
                    { status: "waiting", position: queuePos + 1 },
                    EFFECTIVE_CASE_STUDY_TTL
                );
                return res.json({ status: "waiting", position: queuePos + 1 });
            }

            if (processingPos !== null) {
                await setRedisJson(
                    `case_study_result:${userId}`,
                    { status: "generating", startedAt: Date.now() },
                    EFFECTIVE_CASE_STUDY_TTL
                );
                return res.json({ status: "generating" });
            }

            return res.json({ status: "not_found" });
        }

        if (data.status === "waiting") {
            const queuePos = await redis.lpos("case_study_queue", userId);
            const processingPos = await redis.lpos("case_study_processing", userId);

            if (queuePos === null && processingPos === null) {
                console.warn(`[CaseStudyStatus] Re-queueing lost waiting job for ${userId}`);
                await redis.rpush("case_study_queue", userId);
                await refreshCaseStudyQueuePositions();
                processNextCaseStudy().catch(console.error);
                const repaired = await getRedisJson(`case_study_result:${userId}`);
                return res.json(repaired || { status: "waiting", position: 1 });
            }

            return res.json({
                ...data,
                position: queuePos === null ? data.position || 1 : queuePos + 1
            });
        }

        if (data.status === "generating") {
            const processingPos = await redis.lpos("case_study_processing", userId);
            if (processingPos === null) {
                console.warn(`[CaseStudyStatus] Generating job missing from processing for ${userId}, forcing recovery`);
                recoverStuckCaseStudy().catch(console.error);
                processNextCaseStudy().catch(console.error);
            }
        }

        if (data.status === "done") return res.json({ status: "done", fileName: data.fileName });
        return res.json(data);
    } catch (err) {
        console.error("[CaseStudy] Status error:", err);
        return res.status(500).json({ error: "Status check failed." });
    }
});

// ─── Case Study Download ──────────────────────────────────────────────────────
app.get("/download-case-study/:userId", isAuthenticated, async (req, res) => {
    const { userId } = req.params;

    if (!isOwnUser(req, userId)) {
        console.warn(`[Auth] Case Study download forbidden for ${req.user.studentId} -> ${userId}`);
        return res.status(403).send("Forbidden");
    }

    try {
        const data = await getRedisJson(`case_study_result:${userId}`);
        if (!data) return res.status(404).send("Document not found. Please generate again.");
        if (data.status !== "done") return res.status(400).send("Document not ready yet.");

        const buf      = Buffer.from(data.docBase64, "base64");
        const fileName = data.fileName || `${userId}_CaseStudy.docx`;

        res.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        res.send(buf);

        await redis.del(`case_study_result:${userId}`);
        await redis.del(`case_study_form:${userId}`);
        console.log(`[CaseStudy] Download success for ${userId}`);
    } catch (err) {
        console.error("[CaseStudy] Download error:", err);
        res.status(500).send("Download failed: " + err.message);
    }
});

// ─── Case Study Form Page ─────────────────────────────────────────────────────
app.get("/case-study", isAuthenticated, (req, res) => {
    res.render("case-study", { user: req.user });
});

// ─── Case Study Queue Waiting Page ────────────────────────────────────────────
app.get("/case-study-queue", isAuthenticated, (req, res) => {
    res.render("case-study-queue", { user: req.user });
});

// ─── Start Server ─────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 8080;

if (process.env.VERCEL !== "1") {
    app.listen(PORT, () => {
        console.log(`Server listening on port ${PORT}`);
    });
}

module.exports = app;