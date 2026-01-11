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
const port = 3000;

// Middleware
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(session({
    secret: "cheatcodeiare_secret_key",
    resave: false,
    saveUninitialized: true
}));

// View Engine
app.set("view engine", "ejs");
app.set("views", path.join(__dirname, "views"));

// Static Files
app.use(express.static(path.join(__dirname, "public")));
app.use('/fa', express.static(__dirname + '/node_modules/@fortawesome/fontawesome-free'));

// Gemini Setup
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
const model = genAI.getGenerativeModel({
    model: "gemini-flash-latest",
    generationConfig: { responseMimeType: "application/json" },
    safetySettings: [
        { category: HarmCategory.HARM_CATEGORY_HARASSMENT, threshold: HarmBlockThreshold.BLOCK_NONE },
        { category: HarmCategory.HARM_CATEGORY_HATE_SPEECH, threshold: HarmBlockThreshold.BLOCK_NONE },
        { category: HarmCategory.HARM_CATEGORY_SEXUALLY_EXPLICIT, threshold: HarmBlockThreshold.BLOCK_NONE },
        { category: HarmCategory.HARM_CATEGORY_DANGEROUS_CONTENT, threshold: HarmBlockThreshold.BLOCK_NONE },
    ]
});

// Auth Middleware
const isAuthenticated = (req, res, next) => {
    if (req.session.user) {
        next();
    } else {
        res.redirect("/login");
    }
};

// Routes
app.get("/", (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.get("/login", (req, res) => {
    res.render("login");
});

app.post("/login", (req, res) => {
    const { name, studentId } = req.body;
    if (name && studentId) {
        req.session.user = { name, studentId };
        res.redirect("/ppt");
    } else {
        res.redirect("/login");
    }
});

app.get("/ppt", isAuthenticated, (req, res) => {
    res.render("ppt", { user: req.session.user });
});

app.get("/report", isAuthenticated, (req, res) => {
    res.render("report", { user: req.session.user });
});

const markdownToWordXML = (text) => {
    if (!text) return "";

    // 1. Escape XML special characters
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

        if (!trimmedLine) return; // Skip empty lines

        // Check for Table Start
        if (trimmedLine.startsWith("TABLE:")) {
            inTable = true;
            // Start Table XML
            // Basic table with grid borders
            finalXml += '<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>';
            return;
        }

        if (inTable) {
            // Check if line looks like a table row (has |)
            if (line.includes("|")) {
                finalXml += "<w:tr>";
                const cells = line.split("|");
                cells.forEach(cell => {
                    const cellContent = cell.trim();
                    // Each cell needs a paragraph run
                    finalXml += `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t>${cellContent}</w:t></w:r></w:p></w:tc>`;
                });
                finalXml += "</w:tr>";
            } else {
                // If line does not have |, assume table ended
                inTable = false;
                finalXml += "</w:tbl>";
                // Treat this line as normal text now
                finalXml += `<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t xml:space="preserve">${line}</w:t></w:r></w:p>`;
            }
        } else {
            // Normal Text Paragraph
            // Check for bold (simple heuristic, though user said no bold, we keep it just in case or for headings)
            // But strict rules say "No **bold**", so we just treat as plain text run.
            // We just wrap in standard paragraph
            finalXml += `<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t xml:space="preserve">${line}</w:t></w:r></w:p>`;
        }
    });

    if (inTable) {
        finalXml += "</w:tbl>"; // Close table if still open
    }

    return finalXml;
};

const batchGenerateAnswers = async (questions) => {
    const answers = new Array(10).fill("");
    const batchSize = 1; // Process 1 question at a time to ensure length/quality without timeout

    for (let i = 0; i < questions.length; i += batchSize) {
        const batch = questions.slice(i, i + batchSize);
        console.log(`Processing batch ${i / batchSize + 1}...`);

        const prompt = `
        You are an AI academic content generator for a student-focused AAT report system.

        The system will provide you with questions entered by the student.
        Your task is to generate detailed, exam-ready answers suitable for direct insertion into a Word document.

        Questions to answer:
        ${batch.map((q, idx) => `Question: ${q}`).join('\n')}

        ⚠️ The output will be injected into a DOCX template using a document-generation library.

        🔴 CRITICAL DOCX SAFETY RULES (MANDATORY)

        DO NOT output HTML, Markdown, or XML

        ❌ No <b>, <ul>, <li>, <table>

        ❌ No **bold**, __underline__, or Markdown tables

        DO NOT include raw formatting syntax

        ❌ No inline font sizes

        ❌ No CSS

        ❌ No special layout symbols

        Output must be PLAIN TEXT ONLY, compatible with Word paragraph runs.

        ✍️ FORMATTING RULES (DOCX-SAFE)

        Use capitalized headings instead of formatting tags
        Example:
        INTRODUCTION

        Separate sections using blank lines only

        Use hyphen-based bullets (-) only

        Maintain clean spacing between:

        Headings

        Paragraphs

        Bullet points

        📊 TABLE HANDLING (VERY IMPORTANT)

        If the question requires a table:

        ❌ DO NOT render a visual table
        ✅ Instead, output table data in this structured format:

        TABLE:
        Aspect | Element A | Element B
        Definition | ... | ...
        Advantages | ... | ...
        Limitations | ... | ...


        This data will be converted into a real Word table by backend logic.

        📄 CONTENT RULES

        DO NOT repeat the question

        Minimum length: one full page (400–600 words)

        Academic tone

        Clear logical flow

        Suitable for AAT / university evaluation

        🔠 FONT & STYLE NOTE

        The content will be rendered in 12pt font by the Word document template

        DO NOT mention font size in the output

        ❌ STRICTLY AVOID

        Emojis

        Special symbols

        Decorative characters

        Nested lists

        Overly long lines

        🎯 FINAL OUTPUT INSTRUCTION

        Generate ONLY the answer text, strictly following DOCX-safe rules.
        Do not include explanations or meta comments.

        Output valid JSON:
        {
            "answers": ["Answer String 1"]
        }
        `;

        try {
            const result = await model.generateContent(prompt);
            const responseText = result.response.text()
                .replace(/```json/g, "")
                .replace(/```/g, "")
                .trim();
            const json = JSON.parse(responseText);

            if (json.answers) {
                json.answers.forEach((ans, idx) => {
                    if (i + idx < 10) answers[i + idx] = ans;
                });
            }
        } catch (err) {
            console.error("Batch Error:", err);
            // Fallback or retry logic could go here
        }
    }
    return answers;
};

app.post("/generate-report", isAuthenticated, async (req, res) => {
    try {
        const formData = req.body;
        console.log("Form Data Received:", formData);

        // 1. Gather Questions
        const questions = [];
        for (let i = 1; i <= 10; i++) {
            questions.push(formData[`question${i}`]);
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

        res.setHeader("Content-Disposition", `attachment; filename="${formData.name}_Report.docx"`);
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
        res.send(buf);

    } catch (error) {
        console.error("Report Generation Error:", error);
        res.status(500).send("Error generating report. " + error.message);
    }
});

app.post("/generate-ppt", isAuthenticated, async (req, res) => {
    try {
        const { department, subject, problemStatement } = req.body;
        const { name, studentId } = req.session.user;

        // 1. Generate Content with Gemini
        const prompt = `
        You are an AI assistant generating structured PowerPoint slide content for an AAT presentation.
        Output valid JSON for the following schema:
        {
          "title": "String",
          "introduction": "String",
          "index": ["String", "String", "String", "String", "String"],
          "slides": [
              {
                "heading": "String",
                "bulletPoints": ["String", "String", "String"]
              }
          ],
          "conclusion": "String"
        }
        
        Topic: ${problemStatement}
        Rules:
        - Title must match the topic.
        - Generate 5-7 index items.
        - Generate 5-7 slides based on index.
        - Each slide has 3-5 bullet points.
        - Content must be academic and concise.
        `;

        const result = await model.generateContent(prompt);
        let responseText = result.response.text();
        console.log("Gemini Raw Response:", responseText); // Debug logging

        if (!responseText) {
            throw new Error("Gemini returned empty response. Likely safety block or model error.");
        }

        // Clean up potential markdown if JSON mode misses it (rare but possible)
        responseText = responseText.replace(/```json/g, "").replace(/```/g, "").trim();

        const content = JSON.parse(responseText);

        // 2. Generate PPT with PptxGenJS
        let pres = new PptxGenJS();

        // Theme Colors
        const BG_COLOR = "EDEDEE"; // Light Gray Background
        const IARE_BLUE = "003366";
        const TEXT_MAIN = "000000";
        const THANK_YOU_COLOR = "51237F";

        pres.layout = "LAYOUT_16x9";

        // MASTER 1: Title Slide & Thank You Slide (Image 1 at Top)
        pres.defineSlideMaster({
            title: "TITLE_MASTER",
            background: { color: BG_COLOR },
            objects: [
                // Image 1: Main Header/Banner at the top
                { image: { x: 0, y: 0, w: "100%", h: 1.5, path: path.join(__dirname, "assets", "image1.png") } }
            ]
        });

        // MASTER 2: Content Slides (Image 2 at Top Right)
        pres.defineSlideMaster({
            title: "CONTENT_MASTER",
            background: { color: BG_COLOR },
            objects: [
                // Image 2: Logo at Top Right
                { image: { x: "85%", y: 0.05, w: 1.2, h: 0.7, path: path.join(__dirname, "assets", "image2.png") } },
                // Footer Bar (Blue)
                { rect: { x: 0, y: "95%", w: "100%", h: 0.4, fill: { color: IARE_BLUE } } },
                { slideNumber: { x: "90%", y: "96%", fontSize: 10, color: "FFFFFF" } }
            ]
        });

        // Slide 1: Title Slide
        let slide1 = pres.addSlide({ masterName: "TITLE_MASTER" });

        const NEW_BLUE = "004170"; // User specific color matches AAT - TECH TALK

        // 1. AAT - TECH TALK (Center, 28px, #004170)
        slide1.addText("AAT - TECH TALK", { x: 0.5, y: 1.8, w: "90%", fontSize: 28, color: NEW_BLUE, bold: true, align: "center", fontFace: "Arial" });

        // 2. Topic - Problem Statement
        // "Topic" (28px Arial Headings #004170)
        // "Problem Statement" (24px)
        // "Little left aligned" -> We use x=1.0 to offset from left.

        const contentX = 1.0;
        const contentW = "80%";

        slide1.addText(
            [
                { text: "Topic - ", options: { fontSize: 28, color: NEW_BLUE, fontFace: "Arial", bold: true } },
                { text: `${problemStatement}`, options: { fontSize: 24, color: NEW_BLUE } }
            ],
            { x: contentX, y: 2.5, w: contentW, align: "left" }
        );

        // 3. Student Details (Name, Roll No, Branch, Subject)
        // Labels: 18px, #004170. Values: 18px (Default Black)
        // Left aligned with Topic (x=1.0)

        let startY = 3.5;
        const details = [
            { label: "Name", value: name },
            { label: "Roll No", value: studentId },
            { label: "Branch", value: department },
            { label: "Subject", value: subject }
        ];

        details.forEach((item, i) => {
            slide1.addText(
                [
                    { text: `${item.label}: `, options: { fontSize: 18, color: NEW_BLUE, bold: true } },
                    { text: `${item.value}`, options: { fontSize: 18, color: "000000" } }
                ],
                { x: contentX, y: startY + (i * 0.5), w: contentW, align: "left" }
            );
        });

        // Slide 2: Index
        let slideIndex = pres.addSlide({ masterName: "CONTENT_MASTER" });
        slideIndex.addText("INDEX", { x: 0.5, y: 1.0, w: "90%", fontSize: 24, color: IARE_BLUE, bold: true });

        content.index.forEach((item, idx) => {
            slideIndex.addText(`${idx + 1}. ${item}`, { x: 1, y: 2.0 + (idx * 0.5), w: "80%", fontSize: 18, color: TEXT_MAIN });
        });

        // Slide 3: Introduction
        let slideIntro = pres.addSlide({ masterName: "CONTENT_MASTER" });
        slideIntro.addText("INTRODUCTION", { x: 0.5, y: 1.0, w: "90%", fontSize: 24, color: IARE_BLUE, bold: true });
        slideIntro.addText(content.introduction, { x: 0.5, y: 2.0, w: "90%", fontSize: 18, color: TEXT_MAIN });

        // Content Slides
        content.slides.forEach(slideData => {
            let s = pres.addSlide({ masterName: "CONTENT_MASTER" });
            s.addText(slideData.heading.toUpperCase(), { x: 0.5, y: 1.0, w: "90%", fontSize: 24, color: IARE_BLUE, bold: true });

            let bulletItems = slideData.bulletPoints.map(bp => {
                return { text: bp, options: { bullet: true, fontSize: 18, color: TEXT_MAIN } };
            });

            s.addText(bulletItems, { x: 0.5, y: 2.0, w: "90%", h: "60%" });
        });

        // Conclusion Content Slide (Standard Content Master)
        let slideConc = pres.addSlide({ masterName: "CONTENT_MASTER" });
        slideConc.addText("CONCLUSION", { x: 0.5, y: 1.0, w: "90%", fontSize: 24, color: IARE_BLUE, bold: true });
        slideConc.addText(content.conclusion, { x: 0.5, y: 2.0, w: "90%", fontSize: 18, color: TEXT_MAIN });

        // Final Slide: Thank You (Uses TITLE_MASTER for image1 at top)
        let slideLast = pres.addSlide({ masterName: "TITLE_MASTER" });
        slideLast.addText("THANK YOU", { x: 0.5, y: 1.5, w: "90%", h: "50%", fontSize: 48, color: THANK_YOU_COLOR, bold: true, align: "center", valign: "middle" });

        // Generate and stream response
        const buffer = await pres.write("nodebuffer");

        res.setHeader("Content-Disposition", `attachment; filename="${name}_AAT.pptx"`);
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
        res.send(buffer);

    } catch (error) {
        console.error(error);
        res.status(500).send("Error generating PPT. Please check server logs.");
    }
});

app.listen(port, () => {
    console.log(`Server listening on port ${port}`);
});