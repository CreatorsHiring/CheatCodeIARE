const express = require("express");
const path = require("path");
const session = require("express-session");
const dotenv = require("dotenv");
const { GoogleGenerativeAI, HarmCategory, HarmBlockThreshold } = require("@google/generative-ai");
const PptxGenJS = require("pptxgenjs");
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
                // Institute Name (Header Text) - Changed color to Blue since background bar is gone
                { text: { text: "INSTITUTE OF AERONAUTICAL ENGINEERING", options: { x: 0.2, y: 0.1, w: "60%", h: 0.4, fontSize: 18, color: IARE_BLUE, bold: true, align: "left" } } },
                // Footer Bar (Blue)
                { rect: { x: 0, y: "95%", w: "100%", h: 0.4, fill: { color: IARE_BLUE } } },
                { slideNumber: { x: "90%", y: "96%", fontSize: 10, color: "FFFFFF" } }
            ]
        });

        // Slide 1: Title Slide
        let slide1 = pres.addSlide({ masterName: "TITLE_MASTER" });

        slide1.addText("AAT PRESENTATION", { x: 0.5, y: 1.8, w: "90%", fontSize: 16, color: IARE_BLUE, bold: true, align: "center" });
        slide1.addText(content.title.toUpperCase(), { x: 0.5, y: 2.3, w: "90%", fontSize: 32, color: "000000", bold: true, align: "center" });

        // Student Details
        slide1.addText(`Topic: ${problemStatement}`, { x: 1.5, y: 3.5, w: "80%", fontSize: 16, color: TEXT_MAIN });
        slide1.addText(`Name: ${name}`, { x: 1.5, y: 4.0, w: "80%", fontSize: 16, color: TEXT_MAIN });
        slide1.addText(`Roll No: ${studentId}`, { x: 1.5, y: 4.5, w: "80%", fontSize: 16, color: TEXT_MAIN });
        slide1.addText(`Branch: ${department}`, { x: 1.5, y: 5.0, w: "80%", fontSize: 16, color: TEXT_MAIN });
        slide1.addText(`Subject: ${subject}`, { x: 1.5, y: 5.5, w: "80%", fontSize: 16, color: TEXT_MAIN });

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