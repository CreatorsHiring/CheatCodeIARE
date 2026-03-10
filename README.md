# 🎓 CheatCodeIARE

CheatCodeIARE is a web-based, AI-powered tool specifically designed to help students automatically generate academic presentations (PPTs) and comprehensive reports (DOCX) effortlessly. Simply by providing a problem statement or a set of questions, the app uses Google's Gemini AI to instantly generate structured, professionally formatted, and exam-ready documents.

---

## 🚀 Features

- **🔐 Student Login System:** Session-based authentication to store student details (Name, Roll No/Student ID) to auto-fill documents.
- **🧠 AI-Powered Content Generation:** Uses **Google Gemini API** (`gemini-flash-latest`) to generate highly accurate, academically appropriate content.
- **📊 Automated PPT Generation:** Instantly generates Presentation slides (.pptx) based on a topic or problem statement.
- **📝 Automated Report Generation:** Generates multi-page Word documents (.docx) addressing up to 10 academic questions.
- **🎨 Custom Templates & Master Slides:** Uses predefined PPT themes and DOCX templates to ensure professional and consistent formatting.
- **🌙 Student-Friendly UI:** Designed with a modern aesthetic optimized for ease of use.

---

## ⚙️ How PPT Generation Works

The core functionality of generating PowerPoint presentations involves combining AI response structuring and programmatic slide generation:

1. **User Input:** The student logs in and provides their `Department`, `Subject`, and a specific `Problem Statement` or topic.
2. **AI Content Structuring:**
   - The backend constructs a prompt instructing the Gemini AI to generate a presentation outline matching the problem statement.
   - The AI outputs a strictly formatted **JSON object** containing the presentation `title`, `introduction`, 5-7 `index` items, 5-7 `slides` (each with a heading and 3-5 bullet points), and a `conclusion`.
3. **Slide Assembly (`PptxGenJS`):** 
   - The application uses the `pptxgenjs` library to initialize a new PowerPoint presentation.
   - **Slide Masters:** Pre-defined master slides are applied to ensure consistent branding, utilizing custom assets (`image1.png` for title screens, `image2.png` for logos, and a colored footer).
   - **Data Injection:** The student's name, roll number, and subject are injected directly into the Title Slide.
   - **Slide Creation:** The app programmatically iterates through the AI-generated JSON content, generating individual slides for the Index, Introduction, Content (bullet points), Conclusion, and finally, a "Thank You" slide.
4. **Download:** The presentation is compiled into a `.pptx` buffer and sent to the user as a direct download formatted as `[StudentName]_AAT.pptx`.

---

## 📄 How Report Generation Works

Apart from presentations, CheatCodeIARE excels at generating detailed, exam-ready Word documents:

1. **User Input:** The student enters various academic details (Course Title, Semester, Regulation, etc.) along with up to **10 questions**.
2. **Batched AI Processing:**
   - To ensure high quality, the questions are processed through the Gemini API in batches.
   - Strict pacing and formatting rules are applied to generate plain text content that's safe to be injected into a DOCX format, explicitly avoiding unsupported markdown formats and translating structural components (like tables) into basic string representations.
3. **XML Transformation:** The generated markdown text is programmatically parsed and converted directly into Word-compatible XML string constructs (`<w:p>`, `<w:r>`, `<w:tbl>`, etc.). This ensures smooth paragraphs, structural tables, and proper line breaks.
4. **Template Patching (`docxtemplater` & `pizzip`):**
   - The app reads a base template file (`assets/ReportTemplate.docx`). *(A utility script `patch_template.js` is included to convert `{answer}` tags to `{@answer}` tags within the template to allow raw XML injection).*
   - Using `docxtemplater`, the app iterates through these placeholders and injects the raw XML content generated in the previous step, alongside the student's personal details.
5. **Download:** The generated document is zipped, compressed, and downloaded directly as `[StudentName]_Report.docx`.

---

## 🛠 Tech Stack

- **Frontend:** HTML, Vanilla CSS, JavaScript, EJS
- **Backend:** Node.js, Express.js
- **AI Integration:** `@google/generative-ai` (Gemini API)
- **Document APIs:** `pptxgenjs` (PPTX generation), `docxtemplater` & `pizzip` (DOCX template manipulation)
- **Sessions & Analytics:** `express-session`, `@vercel/analytics`
- **Hosting:** Configured for Vercel

---

## 💻 Getting Started (Local Development)

### 1. Requirements
- Node.js installed
- Google Gemini API Key

### 2. Setup
1. Clone the repository / download the files.
2. Install dependencies:
   ```bash
   npm install
   ```
3. Create a `.env` file in the root directory and add your API key:
   ```env
   GEMINI_API_KEY=your_gemini_api_key_here
   ```

### 3. Run the Application
Start the server:
```bash
node index.js
```
The application will be running at `http://localhost:8080`.

---
