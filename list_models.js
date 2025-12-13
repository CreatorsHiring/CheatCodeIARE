const { GoogleGenerativeAI } = require("@google/generative-ai");
const dotenv = require("dotenv");
dotenv.config();

async function listModels() {
    const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
    try {
        const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" }); // Dummy init to get client
        // The SDK doesn't have a direct listModels method on the client in some versions, 
        // but we can try to use the API directly or check documentation.
        // Actually, checking the error message suggested calling ListModels.

        // Let's try to hit the API directly using fetch to be sure, as SDK encapsulation might hide details.
        const key = process.env.GEMINI_API_KEY;
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${key}`);
        const data = await response.json();

        if (data.models) {
            console.log("Available Models:");
            data.models.forEach(m => {
                if (m.supportedGenerationMethods.includes("generateContent")) {
                    console.log(`- ${m.name}`);
                }
            });
        } else {
            console.log("No models found or error:", data);
        }

    } catch (error) {
        console.error("Error listing models:", error);
    }
}

listModels();
