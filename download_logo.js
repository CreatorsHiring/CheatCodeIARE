const fs = require('fs');
const https = require('https');
const path = require('path');

const fileUrl = "https://placehold.co/400x400/003366/FFFFFF/png?text=IARE+LOGO";
const filePath = path.join(__dirname, 'assets', 'iare_logo.png');

const file = fs.createWriteStream(filePath);

https.get(fileUrl, function (response) {
    response.pipe(file);

    file.on('finish', () => {
        file.close();
        console.log("Download Completed");
    });
}).on('error', (err) => {
    fs.unlink(filePath, () => { }); // Delete the file async. (But we don't check the result) - just cleanup
    console.error("Error downloading file:", err.message);
});
