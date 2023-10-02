const path = require("path");
const mammoth = require("mammoth");
const { convert } = require("html-to-text");
const puppeteer = require("puppeteer");
const TurndownService = require("turndown");
const { Command } = require("commander");
const fs = require("fs");
const wordDoc = path.join(__dirname, "test.docx");

const exportWebsiteAsPdf = async (html, outputPath) => {
	// Create a browser instance
	const browser = await puppeteer.launch({
		headless: "new",
	});

	// Create a new page
	const page = await browser.newPage();

	await page.setContent(html, { waitUntil: "domcontentloaded" });

	// To reflect CSS used for screens instead of print
	await page.emulateMediaType("screen");

	// Download the PDF
	const PDF = await page.pdf({
		path: outputPath,
		margin: { top: "100px", right: "50px", bottom: "100px", left: "50px" },
		printBackground: true,
		format: "A4",
	});

	// Close the browser instance
	await browser.close();

	return PDF;
};

const createHTMLFromWord = async (filePathToWordDoc) => {
	try {
		const result = await mammoth.convertToHtml({ path: filePathToWordDoc });
		const html = result.value; // The generated HTML
		const messages = result.messages; // Any messages, such as warnings during conversion
		console.log({ html, messages });
		return html;
	} catch (error) {
		console.error(error);
	}
};

const convertHTMLToText = (html) => {
	return new Promise((resolve, reject) => {
		const text = convert(html, {
			wordwrap: 130,
		});
		resolve(text);
	});
};

const convertStringToFile = (string, fileName) => {
	return new Promise((resolve, reject) => {
		return fs.writeFile(fileName, string, (err) => {
			if (err) reject(err);
			console.log("Data has been written to file successfully.");
			resolve();
		});
	});
};

const createMarkdownFromHTML = (html) => {
	return new Promise((resolve, reject) => {
		const turndownService = new TurndownService();
		const markdown = turndownService.turndown(html);
		resolve(markdown);
	});
};

(async function () {
	const program = new Command();
	program
		.name("convert-docx")
		.description("CLI to convert docx files")
		.option("-f, --file <file>", "add the specified file", wordDoc)
		.option("-t, --file-type <type>", "type of file to convert: pdf, html, text, md", "html")
        .option("-o, --output-path <output>", "output file path", "./")
		.version("0.1.0");

	program.parse(process.argv);
    const html = await createHTMLFromWord(options.file);
    const fileName = path.basename(options.file);
    const filePath = path.join(options.outputPath, fileName);
    const completeOutputPath = filePath + "." + options.fileType;

    switch (options.fileType) {
        case pdf:
            await exportWebsiteAsPdf(html, completeOutputPath);
            break;
        case html:
            await convertStringToFile(html, completeOutputPath);
            break;
        case text:
            const text = await convertHTMLToText(html);
            await convertStringToFile(text, completeOutputPath);
            break;
        case md:
            const markdown = await createMarkdownFromHTML(html);
            await convertStringToFile(markdown, completeOutputPath);
            break;
        default:
            break;
    }
})();
