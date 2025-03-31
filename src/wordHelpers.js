import { readFileSync, writeFileSync } from "node:fs";
import {
	AlignmentType,
	Paragraph,
	PatchType,
	TextRun,
	patchDocument,
} from "docx";

export function readTemplateFile(track, level) {
	if (track === "IT-Projektmanagement" || track === "Deep Learning") {
		// IT Projektmanagement and Deep Learning don't differentiate between beginner and advanced levels
		return readFileSync(`${process.cwd()}/templates/${track}.docx`);
	}
	return readFileSync(`${process.cwd()}/templates/${track} ${level}.docx`);
}

export async function fillTemplate(document, data) {
	let workshopsIntro = [];
	let workshopParagraphs = [];

	if (data.workshops.length > 0) {
		// Add intro paragraph if workshops were attended
		workshopsIntro.push(
			new TextRun({
				text: `${data.firstName} also took part in workshops with the following companies: `,
				font: "Inter Tight",
				size: "10pt",
				color: "000000",
			})
		);
	}

	if (data.workshops.length > 0) {
		// Add bullet point for each workshop
		for (let workshop of data.workshops) {
			workshopParagraphs.push(
				new Paragraph({
					bullet: { level: 1 },
					children: [
						new TextRun({
							text: workshop,
							font: "Inter Tight",
							size: "10pt",
							color: "000000",
						}),
					],
				})
			);
		}
	}

	return await patchDocument(document, {
		patches: {
			name: {
				type: PatchType.PARAGRAPH,
				children: [
					new TextRun({
						text: data.name,
						font: "Urbanist",
						size: "18pt",
						color: "1e3770",
						bold: true,
					}),
				],
			},
			// Show only the workshops paragraph if workshops were attended
			...(data.workshops?.length > 0
				? {
						workshops: {
							type: PatchType.PARAGRAPH,
							children: workshopsIntro,
						},
					}
				: {
						workshops: {
							type: PatchType.PARAGRAPH,
							children: [],
						},
					}),
			// Show only the workshops list if workshops were attended
			...(data.workshops?.length > 0
				? {
						workshopsList: {
							type: PatchType.DOCUMENT,
							children: workshopParagraphs,
						},
					}
				: {
						workshopsList: {
							type: PatchType.DOCUMENT,
							children: [],
						},
					}),
		},
	});
}

export async function generateCertificate(document, fileName) {
	const wordPath = `${process.cwd()}/certificates/docx/${fileName}.docx`;
	const pdfPath = `${process.cwd()}/certificates/pdf/${fileName}.pdf`;

	writeFileSync(wordPath, document);
}

export async function generateCommentFile(submission) {
	const commentTemplatePath = `${process.cwd()}/templates/Kommentar.docx`;
	const commentPath = `${process.cwd()}/certificates/comment/${
		submission.fileName
	}_comment.docx`;
	const pdfPath = `${process.cwd()}/certificates/pdf/${
		submission.fileName
	}_comment.pdf`;

	const commentTemplate = readFileSync(commentTemplatePath);

	// Create an array to hold all the TextRun objects
	const commentChildren = [];

	if (submission.comment) {
		const commentLines = submission.comment.split("\n");

		// Add each line with proper line breaks
		commentLines.forEach((line, index) => {
			// Add the text content
			commentChildren.push(
				new TextRun({
					text: line,
					alignment: AlignmentType.LEFT,
					font: "Fira Code Light",
					size: "10pt",
					color: "000000",
				})
			);

			// Add a line break if not the last line
			if (index < commentLines.length - 1) {
				commentChildren.push(new TextRun({ break: 1 }));
			}
		});
	}

	const commentDocument = await patchDocument(commentTemplate, {
		patches: {
			comment: {
				type: PatchType.PARAGRAPH,
				alignment: AlignmentType.LEFT,
				children: commentChildren.length > 0 ? commentChildren : [],
			},
		},
	});

	writeFileSync(commentPath, commentDocument);
}
