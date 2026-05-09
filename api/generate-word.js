import {
  Document,
  Packer,
  Paragraph,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
  TextRun,
  PageBreak,
} from "docx";
import { put } from "@vercel/blob";

export default async function handler(req, res) {
  // CORS headers (for testing from anywhere)
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({
      status: "error",
      message: "Only POST method allowed",
    });
  }

  try {
    const body = req.body || {};
    const title = body.title || "Untitled Document";
    const organization = body.organization || "";
    const author = body.author || "";
    const sections = body.sections || [];

    const children = [];

    // ========== TITLE ==========
    children.push(
      new Paragraph({
        text: title,
        heading: HeadingLevel.TITLE,
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
      })
    );

    // ========== METADATA BLOCK ==========
    if (organization) {
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Organization: ", bold: true }),
            new TextRun({ text: organization }),
          ],
        })
      );
    }
    if (author) {
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Prepared By: ", bold: true }),
            new TextRun({ text: author }),
          ],
        })
      );
    }
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "Date: ", bold: true }),
          new TextRun({
            text: new Date().toLocaleDateString("en-US", {
              year: "numeric",
              month: "long",
              day: "numeric",
            }),
          }),
        ],
      })
    );
    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "Version: ", bold: true }),
          new TextRun({ text: "1.0" }),
        ],
        spacing: { after: 400 },
      })
    );

    // ========== TABLE OF CONTENTS (if many sections) ==========
    if (sections.length > 3) {
      children.push(
        new Paragraph({
          text: "Table of Contents",
          heading: HeadingLevel.HEADING_1,
        })
      );
      sections.forEach((section, i) => {
        children.push(
          new Paragraph({
            text: `${i + 1}. ${section.heading || `Section ${i + 1}`}`,
            numbering: { reference: "default-numbering", level: 0 },
          })
        );
      });
      children.push(
        new Paragraph({
          children: [new PageBreak()],
        })
      );
    }

    // ========== SECTIONS ==========
    sections.forEach((section, i) => {
      const heading = section.heading || `Section ${i + 1}`;
      const contentBlocks = section.content_blocks || [];

      // Section heading
      children.push(
        new Paragraph({
          text: `${i + 1}. ${heading}`,
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 300, after: 200 },
        })
      );

      // Content blocks
      contentBlocks.forEach((block) => {
        const blockType = block.type || "paragraph";

        if (blockType === "paragraph") {
          children.push(
            new Paragraph({
              text: block.text || "",
              spacing: { after: 150 },
            })
          );
        } else if (blockType === "subheading") {
          children.push(
            new Paragraph({
              text: block.text || "",
              heading:
                block.level === 3
                  ? HeadingLevel.HEADING_3
                  : HeadingLevel.HEADING_2,
              spacing: { before: 200, after: 100 },
            })
          );
        } else if (blockType === "bullet_list") {
          (block.items || []).forEach((item) => {
            children.push(
              new Paragraph({
                text: item,
                bullet: { level: 0 },
              })
            );
          });
        } else if (blockType === "numbered_list") {
          (block.items || []).forEach((item, idx) => {
            children.push(
              new Paragraph({
                text: `${idx + 1}. ${item}`,
              })
            );
          });
        } else if (blockType === "table") {
          const headers = block.headers || [];
          const rows = block.rows || [];

          if (headers.length > 0 && rows.length > 0) {
            const tableRows = [];

            // Header row
            tableRows.push(
              new TableRow({
                tableHeader: true,
                children: headers.map(
                  (h) =>
                    new TableCell({
                      width: {
                        size: 100 / headers.length,
                        type: WidthType.PERCENTAGE,
                      },
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({ text: String(h), bold: true }),
                          ],
                        }),
                      ],
                    })
                ),
              })
            );

            // Data rows
            rows.forEach((row) => {
              tableRows.push(
                new TableRow({
                  children: row.map(
                    (cell) =>
                      new TableCell({
                        children: [
                          new Paragraph({ text: String(cell) }),
                        ],
                      })
                  ),
                })
              );
            });

            children.push(
              new Table({
                rows: tableRows,
                width: { size: 100, type: WidthType.PERCENTAGE },
              })
            );
            children.push(new Paragraph({ text: "" })); // spacing after table
          }
        } else if (blockType === "quote") {
          children.push(
            new Paragraph({
              children: [
                new TextRun({ text: block.text || "", italics: true }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: { before: 200, after: 200 },
            })
          );
        }
      });
    });

    // ========== FOOTER ==========
    children.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `Generated on ${new Date().toLocaleString("en-US")}`,
            italics: true,
            size: 18,
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 600 },
      })
    );

    // ========== BUILD DOCUMENT ==========
    const doc = new Document({
      creator: "VKT Word Generator",
      title: title,
      styles: {
        default: {
          document: {
            run: {
              font: "Calibri",
              size: 22, // 11pt
            },
          },
        },
      },
      sections: [{ children }],
    });

    const buffer = await Packer.toBuffer(doc);

    // ========== UPLOAD TO VERCEL BLOB ==========
    const safeFilename = title.replace(/[^a-zA-Z0-9-_]/g, "_");
    const uniqueId = Math.random().toString(36).substring(2, 10);
    const blobName = `${safeFilename}_${uniqueId}.docx`;

    const blob = await put(blobName, buffer, {
      access: "public",
      contentType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });

    return res.status(200).json({
      status: "success",
      document_url: blob.url,
      file_name: blobName,
      message: `Word document '${title}' has been created successfully. Click the link to download.`,
    });
  } catch (error) {
    console.error("Error generating document:", error);
    return res.status(500).json({
      status: "error",
      message: "Failed to generate document",
      error: error.message,
    });
  }
}
