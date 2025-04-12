/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, PDFLib, jspdf */
import { Email } from "../models/email.js";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initializeApp();
  }
});

function initializeApp() {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("downloadEmail").addEventListener("click", downloadEmailAsPDF);
}

async function downloadEmailAsPDF() {
  debugger;
  const email = new Email(Office.context.mailbox.item);
  const item = Office.context.mailbox.item;
  if (!item || !item.body) {
    return console.error("Email item is not available.");
  }

  // Convert email body to html
  item.body.getAsync(Office.CoercionType.Html, async (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      return console.error("Failed to retrieve email content:", result.error);
    }

    let emailContent = result.value;

    // Directly convert HTML to image
    const emailImage = await convertHTMLToImage(emailContent);
    if (!emailImage) return;

    const pdf = await createPDF(emailImage);
    const finalPdf = await appendAttachments(pdf);
    downloadPDF(finalPdf);
  });
}

async function convertHTMLToImage(htmlContent) {
  const tempDiv = document.createElement("div");
  tempDiv.innerHTML = htmlContent;
  Object.assign(tempDiv.style, { width: "800px", padding: "10px", position: "absolute", left: "-9999px" });
  document.body.appendChild(tempDiv);

  try {
    const canvas = await html2canvas(tempDiv, { scale: 2 });
    return canvas;
  } catch (error) {
    console.error("Error converting HTML to image:", error);
    return null;
  } finally {
    document.body.removeChild(tempDiv);
  }
}

async function createPDF(canvas) {
  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF("p", "mm", "a4");
  const imgData = canvas.toDataURL("image/png");
  const pdfWidth = pdf.internal.pageSize.getWidth();
  const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
  pdf.addImage(imgData, "PNG", 0, 0, pdfWidth, pdfHeight);
  return pdf;
}

async function appendAttachments(pdfDoc) {
  const { PDFDocument } = PDFLib;
  const mergedPdf = await PDFDocument.create();
  const mainPdf = await PDFDocument.load(pdfDoc.output("arraybuffer"));
  const mainPages = await mergedPdf.copyPages(mainPdf, mainPdf.getPageIndices());
  mainPages.forEach((page) => mergedPdf.addPage(page));

  const attachments = Office.context.mailbox.item.attachments || [];

  for (let att of attachments) {
    if (att.contentType === "application/pdf") {
      // Handle PDF attachments
      const fileContent = await getAttachmentContent(att.id);
      if (fileContent) {
        const attachmentPdf = await PDFDocument.load(fileContent);
        const attachmentPages = await mergedPdf.copyPages(attachmentPdf, attachmentPdf.getPageIndices());
        attachmentPages.forEach((page) => mergedPdf.addPage(page));
      }
    } else if (att.contentType.startsWith("image/")) {
      // Handle image attachments
      const fileContent = await getAttachmentContent(att.id);
      if (fileContent) {
        // Convert binary to Base64 string
        const binary = String.fromCharCode.apply(null, new Uint8Array(fileContent));
        const imgData = window.btoa(binary);

        // Convert base64 to image bytes
        const imageBytes = Uint8Array.from(atob(imgData), (c) => c.charCodeAt(0));

        // Embed the image into the PDF
        const image = await mergedPdf.embedPng(imageBytes);

        // Get the image dimensions
        const { width, height } = image.size();

        // Create a new page and append the image to it directly
        mergedPdf.addPage().drawImage(image, {
          x: (210 - width) / 2, // Center the image horizontally on an A4 size (210mm)
          y: (297 - height) / 2, // Center the image vertically on an A4 size (297mm)
          width: width,
          height: height,
        });
      }
    }
  }

  return mergedPdf;
}

function getAttachmentContent(attachmentId) {
  return new Promise((resolve) => {
    Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
          // If it's a Base64-encoded attachment (either PDF or image)
          resolve(Uint8Array.from(atob(result.value.content), (c) => c.charCodeAt(0)));
        } else {
          console.error("Unsupported attachment content format:", result.value.format);
          resolve(null);
        }
      } else {
        console.error("Failed to get attachment content:", result.error);
        resolve(null);
      }
    });
  });
}

function downloadPDF(mergedPdf) {
  mergedPdf.save().then((mergedPdfBytes) => {
    const blob = new Blob([mergedPdfBytes], { type: "application/pdf" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "email_with_attachments.pdf";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  });
}
