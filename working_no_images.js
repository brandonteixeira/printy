/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("downloadEmail").addEventListener("click", downloadEmailAsPDF);
  }
});

async function downloadEmailAsPDF() {
  const item = Office.context.mailbox.item;

  if (!item || !item.body) {
    console.error("Email item is not available.");
    return;
  }

  // Fetch email content in HTML format
  item.body.getAsync(Office.CoercionType.Html, async (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const emailContent = result.value;

      // Create a hidden div to hold the email content
      let tempDiv = document.createElement("div");
      tempDiv.innerHTML = emailContent;
      tempDiv.style.width = "800px";
      tempDiv.style.padding = "10px";
      document.body.appendChild(tempDiv);

      // Convert HTML to an image using html2canvas
      html2canvas(tempDiv, { scale: 2 }).then(async (canvas) => {
        const imgData = canvas.toDataURL("image/png");

        // Create a PDF from the image
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF("p", "mm", "a4");
        const imgWidth = 210; // A4 width in mm
        const imgHeight = (canvas.height * imgWidth) / canvas.width; // Maintain aspect ratio
        pdf.addImage(imgData, "PNG", 0, 0, imgWidth, imgHeight);

        // Cleanup temp div
        document.body.removeChild(tempDiv);

        // Fetch and append PDF attachments
        const mergedPdf = await appendPDFAttachments(pdf);
        const mergedPdfBytes = await mergedPdf.save();

        // Download final merged PDF
        const blob = new Blob([mergedPdfBytes], { type: "application/pdf" });
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = "email_with_attachments.pdf";
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
      });
    } else {
      console.error("Failed to retrieve email content:", result.error);
    }
  });
}

async function appendPDFAttachments(pdfDoc) {
  const { PDFDocument } = window.PDFLib;
  const mergedPdf = await PDFDocument.create();
  const mainPdf = await PDFDocument.load(pdfDoc.output("arraybuffer"));
  const mainPages = await mergedPdf.copyPages(mainPdf, mainPdf.getPageIndices());

  mainPages.forEach((page) => mergedPdf.addPage(page));

  // Fetch attachments from the email
  const item = Office.context.mailbox.item;
  if (item.attachments && item.attachments.length > 0) {
    for (let att of item.attachments) {
      if (att.contentType === "application/pdf") {
        const fileContent = await getAttachmentContent(att.id);
        if (fileContent) {
          const attachmentPdf = await PDFDocument.load(fileContent);
          const attachmentPages = await mergedPdf.copyPages(attachmentPdf, attachmentPdf.getPageIndices());

          attachmentPages.forEach((page) => mergedPdf.addPage(page));
        }
      }
    }
  }

  return mergedPdf;
}

// Helper function to get attachment content
function getAttachmentContent(attachmentId) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
      if (
        result.status === Office.AsyncResultStatus.Succeeded &&
        result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64
      ) {
        resolve(Uint8Array.from(atob(result.value.content), (c) => c.charCodeAt(0)));
      } else {
        console.error("Failed to get attachment content:", result.error);
        resolve(null);
      }
    });
  });
}
