/**
 * This is a method for converting all pages in a PDF file to PNG images.
 * PNG images are returned as BlobSource[].
 * IMPORTANT: This method uses Drive API. Please enable Drive API at Advanced Google services.
 *
 * @param {Blob} blob Blob of PDF file.
 * @return {BlobSource[]} PNG blobs.
 */
async function convertPDFToPNG_(blob) {
  // Convert PDF to PNG images.
  const cdnjs = "https://cdn.jsdelivr.net/npm/pdf-lib/dist/pdf-lib.min.js";
  eval(UrlFetchApp.fetch(cdnjs).getContentText()); // Load pdf-lib
  const setTimeout = function (f, t) {
    // Overwrite setTimeout with Google Apps Script.
    Utilities.sleep(t);
    return f();
  };
  const data = new Uint8Array(blob.getBytes());
  const pdfData = await PDFLib.PDFDocument.load(data);
  const pageLength = pdfData.getPageCount();
  console.log(`Total pages: ${pageLength}`);
  const obj = { imageBlobs: [], fileIds: [] };
  for (let i = 0; i < pageLength; i++) {
    console.log(`Processing page: ${i + 1}`);
    const pdfDoc = await PDFLib.PDFDocument.create();
    const [page] = await pdfDoc.copyPages(pdfData, [i]);
    pdfDoc.addPage(page);
    const bytes = await pdfDoc.save();
    const blob = Utilities.newBlob(
      [...new Int8Array(bytes)],
      MimeType.PDF,
      `sample${i + 1}.pdf`
    );
    const id = DriveApp.createFile(blob).getId();
    Utilities.sleep(3000); // This is used for preparing the thumbnail of the created file.
    const link = Drive.Files.get(id, { fields: "thumbnailLink" }).thumbnailLink;
    if (!link) {
      throw new Error(
        "In this case, please increase the value of 3000 in Utilities.sleep(3000), and test it again."
      );
    }
    const imageBlob = UrlFetchApp.fetch(link.replace(/\=s\d*/, "=s1000"))
      .getBlob()
      .setName(店鋪代號 + '_' + 店鋪名稱 + '-臨場服務通知單-' + 月 + 日);     ////和PDF檔名不同(一個是底線, 一個是中間短分隔線), 之後用檔案名稱抓需要的檔案
    obj.imageBlobs.push(imageBlob);
    obj.fileIds.push(id);
  }
  obj.fileIds.forEach((id) => DriveApp.getFileById(id).setTrashed(true));
  return obj.imageBlobs;
}

// Please run this function.
async function main_pdf_to_pic() {
  // Retrieve PDF data.
  const fileId = 通知單ID; // Please set file ID of PDF file on Google Drive.
  const blob = DriveApp.getFileById(fileId).getBlob();

  // const url = "###url1###"; // Please set the direct link of the PDF file.
  // const blob = UrlFetchApp.fetch(url).getBlob();

  // Use a method for converting all pages in a PDF file to PNG images.
  const imageBlobs = await convertPDFToPNG_(blob);

  // As a sample, create PNG images as PNG files.
  imageBlobs.forEach((b) => dsfolder.createFile(b));

  // As another sample, create a zip file including the converted PNG images.
  // const zip = Utilities.zip(imageBlobs, "sample.zip");
  // DriveApp.createFile(zip);
}