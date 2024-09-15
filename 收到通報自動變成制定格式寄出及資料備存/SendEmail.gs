function autoemail() {
    
    var mailImages = {};
    var emailBody = '<p>Here are the images:</p>';
    
    // 提取圖片url
    var imageUrls = 附件照片.split(', ');

    for (var i = 0; i < imageUrls.length; i++) {
        var imageUrl = imageUrls[i].trim();
        if (imageUrl) {
            var imageId = imageUrl.split('=')[1]; // URL中提取ID
            if (imageId) {
                var image = DriveApp.getFileById(imageId).getBlob();
                var cidKey = 'image' + i;
                mailImages[cidKey] = image; // 将图片添加到邮件图片对象
                // 限制圖片顯示大小，wide = 400px，高度automation
                emailBody += '<img src="cid:' + cidKey + '" alt="Image ' + i + '" style="max-width:400px; height:auto;"><br>';
            }
        }
    }
    
    // add pic(If it have)
    if (圖片的ID) {
        var originalImage = DriveApp.getFileById(圖片的ID).getAs("image/png");
        mailImages['logo100'] = originalImage; 
        
        emailBody += '<img src="cid:logo100" alt="Original Image" style="max-width:400px; height:auto;"><br>';
    }
    
    
    MailApp.sendEmail({
        to: "server@XXX.com.tw",
        subject: "(\/)@@(\/)事故通報-" + 員編 + "_" + 姓名 + "_" + 單位名稱,
        htmlBody: emailBody,
        inlineImages: mailImages,
        attachments: [要寄信的pdf]
    });
}
