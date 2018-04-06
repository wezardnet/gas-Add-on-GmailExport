/**
 * Gmail のメール本文と添付ファイルを Google ドライブにエクスポートするアドオン提供型スクリプト
 * 
 * @vesion 1.0.0
 * @author Kenji.Nakahara
 * @date: 2018/04/05
 * @update: 2018/04/06
 * @Deployment ID: AKfycbyBBT0K_k-lvzdZnV1kk9QfIh0vYjxaCFOBgCipLIY
 * @include libraries: 
 */

/**
 * マニフェスト 'onTriggerFunction' フィールドで指定されたアドオン起動トリガー
 * 
 * @param {Object} イベントオブジェクト
 * @return {Card} カードオブジェクト
 */
function buildAddOn(e){
	var accessToken = e.messageMetadata.accessToken;
	GmailApp.setCurrentMessageAccessToken(accessToken);

	var messageId = e.messageMetadata.messageId;
	var message = GmailApp.getMessageById(messageId);
	var subject = message.getSubject();					// 件名
	var from = message.getFrom();						// 送信元
	var attachments = message.getAttachments();			// 添付ファイル

	// UI を作る
	var buttonSet = CardService.newButtonSet();
	var exportButton = CardService.newTextButton()
		.setText('Drive Export')
		.setOnClickAction(CardService.newAction()
			.setFunctionName('driveExport')
			.setParameters({'messageId': messageId}));

	var card = CardService.newCardBuilder()
		.setHeader(CardService.newCardHeader()
			.setTitle(subject))
		.addSection(CardService.newCardSection()
			.addWidget(CardService.newKeyValue()
				.setTopLabel('送信元')
				.setContent(from))
			.addWidget(CardService.newKeyValue()
				.setTopLabel('添付ファイル')
				.setContent(attachments.length + ' 個'))
			.addWidget(exportButton))
		.build();

	return card;
}

/**
 * 指定されたメッセージ識別子に対応するメール本文と添付ファイルを Google ドライブに格納する
 * 
 * @param {Object} イベントオブジェクト
 */
function driveExport(e){
	var messageId = e.parameters['messageId'];
	var message = GmailApp.getMessageById(messageId);

	var driveFolder = 'Gmail Export';

	// Gmail Add-on https://developers.google.com/gmail/add-ons/
	// メール本文の PDF 化は https://ctrlq.org/code/19117-save-gmail-as-pdf を参考に実装
	var folders = DriveApp.getFoldersByName(driveFolder);
	var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(driveFolder);

	var html = '';
	html += 'From: ' + message.getFrom() + '<br />';
	html += 'To: ' + message.getTo() + '<br />';
	html += 'Date: ' + message.getDate() + '<br />';
	html += 'Subject: ' + message.getSubject() + '<br />';
	html += '<hr />';
	html += message.getBody().replace(/<img[^>]*>/g, '');
	html += '<hr />';

	var attachments = [];
	var atts = message.getAttachments();
	for ( var i = 0; i < atts.length; i++ ) attachments.push(atts[i]);

	if ( attachments.length > 0 ) {
		var footer = '<strong>添付ファイル</strong><ul>';
		for ( var i = 0; i < attachments.length; i++ ) {
			var file = folder.createFile(attachments[i]);
			footer += '<li><a href="' + file.getUrl() + '">' + file.getName() + '</a></li>';
		}
		html += footer + '</ul>';
	}

	var tempFile = DriveApp.createFile('temp.html', html, 'text/html');
	folder.createFile(tempFile.getAs('application/pdf')).setName(message.getSubject() + '.pdf');
//	tempFile.setTrashed(true);
	deleteFile(tempFile.getId());
}

/**
 * 指定したファイルを Google ドライブから完全削除する
 * 
 * @param {String} Google ドライブから削除するファイル ID
 */
function deleteFile(fileId){
	var token = ScriptApp.getOAuthToken();
	var response = UrlFetchApp.fetch(Utilities.formatString('https://www.googleapis.com/drive/v3/files/%s', fileId), {
		method: 'delete',
		headers: {
			'Authorization': 'Bearer ' + token
		},
		muteHttpExceptions: true
	});
}
