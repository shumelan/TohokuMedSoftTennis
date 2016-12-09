/**
Timestamp: 2016-12-07

 * フォームの回答を自動で返信
 * 必要な設定: 
 * 回答先シート名を”フォームの回答1”---> 題名, イベント名(例: “北医体”)に変更
 * フォームに”メールアドレス”を回答する欄をつくる. 
 * ただし, 必須でなくても可, 部員名簿(member sheet (ms))上のアドレスが指定される
 * 回答者がメールアドレスを入力した場合, 部員名簿のメールアドレス欄が修正される

 * #####: ID
 * .....: name
 * ?????: host mail address 
 * !!!!!: sub address
 **/

function autoReply2(){
    
    // 設定
    var sem     =   "early";	// 半期 "early" or "late"
    var year    =   "2015";		// 年度
    
    var formUrl="";

    // 部員名簿のID
    // 部員名簿のURLの'spreadsheets/d/' から '/edit#gid=0' に挟まれた文字列のこと
    var mssID   =   "#####";
    
	// subject
	
	var subject = "【回答確認】";      //　メール件名: 　"【回答確認】"の後にシート名が追加される
	 

	// header; body; footer
	// HTMLメールが送信できない時に, 送信されるテキストメールの文面
	var header
	= "回答, ありがとうございます.\n"
	+ "回答に変更などあれば, 再度,フォームを送信してください.\n";            // ヘッダー
	
	// text mail
	var body
	= "============================\n";
	
	var sep
	= "============================\n";

	var footer
	= "\n\n"
	+ "************************************************\n"
	+ "......軟式テニス部\n"
	+ "?????@gmail.com\n"
	+ "************************************************";       //フッター
	
	// HTML mail 
	
	var htmlMail 	= 	'<!DOCTYPE html>'
					+	'<html>'
					+	'<head>'
					+		'<meta charset="UTF-8">'
					+		'<style type="text/css">'
					+			'.head{font-size:14px;font-family:"ＭＳ ゴシック",sans-serif;}'
					+			'table, tbody, tfoot,thead,tr,th,td{background: none;margin: 0;padding:0;border:0;font-size:12px;font:inherit;vertical-align:middle;}'
					+			'table{border-collapse:collapse;border-spacing:0;color:#333;font-family:"ＭＳ ゴシック",sans-serif; width:100%;border-collapse:collapse;border-spacing:0;}'
					+			'td,th{border:1px solid transparent;height:40px;transition:all 0.3s;}'
					+			'th{background: #DFDFDF;font-weight: bold;}'
					+			'td{background: #FAFAFA;text-align: center;}'
					+			'.data tr:nth-child(even) td {background: #F1F1F1;}'
					+			'.data tr:nth-child(odd) td {background: #FEFEFE;}'
					+			'.data tfoot tr td{text-align: left;}'
					+			'.sign td{border-bottom: 1px solid #333333;background: white; text-align: left; font-size: 14px;}'
					+		'</style>'
					+	'</head>'
					+	'<body>'
					+		'<div class="head">'
					+			'<p>回答, ありがとうございます.</p>'
					+			'<p>変更があれば, <a href="{{url}}">再度送信</a>してください.</p>'
					+		'</div><br/>'
					+		'<div  class="data">'
					+			'<table>'
					+				'<thead>'
					+					'<tr><th colspan=2>回答</th></tr>'
					+				'</thead>'
					+				'<tfoot>'
					+					'<tr><td colspan=2>送信日時: {{time}}</td></tr>'
					+				'</tfoot>'
					+				'<tbody>{{data}}</tbody>'
					+			'</table>'
					+			'<br/>'
					+		'</div>'
					+		'<div class="sign">'
					+			'<table class="sign">'
					+				'<tbody>'
					+					'<tr><td></td></tr>'
					+					'<tr><td>.....ソフトテニス部<br/><a href="mailto:?????@gmail.com">?????@gmail.com</a></td></tr>'
					+				'</tbody>'
					+			'</table>'
					+		'</div>'
					+	'</body>'
					+	'</html>';
	
	var row 		= 	'<tr>'
					+		'<td>{{col_name}}:</td>'
					+		'<td>{{col_value}}</td>'
					+	'</tr>';
	
	
	var hostml 	= 	"?????@gmail.com"; //執行部メール 
		
	// cc; bcc; reply; to
	var cc      =   "";   // Cc:
	var acc     =   "!!!!!@gmail.com";
	var bcc     =   acc;                         // Bcc:
	var reply   =   hostml;               // Reply-To:
	var to      =   "";                            // To:
	var newAdd  =   "";
	var id      =   "";
	
	// 回答集計したスプレッドシート	
	// 入力カラム名の指定
	var MAIL_COL_NAME   =   'メールアドレス';	//フォームの記載がメールアドレスでなければ更新
	var TIMESTAMP_LABEL =   'タイムスタンプ';
	var ID_COL_NAME     =   '学籍番号' ;
	var NAME_LABEL      =   '名前';
	
	try{
		// スプレッドシートの操作
		var sh      =   SpreadsheetApp.getActiveSheet();
		var shName  =   sh.getName();
		var rows    =   sh.getLastRow();
		var cols    =   sh.getLastColumn();
		var rg      =   sh.getDataRange();
		Logger.log("rows="+rows+" cols="+cols);
		
		//var formUrl= sh.getFormUrl();
		
		
		//htmlMail = htmlMail.replace("{{url}}", formUrl, "g");
		
		
		// メール件名・本文作成と送信先メールアドレス取得
		for (var j = 1; j <= cols; j++ ) {
			var col_name  = rg.getCell(1, j).getValue();    // カラム名
			var col_value = rg.getCell(rows, j).getValue(); // 入力値
			
			if ( col_name === TIMESTAMP_LABEL ) {　//タイムスタンプ
				col_value = Utilities.formatDate(col_value, "JST", "yyyy/MM/dd HH:mm:ss");
				sep += "送信日時 : "+col_value;
				
				
				htmlMail = htmlMail.replace("{{time}}", col_value, "g");
				
			}
			else if (col_name === ID_COL_NAME ){
			    id = col_value;
			    id = id.toLowerCase();
			    rg.getCell(rows, j).setValue(id);
			}
			else if ( col_name === MAIL_COL_NAME ) {　　//メールアドレス
				newAdd = col_value;
			}
			else{
				body += "【"+col_name+"】\n";
				body += "   "+col_value + "\n";
				
				theRow = row.replace("{{col_name}}", col_name, "g");
				theRow = theRow.replace("{{col_value}}", col_value, "g");
				theRow = theRow.replace(theRow, theRow+"{{data}}", "g");
				
				htmlMail = htmlMail.replace("{{data}}", theRow, "g");
			}
		}	//end for (var j = 1; j <= cols; j++ )
		
		body = header + body + sep + footer;
		subject += shName;                  //題名にシート名を追加
		
		htmlMail = htmlMail.replace("{{data}}", "", "g");
		
		
		// id ---> oldAdd
		// member sheet
		
		var mssUrl  =   "https://docs.google.com/spreadsheets/d/" + mssID + "/edit";
		var mss     =   SpreadsheetApp.openByUrl(mssUrl);
		Logger.log(mss.getName());
		
		//if (mss){
		    
		//}else{
		    //var ui = SpreadsheetApp.getUi();
		    //ui.alert("No member sheet!!");
		//}
		
		var idCol   =   3;
		var addCol  =   7;
		var nameCol =   5;
		
		
		var ms = mss.getSheetByName(sem+"."+year);
		
		var mRows = ms.getLastRow();
		var mId;
		
		for (var i =2; i <= mRows; i++) {
		    mId = ms.getRange(i, idCol).getValue();
            
		    if (mId === id){
		        oldAdd = ms.getRange(i, addCol).getValue();
		        var thisRow = i;
		        i = mRows + 1;
		    }
		}
		//id--->oldAdd ends
		
		if (newAdd){
		    to = newAdd;
		    
		    if (newAdd!==oldAdd){
		    ms.getRange(thisRow, addCol).setValue(newAdd);
		    ms.getRange(thisRow, addCol).setBackground('red');
		    
		    }
		    
		}else if (oldAdd){
		    to = oldAdd;
		    
		}else{
		    MailApp.sendEmail(hostml, "【失敗】Googleフォームにメールアドレスが指定されていません", body);
		}
		

		// 送信先オプション
		var options = {};
		if ( cc )    options.cc      = cc;
		if ( bcc )   options.bcc     = bcc;
		if ( reply ) options.replyTo = reply;
		if (htmlMail)options.htmlBody=htmlMail;
		
		
		
		// send Email
		if ( to ) {
			MailApp.sendEmail(to, subject, body, options);
		}else{
			MailApp.sendEmail(hostml, "【失敗】Googleフォームにメールアドレスが指定されていません", body);
		}
	}catch(e){
		MailApp.sendEmail(hostml, "【失敗】Googleフォームからメール送信中にエラーが発生", e.message);
	}
}