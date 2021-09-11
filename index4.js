//以下的四列require裡的內容，請確認是否已經用npm裝進node.js
var linebot = require('linebot');
var express = require('express');
var google = require('googleapis');
var googleAuth = require('google-auth-library');

//以下的引號內請輸入申請LineBot取得的各項資料，逗號及引號都不能刪掉
var bot = linebot({
  channelId: '1567671197',
  channelSecret: '79fd196962ab5b9b839ad53bebfb44af',
  channelAccessToken: 'EI13t2LvkVbM8KA01RjFITLs2khkrYTcpO0BSi4Lidvnu7N02iB7p3Ut/b3rC1425+mEGOOkaEM5WSi7xVN4pDUMwjwrph6Rjzndq70hvvmIIYoGZUgsGBynNFNcynF6OARwN7KFyTPIqAUOFB1HIgdB04t89/1O/w1cDnyilFU='
});

//底下輸入client_secret.json檔案的內容
var myClientSecret={"installed":{"client_id":"481848113484-k2nahgno9n8rihpdc3sovan8pf64nlu3.apps.googleusercontent.com","project_id":"punchinlinebot-database","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://accounts.google.com/o/oauth2/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"Mrk9J8Qb10NuYM6w88hHw5hv","redirect_uris":["urn:ietf:wg:oauth:2.0:oob","http://localhost"]}};

var auth = new googleAuth();
var oauth2Client = new auth.OAuth2(myClientSecret.installed.client_id,myClientSecret.installed.client_secret, myClientSecret.installed.redirect_uris[0]);

//底下輸入sheetsapi.json檔案的內容
oauth2Client.credentials ={"access_token":"ya29.Glt7BWILhp3riyfn2bNJZXPLGu6woKepNmvXffdd8HBJfP2qz2yFBp7B-aiQ-UkxPirEAW1Ai9GiKtd5gReHMz8OI_fg-R3tvB0GASuuNO0gP3sF6Gx-H4IXVHOV","refresh_token":"1/Nw1g6Rg46RF_t9-0FjonjEJFmmbLXW005eo8oraSanU","token_type":"Bearer","expiry_date":1520774722776};

//試算表的ID，引號不能刪掉
var mySheetId='1U5SnTmYGlfgefnNVv5dm5oXSqZhGAqCCM3IFxqaxaOg';

var Menbers=[];
var users=[];
var users_name = [];
var Users_sheet_save = [];
var step = 0;
var msg ;
var MenbersCont = 0;
var locat = 'A1';
var Sheet_data = [];
var SheetsCont = 0;

getMenbers();
getSheets(getMonths(nowMonth(calcTime('+8'))));

function getSheets(Month){
	var sheets = google.sheets('v4');
	sheets.spreadsheets.values.get({
		auth: oauth2Client,
		spreadsheetId: mySheetId,
		range:encodeURI(Month),
	},function(err, response) {
		if (err) {
			console.log('讀取問題檔的API產生問題：' + err);
        return;}
	var rows = response.values;
    if (rows.length == 0) {
        console.log('No data found.');
    } else {
		Sheet_data=rows;
		SheetsCont=Sheet_data[0].length;
		console.log('要問的問題已下載完畢！');
    }
  });
}

//輸入月份回傳國字月份 ex.輸入1 回傳'1月'
function getMonths(Month){
	if( Month == 1){
		return '1月';
	}else if(Month == 2){
		return '2月';
	}else if(Month == 3){
		return '3月';
	}else if(Month == 4){
		return '4月';
	}else if(Month == 5){
		return '5月';
	}else if(Month == 6){
		return '6月';
	}else if(Month == 7){
		return '7月';
	}else if(Month == 8){
		return '8月';
	}else if(Month == 9){
		return '9月';
	}else if(Month == 10){
		return '10月';
	}else if(Month == 11){
		return '11月';
	}else if(Month == 12){
		return '12月';
	}
}

function getMenbers(){
	var sheets = google.sheets('v4');
	sheets.spreadsheets.values.get({
		auth: oauth2Client,
		spreadsheetId: mySheetId,
		range:encodeURI('人員名單'),
	},function(err, response) {
		if (err) {
			console.log('讀取問題檔的API產生問題：' + err);
        return;}
	var rows = response.values;
    if (rows.length == 0) {
        console.log('No data found.');
    } else {
		Menbers=rows;
		MenbersCont=Menbers[0].length;
		console.log('要問的問題已下載完畢！');
    }
  });
}

function MenberNum(Name){
	var i = 1;
	while(i<50){
		if(Menbers[i]!=''){
			i++
		}else{
			break;
		}
	}
	var j = 1;
	while(j<i){
		if(Name == Menbers[j]){
			return j+1;
		}else{
			j++;
		}
	}
}

function convert(num){
    return num <= 26 ? 
    String.fromCharCode(num + 64) : convert(~~((num - 1) / 26)) + convert(num % 26 || 26);
}

//這是將取得的資料儲存進試算表的函式
function appendMyRow(userId) {
   var request = {
      auth: oauth2Client,
      spreadsheetId: mySheetId,
      range:encodeURI('表單回應 1'),
      insertDataOption: 'INSERT_ROWS',
      valueInputOption: 'RAW',
      resource: {
        "values": [
          users[userId].replies
        ]
      }
   };
   var sheets = google.sheets('v4');
   sheets.spreadsheets.values.append(request, function(err, response) {
      if (err) {
         console.log('The API returned an error: ' + err);
         return;
      }
   });
}

function appendMyRowMember(userId) {
   var request = {
      auth: oauth2Client,
      spreadsheetId: mySheetId,
      range:encodeURI('人員名單'),
      insertDataOption: 'INSERT_ROWS',
      valueInputOption: 'RAW',
      resource: {
        "values": [
          users_name[userId].replies
        ]
      }
   };
   var sheets = google.sheets('v4');
   sheets.spreadsheets.values.append(request, function(err, response) {
      if (err) {
         console.log('The API returned an error: ' + err);
         return;
      }
   });
}

function appendMyRowMonths(Month,local,userId) {
   var request = {
      auth: oauth2Client,
      spreadsheetId: mySheetId,
      range:encodeURI( Month +"!"+ local),
      insertDataOption: 'OVERWRITE',
      valueInputOption: 'RAW',
      resource: {
        "values": [
          Users_sheet_save[userId].replies
        ]
      }
   };
   var sheets = google.sheets('v4');
   sheets.spreadsheets.values.append(request, function(err, response) {
      if (err) {
         console.log('The API returned an error: ' + err);
         return;
      }
   });
}

var punchinbutton={
	type: 'template',
	altText: '打卡功能表~',
	template: {
		type: 'buttons',
		thumbnailImageUrl: 'https://media3.giphy.com/media/Cmr1OMJ2FN0B2/giphy.gif',
		title: '哈囉~您好',
		text: '請選擇您要做的:',
		actions: [{
			type: 'postback',
			label: '我要上班',
			data: 'button_打卡'
		}, {
			type: 'postback',
			label: '我要下班',
			data: 'button_下班'
		}]
	}
};

var punchin_confirm_button = {
	type: 'template',
	altText: '你確定要上班打卡嗎?',
	template: {
	type: 'confirm',
	text: '你確定要上班打卡嗎?',
	actions: [{
		type: 'postback',
		label: '是',
		data:'confirm_打卡'
	}, {
		type: 'postback',
		label: '不要',
		data: '請再操作一次'
	}]
	}
};

var punchout_confirm_button = {
	type: 'template',
	altText: '你確定要下班打卡嗎?',
	template: {
		type: 'confirm',
		text: '你確定要下班打卡嗎?',
		actions: [{
			type: 'postback',
			label: '是',
			data: 'confirm_下班'
		}, {
			type: 'postback',
			label: '否',
			data: '請再操作一次'
		}]
	}
};

var take_a_day_off_confrim_button_special_break = {
	type: 'template',
	altText: '你確定要請特休嗎?',
	template: {
		type: 'confirm',
		text: '你確定要請特休嗎?',
		actions: [{
			type: 'postback',
			label: '恩',
			data:'button_特休'
		}, {
			type: 'postback',
			label: '按錯',
			data: '請再操作一次'
		}]
	}
};

var take_a_day_off_confrim_button_sick_leave= {
	type: 'template',
	altText: '你確定要請病假嗎?',
	template: {
		type: 'confirm',
		text: '你確定要請病假嗎?',
		actions: [{
			type: 'postback',
			label: '恩',
			data:'button_病假'
		}, {
			type: 'postback',
			label: '按錯',
			data: '請再操作一次'
		}]
	}
};

var take_a_day_off_confrim_button_casual_leave= {
	type: 'template',
	altText: '你確定要請事假嗎?',
	template: {
		type: 'confirm',
		text: '你確定要請事假嗎?',
		actions: [{
			type: 'postback',
			label: '恩',
			data:'button_事假'
		}, {
			type: 'postback',
			label: '按錯',
			data: '請再操作一次'
		}]
	}
};

var take_a_day_off_button = {
	type: 'template',
	altText: '請假事由~',
	template: {
		type: 'buttons',
		thumbnailImageUrl: 'https://i.ytimg.com/vi/7g_zGtH13LY/maxresdefault.jpg',
		title: '您好~請選擇你要請假的原因',
		text: '請選擇你要請假的原因:',
		actions: [{
			type: 'postback',
			label: '特休',
			data: 'confirm_特休'
		}, {
			type: 'postback',
			label: '事假',
			data: 'confirm_事假'
		},{
			type: 'postback',
			label: '病假',
			data: 'confirm_病假'
		}]
	}
};

//LineBot收到user的文字訊息時的處理函式
bot.on('message', function(event) {
	if (event.message.type === 'text') {
		msg = event.message.text;
		var myId=event.source.userId;
		if (users[myId]===undefined){
			users[myId]=[];
			users[myId].userId=myId;
			users[myId].step=0;
			users[myId].replies=[];
		}
		if (Users_sheet_save[myId]===undefined){
			Users_sheet_save[myId]=[];
			Users_sheet_save[myId].userId=myId;
			Users_sheet_save[myId].step=0;
			Users_sheet_save[myId].replies=[];
		}
		if (users_name[myId]===undefined){
			users_name[myId]=[];
			users_name[myId].userId=myId;
			users_name[myId].step=0;
			users_name[myId].replies=[];
		}
		if (msg == '我要打卡~'){
			getMenbers();
			//12月31清除11月前資料
			if(nowMonth(calcTime('+8')) == 12 & nowDay(calcTime('+8')) == 31){
				Users_sheet_save[myId].replies[0]='';
				var x = 1 ;
				while(x <= 11){
					var y = 2;
					while(y <= 26){
						var z = 3;
						while( z <= 33){
							appendMyRowMonths( x +'月', convert(y)+z , myId);
							z++;
						}
						y++;
					}
					x++;
				}
			}
			//1月31清除12月資料
			else if(nowMonth(calcTime('+8')) == 1 & nowDay(calcTime('+8')) == 31){
				Users_sheet_save[myId].replies[0]='';
				var y = 2;
				while(y <= 26){
					var z = 3;
					while( z <= 33){
						appendMyRowMonths( '1月', convert(y)+z , myId);
						z++;
					}
					y++;
				}
			}
			step = 0;
			event.reply(punchinbutton);
		}else if(msg == '我要請假~'){
			getMenbers();
			step = 0;
			event.reply(take_a_day_off_button);
		}else if( msg == '我要新增人員'){
			getMenbers();
			step = 6;
			event.reply('請輸入您的姓名~');
		}else if (step == 1){
			getMenbers();
			getSheets(getMonths(nowMonth(calcTime('+8'))));
			step = 0;
			if(Sheet_data[(parseInt(MenberNum(msg).toString())*2)-2][(parseInt(nowDay(calcTime('+8')))+2)] != ''){
				event.reply('你今天已經打過上班卡了~');
			}else{
				Users_sheet_save[myId].replies[0]=nowhour(calcTime('+8'));
				appendMyRowMonths(getMonths(nowMonth(calcTime('+8'))),convert((parseInt(MenberNum(msg).toString())*2)-2)+(parseInt(nowDay(calcTime('+8')))+2),myId);
				users[myId].replies[0]=calcTime('+8');
				users[myId].replies[1]=msg;
				users[myId].replies[2]='上班';
				appendMyRow(myId);
				event.reply('上班打卡成功!~');
			}
		}else if(step == 2){
			getMenbers();
			getSheets(getMonths(nowMonth(calcTime('+8'))));
			step = 0;
			if(Sheet_data[(parseInt(MenberNum(msg).toString())*2)-1][(parseInt(nowDay(calcTime('+8')))+2)] != ''){
				event.reply('你今天已經打過下班卡囉了~');
			}else{
				Users_sheet_save[myId].replies[0]=nowhour(calcTime('+8'));
				appendMyRowMonths(getMonths(nowMonth(calcTime('+8'))),convert((parseInt(MenberNum(msg).toString())*2)-1)+(parseInt(nowDay(calcTime('+8')))+2),myId);
				users[myId].replies[0]=calcTime('+8');
				users[myId].replies[1]=msg;
				users[myId].replies[2]='下班';
				appendMyRow(myId);
				event.reply('下班打卡成功!');
			}
		}else if(step == 3){
			getMenbers();
			getSheets(getMonths(nowMonth(calcTime('+8'))));
			step = 0;
			if(Sheet_data[(parseInt(MenberNum(msg).toString())*2)-2][(parseInt(nowDay(calcTime('+8')))+2)] != ''){
				event.reply('你今天已經打過卡囉~');
			}else{
				Users_sheet_save[myId].replies[0]='特休';
				appendMyRowMonths(getMonths(nowMonth(calcTime('+8'))),convert((parseInt(MenberNum(msg).toString())*2)-2)+(parseInt(nowDay(calcTime('+8')))+2),myId);
				appendMyRowMonths(getMonths(nowMonth(calcTime('+8'))),convert((parseInt(MenberNum(msg).toString())*2)-1)+(parseInt(nowDay(calcTime('+8')))+2),myId);
				users[myId].replies[0]=calcTime('+8');
				users[myId].replies[1]=msg;
				users[myId].replies[2]='特休';
				appendMyRow(myId);
				event.reply('特休請假成功!');
			}
		}else if(step == 4){
			getMenbers();
			getSheets(getMonths(nowMonth(calcTime('+8'))));
			step = 0;
			if(Sheet_data[(parseInt(MenberNum(msg).toString())*2)-2][(parseInt(nowDay(calcTime('+8')))+2)] != ''){
				event.reply('你今天已經打過卡囉~');
			}else{
				Users_sheet_save[myId].replies[0]='事假';
				appendMyRowMonths(getMonths(nowMonth(calcTime('+8'))),convert((parseInt(MenberNum(msg).toString())*2)-2)+(parseInt(nowDay(calcTime('+8')))+2),myId);
				appendMyRowMonths(getMonths(nowMonth(calcTime('+8'))),convert((parseInt(MenberNum(msg).toString())*2)-1)+(parseInt(nowDay(calcTime('+8')))+2),myId);
				users[myId].replies[0]=calcTime('+8');
				users[myId].replies[1]=msg;
				users[myId].replies[2]='事假';
				appendMyRow(myId);
				event.reply('事假請假成功!');
			}
		}else if(step == 5){
			getMenbers();
			getSheets(getMonths(nowMonth(calcTime('+8'))));
			step = 0;
			if(Sheet_data[(parseInt(MenberNum(msg).toString())*2)-2][(parseInt(nowDay(calcTime('+8')))+2)] != ''){
				event.reply('你今天已經打過卡囉~');
			}else{
				Users_sheet_save[myId].replies[0]='病假';
				appendMyRowMonths(getMonths(nowMonth(calcTime('+8'))),convert((parseInt(MenberNum(msg).toString())*2)-2)+(parseInt(nowDay(calcTime('+8')))+2),myId);
				appendMyRowMonths(getMonths(nowMonth(calcTime('+8'))),convert((parseInt(MenberNum(msg).toString())*2)-1)+(parseInt(nowDay(calcTime('+8')))+2),myId);
				users[myId].replies[0]=calcTime('+8');
				users[myId].replies[1]=msg;
				users[myId].replies[2]='病假';
				appendMyRow(myId);
				event.reply('病假請假成功!');
			}
		}else if(step == 6){
			getMenbers();
			step = 0;
			users_name[myId].replies[0]=msg;
			appendMyRowMember(myId);
			event.reply('新增'+msg+'成功!');
		}else{
			event.reply('請再操作一次'+Sheet_data[20][1]);
		}
	}
});

bot.on('postback',function (event){
	var myResult = event.postback.data;
	if (myResult == 'button_打卡'){
		event.reply(punchin_confirm_button);
	}else if( myResult == 'button_下班'){
		event.reply(punchout_confirm_button);
	}else if( myResult == 'confirm_打卡'){
		event.reply("請輸入您的姓名~");
		step = 1;
	}else if( myResult == 'confirm_下班'){
		event.reply("請輸入您的姓名~");
		step = 2;
	}else if( myResult == 'confirm_特休'){
		event.reply(take_a_day_off_confrim_button_special_break);
	}else if( myResult == 'confirm_事假'){
		event.reply(take_a_day_off_confrim_button_casual_leave);
	}else if( myResult == 'confirm_病假'){
		event.reply(take_a_day_off_confrim_button_sick_leave);
	}else if( myResult == 'button_特休'){
		event.reply("請輸入您的姓名~");
		step = 3;
	}else if( myResult == 'button_事假'){
		event.reply("請輸入您的姓名~");
		step = 4;
	}else if( myResult == 'button_病假'){
		event.reply("請輸入您的姓名~");
		step = 5;
	}else{
		event.reply('請再操作一次');
	}
});

//回傳當地時間 ex.台灣 = offset = '+8'
function calcTime(offset) {

    // create Date object for current location
    d = new Date();
    // convert to msec
    // add local time zone offset 
    // get UTC time in msec
    utc = d.getTime() + (d.getTimezoneOffset() * 60000);
    
    // create new Date object for different city
    // using supplied offset
    nd = new Date(utc + (3600000*offset));
    
    // return time as a string
    return nd.toLocaleString();
}

//回傳當地小時
function nowhour(calcT){
	var time = [];
	time = calcT.split(' ');
	return time[1];
}

//回傳當地月份
function nowMonth(calcT){
	var time = [];
	time = calcT.split(' ');
	var month = [];
	month = time[0].split('-');
	return month[1];
}

//回傳當地日期
function nowDay(calcT){
	var time = [];
	time = calcT.split(' ');
	var month = [];
	month = time[0].split('-');
	return month[2];
}

const app = express();
const linebotParser = bot.parser();

app.get('/',function(req,res){
    res.send('Hello World!');
});

app.post('/linewebhook', linebotParser);

//因為 express 預設走 port 3000，而 heroku 上預設卻不是，要透過下列程式轉換
var server = app.listen(process.env.PORT || 80, function() {
  var port = server.address().port;
  console.log("App now running on port", port);
});