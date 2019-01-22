xmlhttp=new XMLHttpRequest()	
var bkn = "" 
// js repalaceAll方法
String.prototype.replaceAll  = function(s1,s2){     
    return this.replace(new RegExp(s1,"gm"),s2);     
}
// 实现sleep方法
function wait(ms) {
    var start = Date.now(),
        now = start;
    while (now - start < ms) {
      now = Date.now();
    }
}
//时间戳格式化方法 
function getLocalTime(timestamp) {
    var d = new Date(timestamp);
    var date = (d.getFullYear()) + "-" +
            (d.getMonth() + 1) + "-" +
            (d.getDate()) + " " +
            (d.getHours()) + ":" +
            (d.getMinutes()) + ":" +
            (d.getSeconds());
    return date;
}
function dataToExcel(data,file_name){
	/* 需要导出的JSON数据 */
	// var data = [
	// 	{"name":"John", "city": "Seattle"},
	// 	{"name":"Mike", "city": "Los Angeles"},
	// 	{"name":"Zach", "city": "New York"}
	// ];

	/* 如果没有导入xlsx组件则导入 */
	if(typeof XLSX == 'undefined') XLSX = require('xlsx');

	/* 创建worksheet */
	var ws = XLSX.utils.json_to_sheet(data);

	/* 新建空workbook，然后加入worksheet */
	var wb = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb, ws, "People");

	/* 生成xlsx文件 */
	XLSX.writeFile(wb, file_name+".xlsx");
}
// 获取所有QQ好友并保存
download_qq =function () {
	// 发送请求
	xmlhttp.open("POST","https://qun.qq.com/cgi-bin/qun_mgr/get_friend_list",false);
	xmlhttp.send("bkn="+bkn);
	var json = JSON.parse(xmlhttp.responseText)
	if (json.ec ==0){
		var content = Array();
		for (var i= 1 in json.result){
			var men = json.result[i]
			var gname = men.gname.replaceAll("&nbsp;"," ")
			var mems = men.mems
			for (var j=0;j < mems.length;j++)
			{
				content.push({"qq": mems[j].uin,"备注名":mems[j].name.replaceAll("&nbsp;"," "),"分组名":gname})
			}
			
		}	
		dataToExcel(content, "all_qq_friends");
	}
	else{
		alert('获取失败，详情请查看console')
		console.log( xmlhttp.responseText);
	}
 }
 // 获取指定群的好友信息系并保存
 download_group =function (qun_id,qun_name) {
	// 发送请求
	var start_num = 0	
	var end_num = -1
	var page_num = 20
	var all_num = 0
	var data = Array();
	while (all_num > end_num)	{
		// 设置等待时长，不然当天要封号
		start_num = end_num + 1
		end_num = start_num + page_num
		xmlhttp=new XMLHttpRequest()	
		xmlhttp.open("POST","https://qun.qq.com/cgi-bin/qun_mgr/search_group_members",false);
		// xmlhttp.setRequestHeader("referer", "https://qun.qq.com/member.html")
		// xmlhttp.send("bkn="+bkn+"&gc="+qun_id + "&st="+ start_num+"&end=" +end_num+"&sort=0");
		xmlhttp.send("gc="+qun_id + "&st="+ start_num+"&end=" +end_num+"&sort=0&"+"bkn="+bkn);
		var json = JSON.parse(xmlhttp.responseText);
		if (json.ec ==0){
			all_num = json.count
			var mems_list = json.mems
			console.log(mems_list);	
			for (var j=0;j <json.mems.length;j++)
			{
				var role 
				if (mems_list[j].role==0){
					role = "群主"
				}
				else if(mems_list[j].role==1){
					role = "管理员"
				}
				else{
					role = "成员"
				}
				if (mems_list[j].join_time == 0 ){
					join_time = "2012年5月以前"
				}
				else{
					join_time = getLocalTime(mems_list[j].join_time*1000)
				}
				if(mems_list[j].last_speak_time==0){
					last_speak_time =  "-"
				}
				else{
					last_speak_time = getLocalTime(mems_list[j].last_speak_time*1000)
				}
				data.push({"qq":mems_list[j].uin,"网名":mems_list[j].nick,"群名":qun_name,"备注":mems_list[j].card,"入群时间":join_time,"上次发言时间":last_speak_time,"q龄":mems_list[j].qage,"权限":role,"活跃度":mems_list[j].lv.point})
			}
			wait(200)
			console.log("wait 200ms")
		}
		else{
			console.log( json);
			alert('获取失败，详情请查看console');
			return false;
		}
	}
	dataToExcel(data, qun_id);
 }

// 获取单个cookie
$(function(){
	chrome.cookies.get( { url: 'https://qun.qq.com/member.html', name: 'skey' }, function( cookie ){
		if (cookie != null){
			for (var e =cookie.value, t = 5381, n = 0, o = e.length; o > n; ++n) 
				t += (t << 5) + e.charAt(n).charCodeAt();	
			bkn = 2147483647 & t
			console.log( "bkn="+bkn);		
			// $.post('https://qun.qq.com/cgi-bin/qun_mgr/get_group_list',{async:false,bkn:(2147483647 & t),function (result) {
			// 	console.log( result);		
			// }});
			// 发送请求
			xmlhttp.open("POST","https://qun.qq.com/cgi-bin/qun_mgr/get_group_list",false);
			xmlhttp.send("bkn="+bkn);
			var json = JSON.parse(xmlhttp.responseText)
			// console.log(json.ec);
			if (json.ec == 0){
				var key_list = Array('create','join','manage')
				var group_list = Array()
				for(var key in key_list){
					for(var group in json[key_list[key]]){
						group_list.push(json[key_list[key]][group])
					}
				}
				console.log( group_list);
				$("#td1").text("登陆成功");
				$("#i1").removeAttr("disabled");
				$("#i2").attr("disabled",false);
				$(function () {
					bindSelect();
					$('#info').text($('#qunSelect').val());
				  });
				//将数据集绑定select，重新选群后显示选中群名称
				bindSelect = function () {
					var $qunSelect = $('#qunSelect');
					if (group_list.length > 0) {
						for (var i = 0; i < group_list.length; i++) {
							var item = group_list[i];
							if (i == 0) {
								$qunSelect.append('<option value="' + item.gc + '" selected>' + item.gn + '</option>');
							} 
							else {
								$qunSelect.append('<option value="' + item.gc + '">' + item.gn + '</option>');
							}
						}
					}
				}
			}
			else{
				$("#td1").text("登陆已过期");
				$("#i1").attr("disabled","true");
			 	$("#i2").attr("disabled",true);
			}
		}
		else{
			$("#td1").text("未登录"); 
			$("#i1").attr("disabled","true");
			$("#i2").attr("disabled",true);
		}	
	});
 });

 // 导出按钮点击事件
 $('#i1').click(() => {
	$("#i1").attr("disabled",true);
	$("#i2").attr("disabled",true);
	download_qq();
	$("#i1").attr("disabled",false);
	$("#i2").attr("disabled",false);
});
$('#i2').click(() => {
	var options=$("#qunSelect option:selected");
	var qun_id = options.val()
	var qun_name = options.text()
	console.log(qun_id,qun_name)
	$("#i1").attr("disabled",true);
	$("#i2").attr("disabled",true);
	download_group(qun_id,qun_name);
	$("#i1").attr("disabled",false);
	$("#i2").attr("disabled",false);
});