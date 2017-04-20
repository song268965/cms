/* *
* 给定一个剩余时间（s）动态显示一个剩余时间.
* 当大于一天时。只显示还剩几天。小于一天时显示剩余多少小时，多少分钟，多少秒。秒数每秒减1 *
*/

// 初始化变量
var auctionDate = 0;
var _GMTEndTime = 0;
var showTime = "leftTime";
var _day = '天';
var _hour = '时';
var _minute = '分';
var _second = '秒';
var _end = '已结束';

var cur_date = new Date();
var startTime = cur_date.getTime();
var Temp;
var timerID = null;
var timerRunning = false;

/* 得到日期年月日等加数字后的日期 */
Date.prototype.dateAdd = function(interval,number)
{
var d = this;
var k={"y":"FullYear", "q":"Month", "m":"Month", "w":"Date", "d":"Date", "h":"Hours", "n":"Minutes", "s":"Seconds", "ms":"MilliSeconds"};
var n={"q":3, "w":7};
eval("d.set"+k[interval]+"(d.get"+k[interval]+"()+"+((n[interval]||1)*number)+")");
return d;
};
/* 计算两日期相差的日期年月日等 */
Date.prototype.dateDiff = function(interval,objDate)
{
var d=this, t=d.getTime(), t2=objDate.getTime(), i={};
i["y"]=objDate.getFullYear()-d.getFullYear();
i["q"]=i["y"]*4+Math.floor(objDate.getMonth()/4)-Math.floor(d.getMonth()/4);
i["m"]=i["y"]*12+objDate.getMonth()-d.getMonth();
i["ms"]=objDate.getTime()-d.getTime();
i["w"]=Math.floor((t2+345600000)/(604800000))-Math.floor((t+345600000)/(604800000));
i["d"]=Math.floor(t2/86400000)-Math.floor(t/86400000);
i["h"]=Math.floor(t2/3600000)-Math.floor(t/3600000);
i["n"]=Math.floor(t2/60000)-Math.floor(t/60000);
i["s"]=Math.floor(t2/1000)-Math.floor(t/1000);
return i[interval];
};
/*倒计时,d日期格式为： 2012/1/1 10:10:00*/
function showtimes(showTimeId, d){
var d1 = new Date(d);
var d2 = new Date();
var ds=d2.dateDiff("s" ,d1);
  showtime(showTimeId, ds);
}

/*倒计时,seconds两个时间的秒差*/
function showtime(showTimeId, seconds) {
    now = new Date();
    var ts = parseInt((startTime - now.getTime()) / 1000) + seconds;
    var dateLeft = 0;
    var hourLeft = 0;
    var minuteLeft = 0;
    var secondLeft = 0;
    var hourZero = '';
    var minuteZero = '';
    var secondZero = '';
    if (ts < 0) {
        ts = 0;
        CurHour = 0;
        CurMinute = 0;
        CurSecond = 0;
    }
    else {
        dateLeft = parseInt(ts / 86400);
        ts = ts - dateLeft * 86400;
        hourLeft = parseInt(ts / 3600);
        ts = ts - hourLeft * 3600;
        minuteLeft = parseInt(ts / 60);
        secondLeft = ts - minuteLeft * 60;
    }

    if (hourLeft < 10) {
        hourZero = '0';
    }
    if (minuteLeft < 10) {
        minuteZero = '0';
    }
    if (secondLeft < 10) {
        secondZero = '0';
    }
    if (dateLeft > 0) {
        Temp = "<li><span>" + dateLeft + "</span>" + _day + "</li><li><span>" + hourZero + hourLeft + "</span>" + _hour + "</li><li><span>" + minuteZero + minuteLeft + "</span>" + _minute + "</li><li><span>" + secondZero + secondLeft + "</span>" + _second + "</li>";
    }
    else {
        if (hourLeft > 0) {
            Temp = "<li><span>" + hourLeft + "</span>" + _hour + "</li><li><span>" + minuteZero + minuteLeft + "</span>" + _minute + "</li><li><span>" + secondZero + secondLeft + "</span>" + _second + "</li>";
        }
        else {
            if (minuteLeft > 0) {
                Temp = "<li><span>" + minuteLeft + "</span>" + _minute + "</li><li><span>" + secondZero + secondLeft + "</span>" + _second + "</li>";
            }
            else {
                if (secondLeft > 0) {
                    Temp = "<li><span>" + secondLeft + "</span>" + _second + "</li>";
                }
                else {
                    Temp = '';
                }
            }
        }
    }

    if (seconds <= 0 || Temp == '') {
        Temp = "<strong>" + _end + "</strong>";
        stopclock();
    }
    if (document.getElementById(showTimeId)) {
        $("#"+showTimeId).html(Temp);
        //document.getElementById(showTimeId).innerHTML = ;
    }

    timerID = setTimeout("showtime(\"" + showTimeId + "\"," + seconds + ")", 1000);
    timerRunning = true;
}

var timerID = null;
var timerRunning = false;
function stopclock() {
    if (timerRunning) {
        clearTimeout(timerID);
    }
    timerRunning = false;
}

function macauclock(showTimeId, seconds) {
    //stopclock();
    showtime(showTimeId, seconds);
}

function onload_leftTime(showTimeId, seconds) {
    /* 第一次运行时初始化语言项目 */
    try {
        _GMTEndTime = gmt_end_time;
        // 剩余时间
        _day = day;
        _hour = hour;
        _minute = minute;
        _second = second;
        _end = end;
    }
    catch (e) {
    }

    if (_GMTEndTime > 0) {
        var tmp_val = parseInt(_GMTEndTime) - parseInt(cur_date.getTime() / 1000 + cur_date.getTimezoneOffset() * 60);
        if (tmp_val > 0) {
            auctionDate = tmp_val;
        }
    }

    macauclock(showTimeId, seconds);
    try {
        initprovcity();
    }
    catch (e) {
    }
}

function loadleftTime(showTimeId, seconds) {
    //auctionDate = seconds;
    onload_leftTime(showTimeId, seconds);
}
