(function(){
	var time=0,
		dateStart = new Date().getTime(),
		dateCurrent = dateStart,
		result = 0,
		shadow_top = document.getElementById("currentTime_shadow_top"),
		output = document.getElementById("currentTime"),
		shadow = document.getElementById("currentTime_shadow");
		
	function tickTak() {
		clearTimeout(time);
		dateCurrent = parseInt(((new Date().getTime()) - dateStart) / 1000);
		var h = dateCurrent / 3600 ^ 0,
			m = (dateCurrent - h * 3600) / 60 ^ 0,
			s = dateCurrent - h * 3600 - m * 60,
			r = (h < 10 ? "0" + h : h) + ":" + (m < 10 ? "0" + m : m) + ":" + (s < 10 ? "0" + s : s);
		output.innerText = shadow.innerText = shadow_top.innerText = r;
		setTimeout(tickTak, 10);
	}
	tickTak()
}())