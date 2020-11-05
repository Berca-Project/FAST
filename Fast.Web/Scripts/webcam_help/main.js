/*
*  derived from examples https://developers.google.com/web/updates/2016/03/access-usb-devices-on-the-web  
*  https://webrtc.github.io/samples/src/content/getusermedia/resolution/
*  and
*  https://webrtchacks.com/getusermedia-resolutions-3/
*/

'use strict';

var videoElement = document.querySelector('video');
var videoSelect = document.querySelector('select#videoSource');
var selectors = [videoSelect];

function gotDevices(deviceInfos) {
  // Handles being called several times to update labels. Preserve values.
  var values = selectors.map(function(select) {
    return select.value;
  });
  selectors.forEach(function(select) {
    while (select.firstChild) {
      select.removeChild(select.firstChild);
    }
  });
  for (var i = 0; i !== deviceInfos.length; ++i) {
    var deviceInfo = deviceInfos[i];
    var option = document.createElement('option');
    option.value = deviceInfo.deviceId;
    if (deviceInfo.kind === 'videoinput') {
	option.text = deviceInfo.label || 'camera ' + (videoSelect.length + 1);
	console.log('camera type source/device: ', deviceInfo);
      
	videoSelect.appendChild(option);
    } else {
	//console.log('Some other kind of source/device: ', deviceInfo);
    }
  }
  selectors.forEach(function(select, selectorIndex) {
    if (Array.prototype.slice.call(select.childNodes).some(function(n) {
      return n.value === values[selectorIndex];
    })) {
      select.value = values[selectorIndex];
    }
  });
}

navigator.mediaDevices.enumerateDevices().then(gotDevices).catch(handleError);



function gotStream(stream) {
  window.stream = stream; // make stream available to console
  videoElement.srcObject = stream;
  // Refresh button list in case labels have become available
  return navigator.mediaDevices.enumerateDevices();
}

function start1920() {
    if (window.stream) {
	window.stream.getTracks().forEach(function(track) {
	    track.stop();
	});
    }
    var videoSource = videoSelect.value;
    var constraints = {
	//video: {deviceId: videoSource ? {exact: videoSource} : undefined}
	//video: {deviceId: videoSource ? {exact: videoSource} : undefined}, width:{ min: 1280},  height:{ min: 720  }  
	video: {deviceId: videoSource ? {exact: videoSource} : undefined,   width: {exact: 1280 }, height:{exact: 720}  }
    };
    
    var constraints1920 = {
	//video: {deviceId: videoSource ? {exact: videoSource} : undefined,   width: {exact: 1920 }, height:{exact: 1080}  }
	video: {deviceId: videoSource ? {exact: videoSource} : undefined,   width: {min: 1920 }, height:{min: 1080}  }
	//video: {deviceId: videoSource ? {exact: videoSource} : undefined,   minWidth:1920, minHeight:1080 }
    };
    
    var constraintsVGA = { video: { mandatory: { maxWidth: 320, maxHeight: 320} } };
    var constraintsHD = { video: { mandatory: {  minWidth: 1280,  minHeight: 720  }  }};
    

    navigator.mediaDevices.getUserMedia(constraints1920).
	then(gotStream).then(gotDevices).catch(handleError1920);


}

function start2592() {
    console.log('trying 2592 video master.: ');

    
    if (window.stream) {
	window.stream.getTracks().forEach(function(track) {
	    track.stop();
	});
    }
    var videoSource = videoSelect.value;
    
//    var constraints1920 = {
//	video: {deviceId: videoSource ? {exact: videoSource} : undefined,   width: {min: 1920 }, height:{min: 1080}  }
//    };

    var constraints2592 = {
	video: {deviceId: videoSource ? {exact: videoSource} : undefined,   width: {min: 2592 }, height:{min: 1944}  }
 	//video: {deviceId: videoSource ? {exact: videoSource} : undefined,   width: {min: 2304 }, height:{min: 1536}  }
       };
    
 

    navigator.mediaDevices.getUserMedia(constraints2592).
	then(gotStream).then(gotDevices).catch(handleError2592);


}

function start1280() {
    if (window.stream) {
	window.stream.getTracks().forEach(function(track) {
	    track.stop();
	});
    }
    var videoSource = videoSelect.value;
    var constraints = {
	//video: {deviceId: videoSource ? {exact: videoSource} : undefined}
	//video: {deviceId: videoSource ? {exact: videoSource} : undefined}, width:{ min: 1280},  height:{ min: 720  }  
	video: {deviceId: videoSource ? {exact: videoSource} : undefined,   width: 160, height:120  } // change for capture resolution
    };
    
    navigator.mediaDevices.getUserMedia(constraints).
	then(gotStream).then(gotDevices).catch(handleError);
}

//videoSelect.onchange = start1920;
//start1920();
videoSelect.onchange = start2592;
start2592();

function handleError(error) {
  console.log('navigator.getUserMedia error: ', error);
}

function handleError1920(error) {
    console.log('camera cannot do 1920 HD or safari issue error: ', error);
    start1280();

}
function handleError2592(error) {
    console.log('no 2592 camera mode or safari issue error: ', error);
    start1920();

}
